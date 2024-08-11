import base64
import json
from io import BytesIO

import pytesseract
import streamlit as st
from openai import OpenAI
from openpyxl import Workbook
from pdf2image import convert_from_bytes

IMG2MD_SYSTEM_PROMPT = '''
You are an expert document parser.
Given the image of a document page, you transcribe all text of the page into Markdown format accurately and clearly.

# Instructions 
- Your transcription should include all the text in the image.
- You should reference the text output from OCR to achieve higher accuracy.
- Your output should be formatted clearly, maintaining the original content structure shown in the image.

# Output format
- Transcribe the document word-for-word whenever possible. Do not paraphrase any content nor add clarifying statements.
- Format all tables as Markdown tables.
- Do not split a single integrated table into multiple tables, unless they are different tables.
- Do not split content into multiple lines unless ABSOLUTELY NECESSARY.
- Output only the transcribed content. Do not reply with other content (e.g., here is the content of the document).
- Do not wrap any content in code blocks.
'''

IMG2MD_USER_PROMPT = '''
Here is the text output from OCR for reference:
{ocr}
---
Please parse the image content and transcribe it into a Markdown format, while maintaining the original structure. 
'''

FORMAT_TABLES_SYSTEM_PROMPT = '''
You are an expert in formatting Markdown tables.

# Instructions
- You will be given an image of a document page containing zero or more tables.
- You will also be provided with a raw Markdown text transcription of the document page.
- Your goal is to properly format all the tables in a Markdown format.
- If there are no tables in the provided content, return NO_TABLE_PRESENT and terminate.

# Output format
- A single table should not be split into multiple tables, unless they are separate tables.
- Separate tables should have a clear space between them, typically more than space between two rows within same table.
- Separate tables should have their own headings, titles, or descriptions.
- Each Markdown table's row should have the same number of columns (i.e., max of all rows from that table).
- If a cell spans multiple columns, represent it as a single cell with the appropriate number of dashes.
- At the end of each table, add a new line with #####END-OF-TABLE#####.
'''

FORMAT_TABLES_USER_PROMPT = '''
Here is the raw Markdown text transcription of the document page:
{transcription}
---
Please reference the image and transcription provided.
Please format all the tables into a well-structured Markdown table format.
'''

MD2JSON_SYSTEM_PROMPT = '''
You are an expert in converting Markdown tables into well-formatted JSON objects.

# Instructions
- If there are Markdown tables present in the content, focus on the Markdown tables only and ignore the rest.
- For each table, extract its title.
- For each table, convert the Markdown table's body into a JSON list of rows.
- Each row should be a list of cells, each representing the value of a table cell.
- If a cell contains a number, make that cell an integer or float in the JSON object.
'''

MD2JSON_USER_PROMPT = '''
For each table present, please return a JSON object that contains the title and body (a list of rows).
Each row should be represented by a list of cells.
---
The JSON response should follow this format:
{
  "tables": [
    {
      "title": <table-title>,
      "body": [
        [<cell-value>, <cell-value>, ..., <cell-value>],
        ...
        [<cell-value>, <cell-value>, ..., <cell-value>]
      ]
    },
    ...
  ]
}
'''


def encode_image(image):
    bfr = BytesIO()
    image.save(bfr, format='png')
    return base64.b64encode(bfr.getvalue()).decode('utf-8')


def extract_tables_from_image(client, image, ocr_string):
    messages = [
        {'role': 'system', 'content': IMG2MD_SYSTEM_PROMPT},
        {'role': 'user', 'content': [
            {'type': 'text', 'text': IMG2MD_USER_PROMPT.format(ocr=ocr_string)},
            {'type': 'image_url', 'image_url': {'url': f'data:image/png;base64,{encode_image(image)}'}}
        ]}
    ]
    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=messages
    )
    return completion.choices[0].message.content


def format_markdown_tables(client, markdown_content, image):
    user_prompt = FORMAT_TABLES_USER_PROMPT.format(transcription=markdown_content)
    messages = [
        {'role': 'system', 'content': FORMAT_TABLES_SYSTEM_PROMPT},
        {'role': 'user', 'content': [
            {'type': 'text', 'text': user_prompt},
            {'type': 'image_url', 'image_url': {'url': f'data:image/png;base64,{encode_image(image)}'}}
        ]}
    ]
    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=messages
    )
    return completion.choices[0].message.content


def convert_markdown_to_json(client, markdown_content):
    user_prompt = 'Here is the markdown content:\n' + markdown_content + '\n---\n' + MD2JSON_USER_PROMPT
    messages = [
        {'role': 'system', 'content': MD2JSON_SYSTEM_PROMPT},
        {'role': 'user', 'content': user_prompt}
    ]
    completion = client.chat.completions.create(
        model="gpt-4o",
        response_format={"type": "json_object"},
        messages=messages
    )
    return json.loads(completion.choices[0].message.content)


def create_excel_binary_from_json(json_data):
    wb = Workbook()
    ws = wb.active

    # Add tables to Excel
    for table in json_data:
        ws.append([table['title']])
        ws.append([])
        for row in table['body']:
            ws.append(row)
        ws.append([])
        ws.append([])

    # Save workbook to binary stream
    binary_stream = BytesIO()
    wb.save(binary_stream)
    binary_content = binary_stream.getvalue()

    return binary_content


def process_uploaded_pdf(client):
    uploaded_file = st.session_state['uploaded_pdf']

    if uploaded_file is not None:
        with st.status('Scanning uploaded PDF pages ...', expanded=True) as status:
            # Convert PDF pages into images
            images = convert_from_bytes(uploaded_file.getvalue(), fmt='png', thread_count=8)
            images = images[:min(st.secrets['PDF_PAGE_LIMIT'], len(images))]  # limit pages for cost reasons
            st.session_state['images'] = images

            # Extract table titles from images
            text_extracts = []
            formatted_tables = []
            table_extracts = []
            for i, image in enumerate(images):
                status.write(f'Reading tables from page {i + 1} ...')

                # Convert to string using OCR
                ocr_string = pytesseract.image_to_string(image)

                # Use OpenAI to extract Markdown text content from image
                text_from_image = extract_tables_from_image(client, image, ocr_string)
                text_extracts.append(text_from_image)

                # Format Markdown tables properly
                formatted_table = format_markdown_tables(client, text_from_image, image)
                formatted_tables.append(formatted_table)

                # Use OpenAI to convert Markdown to JSON
                if 'NO_TABLE_PRESENT' not in formatted_table:
                    table_json = convert_markdown_to_json(client, formatted_table)
                    table_extracts.extend(table_json['tables'])

            # Update status and session states
            status.update(label='✅ PDF processing complete', expanded=False)
            st.session_state['text_extracts'] = text_extracts
            st.session_state['formatted_tables'] = formatted_tables
            st.session_state['table_extracts'] = table_extracts


if __name__ == '__main__':
    st.title('Convert PDF Tables to Excel')

    # Initialize OpenAI Client
    openai_client = OpenAI(api_key=st.secrets['OPENAI_API_KEY'])

    # Upload scanned PDF for processing
    uploaded_pdf = st.file_uploader(
        f"Upload a PDF (of at most {st.secrets['PDF_PAGE_LIMIT']} pages) to convert to Excel",
        type='pdf',
        on_change=process_uploaded_pdf,
        args=(openai_client,),
        key='uploaded_pdf'
    )

    # Show tables
    if 'table_extracts' in st.session_state:
        st.divider()
        excel_workbook = create_excel_binary_from_json(st.session_state['table_extracts'])
        st.download_button(
            label='Download Excel Output',
            data=excel_workbook,
            file_name='Converted_Tables.xlsx',
            mime='application/vnd.ms-excel'
        )
