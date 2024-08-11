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
- Remove the page number and any other content irrelevant to the tables.
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
- For each table row, retain the row header if present. Remove **bold** or any other formatting.
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
    ws_original = wb.active

    # Add tables to Excel by page
    for page in sorted(json_data.keys()):
        ws = wb.create_sheet(f'Page {page}')
        for table in json_data[page]:
            ws.append([table['title']])
            ws.append([])
            for row in table['body']:
                ws.append(row)
            ws.append([])
            ws.append([])

    # Save workbook to binary stream
    wb.remove(ws_original)
    binary_stream = BytesIO()
    wb.save(binary_stream)
    binary_content = binary_stream.getvalue()

    return binary_content


def process_uploaded_pdf(client):
    uploaded_file = st.session_state['uploaded_pdf']

    if uploaded_file is not None:
        st.session_state['file_name'] = uploaded_file.name

        with st.status('Scanning uploaded PDF pages ...', expanded=True) as status:
            # Convert PDF pages into images
            images = convert_from_bytes(uploaded_file.getvalue(), fmt='png', thread_count=8)
            images = images[:min(st.secrets['PDF_PAGE_LIMIT'], len(images))]  # limit pages for cost reasons
            st.session_state['images'] = images

            # Extract table titles from images
            text_extracts = []
            formatted_tables = []
            table_extracts = {}
            for i, image in enumerate(images):
                # Create an empty holder to show page level details
                status.update(label=f'Processing page {i + 1} ...', expanded=True)
                placeholder = st.empty()

                # Convert to string using OCR
                placeholder.write('Scanning page using OCR ...')
                ocr_string = pytesseract.image_to_string(image)

                # Use OpenAI to extract Markdown text content from image
                placeholder.write('Transcribing content using AI vision ...')
                text_from_image = extract_tables_from_image(client, image, ocr_string)
                text_extracts.append(text_from_image)

                # Format Markdown tables properly
                placeholder.write('Extracting tables using AI ...')
                formatted_table = format_markdown_tables(client, text_from_image, image)
                formatted_tables.append(formatted_table)

                # Use OpenAI to convert Markdown to JSON
                if 'NO_TABLE_PRESENT' not in formatted_table:
                    placeholder.write('Converting tables into Excel ...')
                    table_json = convert_markdown_to_json(client, formatted_table)
                    table_extracts[i+1] = table_json['tables']

                # Empty placeholder at the end of processing each page
                placeholder.empty()

            # Update status and session states
            status.update(label='✅ PDF processing complete', expanded=False)
            st.session_state['text_extracts'] = text_extracts
            st.session_state['formatted_tables'] = formatted_tables
            st.session_state['table_extracts'] = table_extracts
    else:
        st.session_state.pop('file_name', None)
        st.session_state.pop('text_extracts', None)
        st.session_state.pop('formatted_tables', None)
        st.session_state.pop('table_extracts', None)


if __name__ == '__main__':
    st.title('PDF ➡️ Excel')

    # Initialize OpenAI Client
    openai_client = OpenAI(api_key=st.secrets['OPENAI_API_KEY'])

    # Upload scanned PDF for processing
    uploaded_pdf = st.file_uploader(
        f"Upload a PDF (of at most {st.secrets['PDF_PAGE_LIMIT']} pages) and convert its tables into an Excel file",
        type='pdf',
        on_change=process_uploaded_pdf,
        args=(openai_client,),
        key='uploaded_pdf'
    )

    # Show tables
    if 'table_extracts' in st.session_state:
        st.divider()

        if len(st.session_state['table_extracts']) > 0:
            # Set file names
            raw_file_name = st.session_state['file_name']
            excel_file_name = raw_file_name[:raw_file_name.rfind('.')] + '.xlsx'
            markdown_file_name = raw_file_name[:raw_file_name.rfind('.')] + '.md'

            # Build Excel and Markdown outputs
            excel_workbook = create_excel_binary_from_json(st.session_state['table_extracts'])

            # Articulate how users can use the product
            download_description = """
            In the Excel file below, you will find extracted tables from the uploaded PDF.
            Optionally, you may also download the transcription in Markdown format too view AI's interim steps.
            """
            st.write(download_description)

            # Add download buttons
            st.download_button(
                label='Download Excel',
                data=excel_workbook,
                file_name=excel_file_name,
                mime='application/vnd.ms-excel',
                type='primary'
            )

            st.download_button(
                label='Download Transcription',
                data='\n\n'.join(st.session_state['text_extracts']).encode('utf-8'),
                file_name=markdown_file_name,
                mime='text/markdown'
            )
        else:
            st.write('No tables are found in the uploaded PDF.')
