import base64
import json
import streamlit as st
import pytesseract
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes

IMG2MD_SYSTEM_PROMPT = '''
You are an expert document parser.
Given the image of a document page, you transcribe all text of the page into Markdown format accurately and clearly.

# Instructions 
- Your transcription should include all the text in the image.
- You could reference the text output from OCR output to achieve higher accuracy.
- Your output should be formatted clearly, maintaining the original content structure shown in the image.

# Output format
- Transcribe the document word-for-word whenever possible. Do not paraphrase any content nor add clarifying statements.
- Do not split a single table into multiple tables. However, keep different tables separate.
- Do not split content into multiple lines unless ABSOLUTELY NECESSARY.
- Output only the transcribed content. Do not reply with other content (e.g., here is the content of the document).
- Only wrap the Markdown tables in markdown code blocks. Do not wrap other content in code blocks.
'''

IMG_TITLE_USER_PROMPT = '''
Here is the output from OCR for reference:
{ocr}
---
Please parse the image content and transcribe into a Markdown format, while maintaining the original structure. 
'''

MD2JSON_SYSTEM_PROMPT = '''
You are an expert in converting Markdown tables into well-formatted JSON objects.

# Instructions
- If there is no Markdown table in the provided content, return {"tables": "NO_TABLE_PRESENT"} and terminate.
- If there are Markdown tables present in the content, focus on the Markdown tables only and ignore the rest.
- For each table, extract its title.
- For each table, convert the Markdown table's body into a JSON list of rows.
- Each row should be a list of items, where each item represents the value of a cell.
- For each cell, if the value is a number, convert the string into an integer or a float.
- Apply the correct magnitude for each cell (e.g., multiply by 1,000,000 if the cell should be in millions).
'''

MD2JSON_USER_PROMPT = '''
For each table present, please return a JSON object that contains the title and body (a list of rows).
Each row should be represented by a list of cells.
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
            {'type': 'text', 'text': IMG_TITLE_USER_PROMPT.format(ocr=ocr_string)},
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


def process_uploaded_pdf(client):
    uploaded_file = st.session_state['uploaded_pdf']

    if uploaded_file is not None:
        with st.status('Scanning uploaded PDF pages ...') as status:
            # Convert PDF pages into images
            images = convert_from_bytes(uploaded_file.getvalue(), fmt='png', thread_count=8)
            st.session_state['images'] = images

            # Extract table titles from images
            text_extracts = []
            table_extracts = []
            for i, image in enumerate(images):
                status.update(label=f'Reading tables from page #{i+1} ...')

                # Convert to string using OCR
                ocr_string = pytesseract.image_to_string(image)

                # Use OpenAI to extract Markdown text content from image
                text_from_image = extract_tables_from_image(client, image, ocr_string)
                text_extracts.append(text_from_image)

                # Use OpenAI to convert Markdown to JSON
                table_json = convert_markdown_to_json(client, text_from_image)
                if table_json['tables'] != 'NO_TABLE_PRESENT' and type(table_json['tables']) == list:
                    table_extracts.extend(table_json['tables'])

            status.update(label='PDF processing complete')
            st.session_state['text_extracts'] = text_extracts
            st.session_state['table_extracts'] = table_extracts


if __name__ == '__main__':
    st.title('Convert PDF to Excel')

    # Initialize OpenAI Client
    openai_client = OpenAI(api_key=st.secrets['OPENAI_API_KEY'])

    # Upload scanned PDF for processing
    uploaded_pdf = st.file_uploader(
        'Upload a scanned CPP application form (ISP-1000)',
        type='pdf',
        on_change=process_uploaded_pdf,
        args=(openai_client,),
        key='uploaded_pdf'
    )

    # Show tables
    if 'table_extracts' in st.session_state:
        for t in st.session_state['table_extracts']:
            st.subheader(t['title'])
            st.json(t['body'])
