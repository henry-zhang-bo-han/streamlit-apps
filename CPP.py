import base64
import json
import time
from datetime import datetime
from io import BytesIO

import streamlit as st
from openai import OpenAI
from pdf2image import convert_from_bytes

IMG2JSON_SYSTEM_PROMPT = '''
You are an expert document image analyzer. Given an image of a document page, you will answer specific user questions
based on the content of the image. You will transcribe the text in the image word by word whenever possible.
You will not paraphrase any content nor add any clarifying statements.
You will answer user questions and provide the answers in a JSON format specified by the user.
'''

IMG2JSON_USER_PROMPT = '''
Fill in the blanks with the correct answers based on the image content. Return the answers in the JSON format below:
'''

FIELDS_LIST = [
    {'ID': 'Social Insurance Number', 'type': 'int', 'page': 0},
    {'ID': 'Preferred Language', 'type': 'str', 'options': ['English', 'French'], 'page': 0},
    {'ID': 'First Name', 'type': 'str', 'page': 0},
    {'ID': 'Last Name', 'type': 'str', 'page': 0},
    {'ID': 'Date of Birth', 'type': 'str', 'format': 'YYYY-MM-DD', 'page': 0},
    {'ID': 'Address', 'type': 'str', 'page': 0},
    {'ID': 'Telephone', 'type': 'int', 'page': 0},
    {'ID': 'Email', 'type': 'str', 'page': 0},
    {'ID': 'Branch Number', 'type': 'int', 'page': 0},
    {'ID': 'Institution Number', 'type': 'int', 'page': 0},
    {'ID': 'Account Number', 'type': 'int', 'page': 0},
    {'ID': 'Name on the Account', 'type': 'str', 'page': 0},
    {'ID': 'Pension Sharing with Spouse', 'type': 'str', 'options': ['Yes', 'No', 'Not Applicable'], 'page': 1},
    {'ID': 'Other Country', 'type': 'str', 'page': 1},
    {'ID': 'Other Country from Date', 'type': 'str', 'format': 'YYYY-MM-DD', 'page': 1},
    {'ID': 'Other Country to Date', 'type': 'str', 'format': 'YYYY-MM-DD', 'page': 1},
    {'ID': 'Other Country Insurance Number', 'type': 'int', 'page': 1},
    {'ID': 'Applied in Other Country', 'type': 'str', 'options': ['Yes', 'No'], 'page': 1},
    {'ID': 'Separated or Divorced', 'type': 'str', 'options': ['Yes', 'No'], 'page': 1},
    {'ID': 'Current Marital Status', 'type': 'str',
     'options': ['Single', 'Married', 'Separated', 'Divorced', 'Common-law', 'Surviving spouse or common-law partner'],
     'page': 1},
    {'ID': 'Pension Start', 'type': 'str', 'options': ['As soon as I qualify', 'As of'], 'page': 5},
    {'ID': 'As of Date', 'type': 'str', 'format': 'YYYY-MM', 'page': 5},
    {'ID': 'Deduct Federal Income Tax', 'type': 'str', 'options': ['Yes', 'No'], 'page': 5},
    {'ID': 'Federal Income Tax ($)', 'type': 'int', 'page': 5},
    {'ID': 'Federal Income Tax (%)', 'type': 'int', 'page': 5},
    {'ID': 'Applicant Signature', 'type': 'str', 'page': 6},
    {'ID': 'First Name of Witness', 'type': 'str', 'page': 6},
    {'ID': 'Last Name of Witness', 'type': 'str', 'page': 6},
    {'ID': 'Telephone of Witness', 'type': 'int', 'page': 6},
    {'ID': 'Address of Witness', 'type': 'str', 'page': 6},
    {'ID': 'Signature of Witness', 'type': 'str', 'page': 6}
]

PAGES = set([field['page'] for field in FIELDS_LIST])
FIELD_OPTIONS = {field['ID']: field['options'] for field in FIELDS_LIST if 'options' in field}

CALCULATION_FIELDS = [
    {'ID': 'Baseline Monthly Payment (Starting at Age 65)', 'type': 'float'},
    {'ID': 'Corrected Start Date', 'type': 'str', 'format': 'YYYY-MM-DD'},
    {'ID': 'Delta Between Corrected Start Date and Date Turning 65 in Months', 'type': 'int'},
    {'ID': 'Percentage Impact on Monthly Payment', 'type': 'float'},
    {'ID': 'Adjusted Monthly Payment', 'type': 'float'},
    {'ID': 'Rationale', 'type': 'str'}
]


def convert_pdf_to_images(pdf):
    images = convert_from_bytes(pdf, fmt='png', thread_count=8)
    return images


def encode_image(image):
    bfr = BytesIO()
    image.save(bfr, format='png')
    return base64.b64encode(bfr.getvalue()).decode('utf-8')


def construct_img2json_user_prompt(fields_list, i):
    filtered_field_list = [field for field in fields_list if field['page'] == i]
    user_prompt_list = [IMG2JSON_USER_PROMPT, '{']
    for field in filtered_field_list:
        format_string = f'"{field["ID"]}": type {field["type"]}'
        if 'options' in field:
            format_string += f', choose one from {field["options"]}'
        if 'format' in field:
            format_string += f', format {field["format"]}'
        user_prompt_list.append(format_string)
    user_prompt_list.append('}')
    return '\n'.join(user_prompt_list)


def extract_json_from_image(client, image, user_prompt):
    messages = [
        {'role': 'system', 'content': IMG2JSON_SYSTEM_PROMPT},
        {'role': 'user', 'content': [
            {'type': 'text', 'text': user_prompt},
            {'type': 'image_url', 'image_url': {'url': f'data:image/png;base64,{encode_image(image)}'}}
        ]}
    ]
    completion = client.chat.completions.create(
        model="gpt-4o",
        response_format={"type": "json_object"},
        messages=messages
    )
    return json.loads(completion.choices[0].message.content)


def convert_id_to_key(id_string):
    return 'input_' + id_string.lower().replace(' ', '_')


def process_uploaded_pdf():
    uploaded_file = st.session_state['uploaded_pdf']
    if uploaded_file is not None:
        with st.status('Processing uploaded PDF...', expanded=False) as status:
            # Save uploaded PDF
            status.update(label='Uploading PDF ...')

            # Convert PDF to images
            status.update(label='Scanning PDF pages ...')
            screenshots = convert_pdf_to_images(uploaded_file.getvalue())

            # Extract JSON from images
            st.session_state['extracted_values'] = {}
            for i in PAGES:
                status.update(label=f'Extracting answers from page {i + 1} ...')
                user_prompt = construct_img2json_user_prompt(FIELDS_LIST, i)
                json_data = extract_json_from_image(openai_client, screenshots[i], user_prompt)
                st.session_state['extracted_values'][i] = json_data

            # Finalize processing
            status.update(label='PDF successfully processed ...')
            st.session_state['has_uploaded_pdf'] = True
    else:
        st.session_state.pop('has_uploaded_pdf', None)
        st.session_state.pop('extracted_values', None)
        st.session_state.pop('age_eligibility_assessment', None)
        st.session_state.pop('past_contributions_assessment', None)
        st.session_state.pop('cpp_eligible', None)
        st.session_state['toggle_confirm_accuracy'] = False


def determine_age_eligibility(date_of_birth, client):
    today = datetime.today().strftime('%Y-%m-%d')
    user_prompt = f'''
    The applicant was born on {date_of_birth} (YYYY-MM-DD).
    Today's date is {today} (YYYY-MM-DD).
    For the applicant to quality for CPP, they must be at least 60 years old as of today.
    If the applicant is at least 60 years old, return AGE_REQUIREMENT_MET. Otherwise, return AGE_REQUIREMENT_NOT_MET.
    In addition, provide the rationale of the assessment in one sentence.
    '''
    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[{'role': 'user', 'content': user_prompt}]
    )
    return completion.choices[0].message.content


def calculate_payment(date_of_birth, pension_start, start_date, contribution_amount, client):
    # Determine start date of payment
    date_of_birth = datetime.strptime(date_of_birth, '%Y-%m-%d')
    date_turning_60 = date_of_birth.replace(year=date_of_birth.year + 60)
    date_turning_65 = date_of_birth.replace(year=date_of_birth.year + 65)
    date_turning_70 = date_of_birth.replace(year=date_of_birth.year + 70)

    if pension_start == 'As soon as I qualify':
        pension_start_date = max(date_turning_60, datetime.today())
    else:
        pension_start_date = datetime.strptime(start_date, '%Y-%m')
        pension_start_date = min(pension_start_date, date_turning_70)
        pension_start_date = max(pension_start_date, date_turning_60)

    # Calculate the difference with 65th birthday
    delta_months = ((pension_start_date.year - date_turning_65.year) * 12 +
                    pension_start_date.month - date_turning_65.month)

    # Calculate the percentage difference in payment
    percentage_impact = 0.0
    if delta_months < 0:
        percentage_impact = max(0.6 * delta_months, -36)
    elif delta_months > 0:
        percentage_impact = min(0.7 * delta_months, 42)

    # Calculate the adjusted monthly payment
    payment = contribution_amount / 240 * (1 + percentage_impact / 100)

    # Construct the user prompt
    user_prompt = f'''
    The applicant was born on {date_of_birth} (YYYY-MM-DD).
    The applicant intends to start receiving CPP pension benefits on {start_date} (YYYY-MM).
    The applicant should start receiving benefits between the age of 60 and 70.
    The adjusted pension start date should be {pension_start_date.strftime('%Y-%m-%d')}.
    The applicant has made a total of ${contribution_amount:,.2f} in CPP contributions.
    The total contributions should be spread across 20 years if the applicant starts receiving benefits at age of 65.
    If the pension starts before age 65, it will be reduced by 0.6% for each month (or 7.2% per year) before the 65th
    birthday. The maximum reduction is 36% if the applicant starts their pension at age 60. This reduction is permanent.
    If the pension starts after age 65, it will increase by 0.7% for each month (or 8.4% per year) after the 65th
    birthday. The maximum increase is 42% if starting the pension at age 70. The increase is permanent, and there
    are no more increases after age 70.
    The percentage delta between the projected monthly payment and the baseline monthly payment is {percentage_impact}%.
    As such, the adjusted monthly payment is ${payment:,.2f}.
    ---
    Write a clear paragraph explaining the rationale behind the calculation. Retain key figures.
    Do not have more than 2 decimal places for any number.
    '''

    # Use OpenAI to make calculations
    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[{'role': 'user', 'content': user_prompt}]
    )

    return {
        'pension_start_date': pension_start_date,
        'delta_months': delta_months,
        'percentage_impact': percentage_impact,
        'payment': payment,
        'rationale': completion.choices[0].message.content

    }


def toggle_inputs():
    if st.session_state['toggle_confirm_accuracy'] is False:
        st.session_state.pop('age_eligibility_assessment', None)
        st.session_state.pop('past_contributions_assessment', None)
        st.session_state.pop('cpp_eligible', None)
        st.session_state.pop('payment_calculation', None)


def toggle_eligibility():
    if st.session_state['toggle_confirm_eligibility'] is False:
        st.session_state.pop('payment_calculation', None)


if __name__ == '__main__':
    # Page title
    st.title('CPP Application Processing')

    # Initialize OpenAI Client
    openai_client = OpenAI(
        api_key=st.secrets['OPENAI_API_KEY']
    )

    # Upload scanned PDF for processing
    uploaded_pdf = st.file_uploader(
        'Upload a scanned CPP application form (ISP-1000)',
        type='pdf',
        on_change=process_uploaded_pdf,
        key='uploaded_pdf'
    )

    # Display uploaded PDF
    if 'has_uploaded_pdf' in st.session_state:
        st.divider()
        st.subheader('Validate AI extracted inputs')
        for idx in PAGES:
            if 'toggle_confirm_accuracy' not in st.session_state:
                expanded = True
            else:
                expanded = not st.session_state['toggle_confirm_accuracy']

            with st.expander(f'Inputs extracted from page {idx + 1}', expanded=expanded):
                for k, v in st.session_state['extracted_values'][idx].items():
                    if k in FIELD_OPTIONS:
                        st.selectbox(k, options=FIELD_OPTIONS[k], key=convert_id_to_key(k))
                    else:
                        st.text_input(k, value=v, key=convert_id_to_key(k))

        st.toggle(
            'I confirm the accuracy of extracted inputs',
            key='toggle_confirm_accuracy',
            on_change=toggle_inputs
        )

    # Display eligibility assessment
    if 'toggle_confirm_accuracy' in st.session_state and st.session_state['toggle_confirm_accuracy']:
        st.divider()
        st.subheader('Confirm AI assessed eligibility')

        # Assess age eligibility
        if 'age_eligibility_assessment' not in st.session_state:
            with st.spinner('Assessing age eligibility ...'):
                st.session_state['age_eligibility_assessment'] = determine_age_eligibility(
                    st.session_state['input_date_of_birth'],
                    openai_client
                )

        first_name = st.session_state['input_first_name']

        if age_eligible := 'AGE_REQUIREMENT_MET' in st.session_state['age_eligibility_assessment']:
            age_eligibility_header = f'✅ {first_name} is 60+ years old'
        else:
            age_eligibility_header = f'❌ {first_name} is under 60 years old'

        with st.expander(age_eligibility_header):
            st.write(st.session_state['age_eligibility_assessment'])

        # Assess past contributions
        if 'past_contributions_assessment' not in st.session_state:
            with st.spinner('Assessing past contributions ...'):
                time.sleep(2)
                if 'CONTRIBUTIONS' in st.secrets:
                    st.session_state['past_contributions_assessment'] = float(st.secrets['CONTRIBUTIONS'])
                else:
                    st.session_state['past_contributions_assessment'] = 0.0

        contributions = st.session_state['past_contributions_assessment']
        if contributions > 0:
            contributions_header = f'✅ {first_name} has made valid CPP contributions'
            contributions_body = f'Total past contributions: ${contributions:,.2f}'
        else:
            contributions_header = f'❌ {first_name} has no past CPP contributions'
            contributions_body = f'No past contributions made to CPP'

        with st.expander(contributions_header):
            st.write(contributions_body)

        # Assess overall eligibility
        if 'cpp_eligible' not in st.session_state:
            st.session_state['cpp_eligible'] = age_eligible and contributions > 0

        if st.session_state['cpp_eligible']:
            st.write(f'✅ {first_name} is eligible for CPP benefits')
        else:
            st.write(f'❌ {first_name} is not eligible for CPP benefits')

        st.toggle(
            'I confirm the correctness of eligibility assessments',
            key='toggle_confirm_eligibility',
            on_change=toggle_eligibility
        )

    # Display payment calculation
    if 'toggle_confirm_eligibility' in st.session_state and st.session_state['toggle_confirm_eligibility']:
        if 'cpp_eligible' in st.session_state and st.session_state['cpp_eligible']:
            st.divider()
            st.subheader('Confirm AI calculated payment')

            # Calculate payment
            if 'payment_calculation' not in st.session_state:
                with st.spinner('Calculating payment ...'):
                    st.session_state['payment_calculation'] = calculate_payment(
                        st.session_state['input_date_of_birth'],
                        st.session_state['input_pension_start'],
                        st.session_state['input_as_of_date'],
                        st.session_state['past_contributions_assessment'],
                        openai_client
                    )

            col1, col2, col3 = st.columns(3)

            with col1:
                with st.container(border=True):
                    st.metric('Past contributions', f'${st.session_state["past_contributions_assessment"]:,.2f}',
                              'in total', delta_color='off')

            with col2:
                with st.container(border=True):
                    st.metric(
                        'Pension start date',
                        st.session_state['payment_calculation']['pension_start_date'].strftime('%Y-%m-%d'),
                        f"{st.session_state['payment_calculation']['delta_months']} mo from 65th birthday"
                    )

            with col3:
                with st.container(border=True):
                    st.metric('Adjusted monthly payment',
                              f'${st.session_state["payment_calculation"]["payment"]:,.2f}',
                              f'{st.session_state["payment_calculation"]["percentage_impact"]:.2f}% vs. baseline')

            with st.expander(
                    f'Detailed rationale behind ${st.session_state["payment_calculation"]["payment"]:,.2f} calculation',
                    expanded=True):
                st.write(st.session_state['payment_calculation']['rationale'].replace('$', '\\$'))

            st.toggle('I confirm the correctness of payment calculation', key='toggle_confirm_payment')

    # Display communication button
    if 'toggle_confirm_eligibility' in st.session_state and st.session_state['toggle_confirm_eligibility']:
        if 'cpp_eligible' in st.session_state and st.session_state['cpp_eligible']:
            if 'toggle_confirm_payment' in st.session_state and st.session_state['toggle_confirm_payment']:
                st.divider()
                st.button('Approve CPP application', on_click=st.balloons)
        elif 'cpp_eligible' in st.session_state and not st.session_state['cpp_eligible']:
            st.divider()
            st.button('Reject CPP application', on_click=st.snow)
