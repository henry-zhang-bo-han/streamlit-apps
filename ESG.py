import streamlit as st

with open('data/esg.md') as f:
    st.title('Merck ESG Report')
    st.write(f.read())
