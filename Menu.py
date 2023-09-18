
import streamlit as st
import pandas as pd



st.set_page_config(page_title='SEC Test')
st.title('SEC TEST')
st.subheader('<--- Choose an action from the left menu')


st.sidebar.markdown("# Main page 🎈")
###################create options as buttons##################
with st.form("Set file path"):
    st.write("Remember to download PO data into Testdata.xls")
    in_file = st.file_uploader("Path to Testdata.xls")
    submitted = st.form_submit_button("Set file path")
    if submitted:
        st.session_state.file = in_file
        st.write("Path is set to: "+ str(in_file))


##################retrieve info from excel ***************************
