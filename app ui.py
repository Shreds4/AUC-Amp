
import streamlit as st
import pandas as pd
from final_data_analysis import normalize_data, analyse_intervals

st.title("📊 AUC & Amplitude Analysis Tool")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
cutpoints = st.text_input("Enter timepoints to split (comma-separated)", "2830")

if uploaded_file and cutpoints:
    try:
        st.success("File uploaded. Processing...")
        timepoints = [int(tp.strip()) for tp in cutpoints.split(',') if tp.strip().isdigit()]

        original_df = pd.read_excel(uploaded_file)
        normalized_df = normalize_data(uploaded_file)

        auc_tables, auc_sums, amp_df, meta_df = analyse_intervals(normalized_df, timepoints)

        st.subheader("📈 Normalized Data")
        st.dataframe(normalized_df)

        st.subheader("📐 AUC Sums")
        st.dataframe(auc_sums)

        st.subheader("📏 Amplitudes")
        st.dataframe(amp_df)

        st.subheader("📋 Summary Stats")
        st.dataframe(meta_df)

    except Exception as e:
        st.error(f"Something went wrong: {e}")
