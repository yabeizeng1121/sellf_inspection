import streamlit as st
import pandas as pd
from datetime import datetime
import io
import zipfile
from collections import Counter
import random

# Session state storage for memory caching
if "final_data" not in st.session_state:
    st.session_state.final_data = pd.DataFrame()

# Reason options
REASONS = {
    "00": "Qualified",
    "01": "No Address Info",
    "02": "Location Not Clear",
    "03": "No Clear Shipping Label",
    "04": "Public or Unsafe Area",
    "05": "Invalid Mailbox Delivery",
    "06": "Leave Outside of Building",
    "07": "Wrong Address",
    "08": "Wrong Parcel Photo",
    "09": "No POD",
    "10": "Inappropriate Delivery",
}


def process_file(file, selected_date):
    df = pd.read_excel(file)
    df = df.drop(columns=["199_pathtime"], errors="ignore")
    df = df.drop_duplicates()
    df = df[~df["service_number"].astype(str).str.startswith("550")]
    df = df[df["state"] == 203]

    grouped = df.groupby("service_number")
    sampled_df = grouped.apply(
        lambda x: x.sample(n=min(30, len(x)), random_state=42)
    ).reset_index(drop=True)
    return sampled_df


def user_input_interface(df, selected_date):
    records = []
    for service in df["service_number"].unique():
        st.markdown(f"### Service Number: {service}")
        tno_list = df[df["service_number"] == service]["tno"].tolist()
        tno_string = "\n".join(map(str, tno_list))
        st.text_area("📋 Copy these 15 TNOs:", value=tno_string, height=100)

        dsp_name = st.text_input(f"DSP Name for {service}", key=f"dsp_{service}")
        subset = df[df["service_number"] == service]
        for i, row in subset.iterrows():
            st.markdown(
                f"**TNO: {row['tno']} | Driver ID: {row.get('Driver id', 'N/A')}**"
            )
            qualified = st.radio(
                f"Is this a qualified POD?", ["Yes", "No"], key=f"qual_{i}"
            )
            reason = "00"
            if qualified == "No":
                reason = st.selectbox(
                    "Select Fail Reason",
                    list(REASONS.keys())[1:],
                    format_func=lambda x: REASONS[x],
                    key=f"reason_{i}",
                )
            records.append(
                {
                    "tno": row["tno"],
                    "DSP": dsp_name,
                    "Date": selected_date.strftime("%Y-%m-%d"),
                    "Quality": "Yes" if qualified == "Yes" else "No",
                    "Reason": REASONS[reason],
                    "Driver id": row["service_number"],
                }
            )
    return pd.DataFrame(records)


def generate_reports(final_df):
    grouped = final_df.groupby("DSP")
    summary_blocks = []

    files = {}
    for dsp, group in grouped:
        total = len(group)
        qualified_count = sum(group["Quality"] == "Yes")
        rate = round(qualified_count / total * 100, 2)

        if rate == 100:
            zh_summary = f"中文版：今天【{dsp}】POD抽查共【{total}】件，100%合格， 不错继续保持！"
            en_summary = f"English: Today, DSP {dsp} has {total} PODs checked, 100% qualified. Great job, keep it up!"
        else:
            service_rates = group.groupby("Driver id").apply(
                lambda x: (x["Quality"] == "No").sum()
            )
            main_service = service_rates.idxmax()
            reason_mode = group[group["Driver id"] == main_service]["Reason"]
            most_common_reason = Counter(reason_mode).most_common(1)[0][0]
            zh_summary = f"中文版：今天【{dsp}】POD共查【{total}】件，合格率为【{rate}%】，其中司机【{main_service}】有不合格件，主要原因是【{most_common_reason}】"
            en_summary = f"English: Today, DSP {dsp} had {total} PODs checked with a {rate}% pass rate. Service number {main_service} had some failures, mainly due to {most_common_reason}."

        summary_blocks.append((zh_summary, en_summary))

        # Write group data only
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            group.to_excel(writer, index=False, sheet_name="Result")
        output.seek(0)
        files[f"{dsp}_report.xlsx"] = output.read()

    # Display all summaries
    for zh, en in summary_blocks:
        st.markdown(f"**{zh}**")
        st.markdown(f"*{en}*")
        st.markdown("---")

    # Zip all files
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
        for name, content in files.items():
            zip_file.writestr(name, content)
    zip_buffer.seek(0)
    return zip_buffer


# Streamlit interface
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to", ["📤 Upload & Inspect", "📊 Report"])

if page == "📤 Upload & Inspect":
    st.title("Excel POD Qualification Checker")
    today = st.date_input("📅 Select Today's Date", value=datetime.today())
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if uploaded_file:
        st.subheader("Step 1: Filter and Sample")
        sampled_df = process_file(uploaded_file, today)

        st.subheader("Step 2: Review and Annotate")
        final_data = user_input_interface(sampled_df, today)
        if st.button("✅ Save Results"):
            st.session_state.final_data = final_data
            st.success("Results saved! You can now go to the Report page.")
            # Create downloadable Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                final_data.to_excel(
                    writer, index=False, sheet_name="Final Annotated Data"
                )
            output.seek(0)

            st.download_button(
                label="📥 Download Final Annotated Excel",
                data=output,
                file_name="Final_Annotated_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.title("DSP Quality Report Generator")
    if not st.session_state.final_data.empty:
        zip_file = generate_reports(st.session_state.final_data)
        st.download_button(
            "📥 Download All DSP Reports (ZIP)", zip_file, file_name="DSP_Reports.zip"
        )
    else:
        st.warning("❗ Please first upload and annotate data in the previous page.")
