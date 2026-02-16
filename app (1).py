
import tempfile
import streamlit as st
from engine import build_master
from report import export_excel, export_pdf

st.set_page_config(page_title="Portfolio Consolidator", layout="wide")
st.title("Portfolio Consolidator (FINAL)")

col1, col2, col3 = st.columns(3)
with col1:
    f_new = st.file_uploader("NEW PORTFOLIOS.xlsx", type=["xlsx"])
    f_yasser = st.file_uploader("Yasser.xlsx", type=["xlsx"])
with col2:
    f_cfh = st.file_uploader("CFH.xlsx", type=["xlsx"])
    f_pos = st.file_uploader("positions_by_group.xlsx", type=["xlsx"])
with col3:
    f_cust = st.file_uploader("Customers Position.xlsx", type=["xlsx"])

if st.button("Process", type="primary"):
    missing = [n for n,f in [("NEW PORTFOLIOS.xlsx",f_new),("Yasser.xlsx",f_yasser),("CFH.xlsx",f_cfh),("positions_by_group.xlsx",f_pos),("Customers Position.xlsx",f_cust)] if f is None]
    if missing:
        st.error("Missing files: " + ", ".join(missing))
        st.stop()

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as d:
        p_new=f"{d}/NEW PORTFOLIOS.xlsx"
        p_y=f"{d}/Yasser.xlsx"
        p_c=f"{d}/CFH.xlsx"
        p_p=f"{d}/positions_by_group.xlsx"
        p_u=f"{d}/Customers Position.xlsx"

        for up, path in [(f_new,p_new),(f_yasser,p_y),(f_cfh,p_c),(f_pos,p_p),(f_cust,p_u)]:
            with open(path, "wb") as out:
                out.write(up.getbuffer())

        summary_df, holdings_df, totals_df, matrix_df = build_master(p_new,p_y,p_c,p_p,p_u)

        out_xlsx=f"{d}/Consolidated.xlsx"
        out_pdf=f"{d}/Consolidated.pdf"
        export_excel(out_xlsx, summary_df, holdings_df, totals_df, matrix_df)
        export_pdf(out_pdf, totals_df, matrix_df, title="Master Allocation Comparison (Consolidated)")

        st.success("Done.")
        st.download_button("Download Excel", open(out_xlsx,"rb").read(), file_name="Consolidated.xlsx")
        st.download_button("Download PDF", open(out_pdf,"rb").read(), file_name="Consolidated.pdf")
