
import tempfile
import streamlit as st
from engine import build_master
from report import export_excel, export_pdf

st.set_page_config(page_title="Portfolio Consolidator", layout="wide")
st.title("Portfolio Consolidator")
st.caption("Upload any subset of templates. The app will process only what you upload.")

st.subheader("FX Settings")
fx_usd_to_egp = st.number_input("USD → EGP exchange rate (Arqaam NAV/Cash are in USD)", min_value=0.0, value=0.0, step=0.1)
fx_usd_to_egp = None if fx_usd_to_egp == 0.0 else float(fx_usd_to_egp)

colA, colB, colC = st.columns(3)

with colA:
    use_arqaam = st.checkbox("Include Arqaam (was New Portfolios)", value=True)
    f_new = st.file_uploader("Arqaam template (NEW PORTFOLIOS.xlsx)", type=["xlsx"], disabled=not use_arqaam)
    use_yasser = st.checkbox("Include Yasser (Yasser + R&R)", value=True)
    f_yasser = st.file_uploader("Yasser.xlsx", type=["xlsx"], disabled=not use_yasser)

with colB:
    use_cfh = st.checkbox("Include CFH", value=True)
    f_cfh = st.file_uploader("CFH.xlsx", type=["xlsx"], disabled=not use_cfh)
    use_pos = st.checkbox("Include Positions by Group", value=True)
    f_pos = st.file_uploader("positions_by_group.xlsx", type=["xlsx"], disabled=not use_pos)

with colC:
    use_emad = st.checkbox("Include Emad (Retail>5M)", value=True)
    f_cust = st.file_uploader("Customers Position.xlsx", type=["xlsx"], disabled=not use_emad)

if st.button("Process", type="primary"):
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as d:
        def save(uploaded, name):
            if uploaded is None:
                return None
            path = f"{d}/{name}"
            with open(path, "wb") as out:
                out.write(uploaded.getbuffer())
            return path

        p_new = save(f_new if use_arqaam else None, "NEW PORTFOLIOS.xlsx")
        p_y   = save(f_yasser if use_yasser else None, "Yasser.xlsx")
        p_c   = save(f_cfh if use_cfh else None, "CFH.xlsx")
        p_p   = save(f_pos if use_pos else None, "positions_by_group.xlsx")
        p_u   = save(f_cust if use_emad else None, "Customers Position.xlsx")

        summary_df, holdings_df, totals_df, matrix_df = build_master(
            new_portfolios_path=p_new,
            fx_usd_to_egp=fx_usd_to_egp,
            yasser_path=p_y,
            cfh_path=p_c,
            positions_by_group_path=p_p,
            customers_position_path=p_u,
        )

        out_xlsx = f"{d}/Consolidated.xlsx"
        out_pdf  = f"{d}/Consolidated.pdf"

        export_excel(out_xlsx, summary_df, holdings_df, totals_df, matrix_df)
        export_pdf(out_pdf, totals_df, matrix_df, title="Master Allocation Comparison (Consolidated)")

        st.success("Done.")
        st.download_button("Download Excel", open(out_xlsx,"rb").read(), file_name="Consolidated.xlsx")
        st.download_button("Download PDF", open(out_pdf,"rb").read(), file_name="Consolidated.pdf")
