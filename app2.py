"""
JKCL AP Allocation Tool — Streamlit UI
Run:  streamlit run app.py
Any user: upload the Exela DCN-level detail dump -> download the publish-ready report.
"""
import io, os, json, datetime as dt
import pandas as pd
import streamlit as st
import allocation_engine as E
import build_output as B

st.set_page_config(page_title="JKCL AP Allocation Tool", layout="wide")
HERE = os.path.dirname(os.path.abspath(__file__))

st.markdown("""
<style>
 .stApp {background:#FAFBFC;}
 h1,h2,h3 {color:#1F3B4D; font-family:Calibri,Arial,sans-serif;}
 .metric-card{background:#fff;border:1px solid #E3E8EC;border-radius:10px;padding:14px 18px;}
</style>""", unsafe_allow_html=True)

st.title("JKCL AP — Allocation Tool")
st.caption("New allocation logic · user-independent · upload → classify → route → publish")

cfg = E.load_config()

with st.sidebar:
    st.subheader("How it works")
    st.write("1. Export the Exela **DCN-level detail** (TotalPending → Grand Total → Download Excel).")
    st.write("2. Drop it below.")
    st.write("3. Download the allocated, publish-ready workbook.")
    st.divider()
    st.subheader("Logic in force")
    st.write(f"• Vendor groups mapped: **{len(cfg['vendor_group_utility'])}**")
    st.write(f"• Roster: **{len(cfg['users'])}** users / 4 teams")
    st.write("• Routing: Freight → White → HO-Grey (by category) → Plant; Auto → API users")
    st.write("• Plant/HO-White: Parked→Parker, to-post→Poster, then load-balanced")
    with st.expander("Edit mappings (advanced)"):
        st.caption("All rules live in allocation_config.json — edit there to change logic without code.")

up = st.file_uploader("Exela DCN-level detailed report (.xlsx)", type=["xlsx"])

if up:
    tmp = os.path.join(HERE, "_uploaded.xlsx")
    with open(tmp, "wb") as f:
        f.write(up.getbuffer())
    with st.spinner("Classifying and allocating…"):
        res = E.run(tmp, cfg)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Invoices", f"{len(res):,}")
    c2.metric("Teams routed", res["Team"].nunique())
    c3.metric("Parker", f"{(res['Role']=='Parker').sum():,}")
    c4.metric("Poster", f"{(res['Role']=='Poster').sum():,}")
    c5.metric("Needs review", f"{(res['Needs Review']!='').sum():,}")

    a, b = st.columns([1, 1])
    with a:
        st.subheader("By team")
        st.bar_chart(res["Team"].value_counts())
    with b:
        st.subheader("By assignee (top 15)")
        st.bar_chart(res["Assigned To"].value_counts().head(15))

    if (res["Needs Review"] != "").any():
        st.warning(f"{(res['Needs Review']!='').sum()} rows flagged for review "
                   "(unmapped vendor group / blank / verify auto-plant). See the **Review** tab in the download.")

    st.subheader("Preview")
    st.dataframe(res.head(200), use_container_width=True, height=320)

    out_path = os.path.join(HERE, "JKCL_Allocation_Output.xlsx")
    B.build(res, out_path)
    with open(out_path, "rb") as f:
        st.download_button(
            "⬇ Download publish-ready workbook",
            f, file_name=f"Pendency allocation new logic - {dt.date.today():%d_%b'%y}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload the Exela detailed report to begin.")
