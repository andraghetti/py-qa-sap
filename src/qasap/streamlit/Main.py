import sys
from pathlib import Path

import pandas as pd

import streamlit as st
from streamlit import runtime
from streamlit.web import cli as stcli

from qasap.models.sheet import EXTRA_SHEETS, Introduction, MasterDetails

CURRENT_DIRECTORY = Path(__file__).parent
STATIC_DIR = CURRENT_DIRECTORY / "static"
CSS_FILE = CURRENT_DIRECTORY / "style.css"


def add_header():
    st.title("Migration QA")


def add_hr(margin_rem=0):
    st.markdown(
        f"""<hr style="height:1px;background-color:#224;margin-top:{margin_rem}rem;margin-bottom:{margin_rem}rem;"/>""",
        unsafe_allow_html=True,
    )


def set_page_config():
    st.set_page_config(
        page_title="Migration QA - PyQASAP",
        page_icon=str(STATIC_DIR / "favicon.ico"),
        menu_items={
            "About": "https://github.com/andraghetti/py-qa-sap",
        },
    )


def set_style():
    st.markdown(f"<style>{CSS_FILE.read_text()}</style>", unsafe_allow_html=True)


def add_file_input():
    file = st.file_uploader(label="file-uploader", key="file-uploader-main")
    if file:
        mime_type = file.type
        if mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            return file
        else:
            st.error(f"Please upload an Excel (xslx) file. Got {mime_type}")


def read_xls(file):
    if not file:
        return
    xls = pd.read_excel(file, sheet_name=None, engine="openpyxl", skiprows=[])
    return xls


def select_sheets(xls):
    sheet_names = xls.keys()
    with st.expander("Click to expand", expanded=True):
        num_columns = 3  # Number of columns you want
        cols = st.columns(num_columns)
        selected_sheets = []
        for i, sheet in enumerate(sheet_names):
            col = cols[i % num_columns]
            if col.checkbox(sheet, key=sheet):
                selected_sheets.append(sheet)
    return selected_sheets


color_to_filed_status = {"Mandatory": "green", "Optional": "orange", "Blank": "red"}


def render_fields_radiobox(sheet_name, sheet):
    with st.expander("Click to expand", expanded=True):
        all_fields = sheet.loc[3]
        # questo solo per maste details, qui dovremo mettere sheet.fields

        num_columns = 4
        cols = st.columns(num_columns)
        for i, field in enumerate(all_fields):
            with cols[i % num_columns]:
                color = color_to_filed_status["Optional"]
                selected = st.radio(
                    f":{color}[{field}]",
                    ["Mandatory", "Optional", "Blank"],
                    key=f"{sheet_name}-{i}-{field}",
                )
                # status[field] = selected


def render_selected_sheets(xls, selected_sheets):
    for sheet in selected_sheets:
        if sheet == "Introduction":
            Introduction(xls[sheet]).render()
        if sheet == "Master Details":
            MasterDetails(xls[sheet]).render()

        if sheet not in [s.name for s in EXTRA_SHEETS]:
            render_fields_radiobox(sheet, xls[sheet])
            st.table(xls[sheet])


def main():
    set_page_config()
    set_style()

    add_header()

    file = add_file_input()
    if not file:
        return
    xls = read_xls(file)
    if not xls:
        return
    selected_sheets = select_sheets(xls)
    render_selected_sheets(xls, selected_sheets)


if __name__ == "__main__":
    if runtime.exists():
        main()
    else:
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())
