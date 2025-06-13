
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Souhrn metrik", layout="wide")
st.title("📊 Automatický výpočet metrik z Excelu")

def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df["Uzavřená práce"] = pd.to_datetime(df["Uzavřená práce"], errors="coerce")
    df["Datum"] = df["Uzavřená práce"].dt.date

    # Součet SKP
    df_prodej_vydat = df[(df["ID pracovní třídy"] == "Prodej") & (df["Typ práce"] == "Vydat")].copy()
    df_prodej_vydat["Množství upravené"] = df_prodej_vydat.apply(
        lambda row: row["Množství práce"] * 24 if row["Jednotka"] == "PAL" else row["Množství práce"], axis=1
    )
    soucet_skp = df_prodej_vydat.groupby("Datum")["Množství upravené"].sum().reset_index(name="Součet SKP")

    # Počet SKU
    pocet_sku = df_prodej_vydat.groupby("Datum").size().reset_index(name="Počet SKU")

    # Počet objednávek
    pocet_obj = df_prodej_vydat.groupby("Datum")["Číslo objednávky"].nunique().reset_index(name="Počet objednávek")

    # Počet natónovaných SKP
    df_tonovane = df[
        (df["ID pracovní třídy"].isin(["PO_Pozn", "Výroba"])) & (df["Typ práce"] == "Vložit")
    ].copy()
    df_tonovane["Množství upravené"] = df_tonovane.apply(
        lambda row: row["Množství práce"] * 24 if row["Jednotka"] == "PAL" else row["Množství práce"], axis=1
    )
    tonovane_skp = df_tonovane.groupby("Datum")["Množství upravené"].sum().reset_index(name="Počet natónovaných SKP")

    # Počet tónovaných objednávek (počítáno jako počet řádků)
    tonovane_obj = df_tonovane.groupby("Datum").size().reset_index(name="Počet tónovaných objednávek")

    # Sloučení všech metrik
    df_summary = soucet_skp.merge(pocet_sku, on="Datum", how="outer") \
                           .merge(pocet_obj, on="Datum", how="outer") \
                           .merge(tonovane_skp, on="Datum", how="outer") \
                           .merge(tonovane_obj, on="Datum", how="outer") \
                           .sort_values("Datum")

    return df_summary

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Souhrn")
    output.seek(0)
    return output

uploaded_file = st.file_uploader("Nahraj Excel soubor", type=["xlsx"])

if uploaded_file:
    df_summary = process_file(uploaded_file)
    st.success("✅ Metriky byly úspěšně spočítány.")
    st.dataframe(df_summary, use_container_width=True)

    excel_data = to_excel(df_summary)
    st.download_button(
        label="📥 Stáhnout výsledky jako Excel",
        data=excel_data,
        file_name="souhrn_metrik.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
