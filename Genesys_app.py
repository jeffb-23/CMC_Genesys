import io
import pandas as pd
import streamlit as st

MAX_WIDTH = 22
MAX_LENGTH = 15
MAX_HEIGHT = 11

CARDBOARD_OPTIONS = [23, 39]  # inches

st.title("CMC Genesys Item Size Analysis")

st.write("Machine Max Box Size:")
st.write(f"- Width: {MAX_WIDTH} in")
st.write(f"- Length: {MAX_LENGTH} in")
st.write(f"- Height: {MAX_HEIGHT} in")

st.write("Cardboard Width Options:")
st.write("- 23 in")
st.write("- 39 in")

uploaded = st.file_uploader("Upload CSV or Excel (XLSX)", type=["csv", "xlsx"])

def load_file(file):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def pick_cardboard(width_series: pd.Series, height_series: pd.Series) -> pd.Series:
    """
    Chooses 23 if (W+H)<=23, else 39 if (W+H)<=39, else 'No Fit'.
    """
    wrap = width_series + height_series
    result = pd.Series(["No Fit"] * len(wrap), index=wrap.index, dtype="object")
    # set 39 first, then overwrite with 23 where applicable
    result.loc[wrap <= CARDBOARD_OPTIONS[1]] = str(CARDBOARD_OPTIONS[1])
    result.loc[wrap <= CARDBOARD_OPTIONS[0]] = str(CARDBOARD_OPTIONS[0])
    return result

if uploaded is not None:
    df = load_file(uploaded)

    st.subheader("Preview of Uploaded Data")
    st.dataframe(df.head())

    st.subheader("Column Mapping")
    columns = list(df.columns)

    col1, col2 = st.columns(2)
    with col1:
        item_col = st.selectbox("Select Item ID column", columns)
        height_col = st.selectbox("Select Height column", columns)
    with col2:
        width_col = st.selectbox("Select Width column", columns)
        length_col = st.selectbox("Select Length column", columns)

    if st.button("Process File"):
        # Standardize and coerce to numeric
        item_id = df[item_col]
        height = pd.to_numeric(df[height_col], errors="coerce")
        width = pd.to_numeric(df[width_col], errors="coerce")
        length = pd.to_numeric(df[length_col], errors="coerce")

        invalid = pd.DataFrame({"Height": height, "Width": width, "Length": length}).isna().any(axis=1)

        # Volume (only valid when all dimensions are valid)
        volume = height * width * length
        volume = volume.where(~invalid, pd.NA)

        # Machine fit
        fits_machine = (
            (height <= MAX_HEIGHT) &
            (width <= MAX_WIDTH) &
            (length <= MAX_LENGTH)
        )

        status = pd.Series(["No OK"] * len(df), dtype="object")
        status.loc[fits_machine & (~invalid)] = "OK"

        # Cardboard width (based on Width + Height)
        cardboard_width = pd.Series(["No Fit"] * len(df), dtype="object")
        valid_for_cardboard = (~pd.DataFrame({"Height": height, "Width": width}).isna().any(axis=1))
        cardboard_width.loc[valid_for_cardboard] = pick_cardboard(
            width.loc[valid_for_cardboard],
            height.loc[valid_for_cardboard]
        )

        # Build output with required column order
        df_out = pd.DataFrame({
            "Item ID": item_id,
            "Height": height,
            "Width": width,
            "Length": length,
            "Volume": volume,
            "Status": status,
            "Optimal Cardboard Width": cardboard_width
        })

        # --- SUMMARY (OK vs No OK) ---
        total = len(df_out)
        ok_count = int((df_out["Status"] == "OK").sum())
        no_ok_count = int((df_out["Status"] == "No OK").sum())
        ok_pct = (ok_count / total * 100) if total else 0.0
        no_ok_pct = (no_ok_count / total * 100) if total else 0.0

        summary_df = pd.DataFrame(
            {
                "Metric": ["Total Items", "OK Count", "OK %", "No OK Count", "No OK %"],
                "Value": [total, ok_count, round(ok_pct, 2), no_ok_count, round(no_ok_pct, 2)],
            }
        )

        # Display
        st.success("Processing complete!")
        st.subheader("Summary")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total", total)
        c2.metric("OK", f"{ok_count} ({ok_pct:.2f}%)")
        c3.metric("No OK", f"{no_ok_count} ({no_ok_pct:.2f}%)")

        st.subheader("Output Preview")
        st.dataframe(df_out)

        # --- DOWNLOAD CSV (data + footer summary lines) ---
        csv_data = df_out.to_csv(index=False)
        csv_summary = (
            "\n\nSummary\n"
            f"Total Items,{total}\n"
            f"OK Count,{ok_count}\n"
            f"OK %, {ok_pct:.2f}\n"
            f"No OK Count,{no_ok_count}\n"
            f"No OK %, {no_ok_pct:.2f}\n"
        )
        csv_bytes = (csv_data + csv_summary).encode("utf-8")

        st.download_button(
            "Download Output CSV (with Summary)",
            data=csv_bytes,
            file_name="CMC_Genesys_Item_Size_Analysis_Output.csv",
            mime="text/csv",
        )

        # --- DOWNLOAD EXCEL (Results + Summary sheet) ---
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Results")
            summary_df.to_excel(writer, index=False, sheet_name="Summary")

        st.download_button(
            "Download Output Excel (Results + Summary)",
            data=excel_buffer.getvalue(),
            file_name="CMC_Genesys_Item_Size_Analysis_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
