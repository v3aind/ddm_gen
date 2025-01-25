import pandas as pd
import streamlit as st
from io import BytesIO

# Streamlit UI for the app
st.title("Excel File Generator for Roaming Data")
st.markdown("Upload the input files and generate the required Excel outputs.")

# Upload files
file1 = st.file_uploader("Upload Roaming_SC_Completion.xlsx", type=["xlsx"])
file2 = st.file_uploader("Upload Product Spec Roaming.xlsx", type=["xlsx"])

if file1 and file2:
    try:
        # Load the input files
        file1_df = pd.read_excel(file1)
        file2_df = pd.read_excel(file2)

        # Validate required columns in file2
        required_columns_file2 = [
            "Keywords", "Shortcode", "Unreg", "Keyword Alias1", "Keyword Alias2",
            "Commercial Name", "SIM Action", "SIM Validity", "Package Validity", "Renewal", "PricePre"
        ]
        missing_columns = [col for col in required_columns_file2 if col not in file2_df.columns]
        if missing_columns:
            st.error(f"Missing required columns in Product Spec Roaming.xlsx: {', '.join(missing_columns)}")
        else:
            # Process each row in file2
            for index, row in file2_df.iterrows():
                keyword = row["Keywords"]

                # Match PO ID from file1 based on the keyword
                matching_rows = file1_df.loc[file1_df['Keyword'] == keyword, 'POID']
                if not matching_rows.empty:
                    po_id_from_file1 = matching_rows.iloc[0]
                    output_file_name = f"{po_id_from_file1}.xlsx"

                    # Create ExcelWriter in memory
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # "PO-Master" sheet
                        po_master_data = {
                            "PO ID": [po_id_from_file1],
                            "Family": ["ROAMINGSINGLECOUNTRY"],
                            "Family Code": ["RSC"]
                        }
                        po_master_df = pd.DataFrame(po_master_data)
                        po_master_df.to_excel(writer, sheet_name="PO-Master", index=False)

                        # "Keyword-Master" sheet
                        keyword_master_data = {
                            "Keyword": [
                                row["Keywords"], row["Keywords"], row["Keywords"],
                                "AKTIF_P26", "AKTIF", row["Unreg"]
                            ],
                            "Short Code": [
                                str(int(row["Shortcode"])), "124", "929", "122", "122", "122"
                            ],
                            "Keyword Type": [
                                "Master", "Master", "Master", "Dormant", "Dormant", "UNREG"
                            ]
                        }
                        keyword_master_df = pd.DataFrame(keyword_master_data)
                        keyword_master_df.to_excel(writer, sheet_name="Keyword-Master", index=False)

                        # "Keyword-Alias" sheet
                        keyword_alias_data = {
                            "Keyword": [row["Keywords"], row["Keywords"]],
                            "Short Code": [
                                str(int(row["Shortcode"])), str(int(row["Shortcode"]))
                            ],
                            "Keyword Aliases": [row["Keyword Alias1"], row["Keyword Alias2"]]
                        }
                        keyword_alias_df = pd.DataFrame(keyword_alias_data)
                        keyword_alias_df.to_excel(writer, sheet_name="Keyword-Alias", index=False)

                        # "Ruleset-Header" sheet
                        ruleset_header_data = {
                            "Ruleset ShortName": [
                                f"{po_id_from_file1}:MRPRE00",
                                f"{po_id_from_file1}:MRACT00",
                                f"{po_id_from_file1}:MRACT00",
                                f"{po_id_from_file1}:MR0000"
                            ],
                            "Keyword": [row["Keywords"], "AKTIF_P26", "AKTIF", row["Keywords"]],
                            "Keyword Type": ["", "", "", ""],
                            "Commercial Name Bahasa": [
                                row["Commercial Name"], row["Commercial Name"],
                                row["Commercial Name"], row["Commercial Name"]
                            ],
                            "Commercial Name English": [
                                row["Commercial Name"], row["Commercial Name"],
                                row["Commercial Name"], row["Commercial Name"]
                            ],
                            "Variant Type": ["00", "00", "00", "00"],
                            "SubVariant Type": ["PRE00", "ACT00", "ACT00", "0000"],
                            "SimCard Validity": [
                                row["SIM Action"], row["SIM Action"],
                                row["SIM Action"], row["SIM Action"]
                            ],
                            "LifeTime Validity": [
                                str(int(row["SIM Validity"])) if pd.notna(row["SIM Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else ""
                            ],
                            "MaxLife Time": ["360", "360", "360", "360"],
                            "UPCC Package Code": [
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0]
                                if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0]
                                if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0]
                                if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0]
                                if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else ""
                            ],
                            "Claim Command": ["", "", "", ""],
                            "Flag Auto": [
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP"
                            ],
                            "Progression Renewal": ["", "", "", ""],
                            "Reminder Group Id": ["GROUP18", "GROUP18", "GROUP18", "GROUP18"],
                            "Amount": [
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                                0, 0,
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0
                            ],
                            "Reg Subaction": ["1", "1", "1", "1"]
                        }
                        ruleset_header_df = pd.DataFrame(ruleset_header_data)
                        ruleset_header_df.to_excel(writer, sheet_name="Ruleset-Header", index=False)

                    st.success(f"Generated file for keyword: {keyword}")

                    # Provide download button
                    st.download_button(
                        label=f"Download {output_file_name}",
                        data=output.getvalue(),
                        file_name=output_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(f"No matching POID found for keyword: {keyword}")

    except Exception as e:
        st.error(f"An error occurred: {e}")
