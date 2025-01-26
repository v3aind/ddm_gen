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

        # Validate required columns
        required_columns_file2 = ["Keywords", "Shortcode", "Unreg", "Keyword Alias1", "Keyword Alias2", "Commercial Name", "SIM Action", "SIM Validity", "Package Validity", "Renewal", "PricePre"]
        for col in required_columns_file2:
            if col not in file2_df.columns:
                raise KeyError(f"Missing required column '{col}' in Product Spec Roaming.xlsx")
        
        for index, row in file2_df.iterrows():
            keyword = row["Keywords"]
        
            # Get PO ID from file1_df based on some criteria (e.g., matching keyword)
            matching_rows = file1_df.loc[file1_df['Keyword'] == keyword, 'POID']
        
            if not matching_rows.empty:
                po_id_from_file1 = matching_rows.iloc[0]
                output_file_name = f"{po_id_from_file1}.xlsx"
        
                # Create a Pandas ExcelWriter
                with pd.ExcelWriter(output_file_name, engine='xlsxwriter') as writer:
        
                    # Create the "PO-Master" sheet
                    po_master_data = {
                        "PO ID": [po_id_from_file1],
                        "Family": ["ROAMINGSINGLECOUNTRY"],
                        "Family Code": ["RSC"]
                    }
                    po_master_df = pd.DataFrame(po_master_data)
                    po_master_df.to_excel(writer, sheet_name="PO-Master", index=False)
        
                    # Create the "Keyword-Master" sheet
                    keyword_master_data = {
                        "Keyword": [
                            row["Keywords"],  # 1st row
                            row["Keywords"],  # 2nd row
                            row["Keywords"],  # 3rd row
                            "AKTIF_P26",      # 4th row
                            "AKTIF",          # 5th row
                            row["Unreg"]      # 6th row from file2 column "Unreg"
                        ],
                        "Short Code": [
                            str(int(row["Shortcode"])),  # 1st row from file2 without .0
                            "124",                        # 2nd row
                            "929",                        # 3rd row
                            "122",                        # 4th row
                            "122",                        # 5th row
                            "122"                         # 6th row
                        ],
                        "Keyword Type": [
                            "Master",          # 1st row
                            "Master",          # 2nd row
                            "Master",          # 3rd row
                            "Dormant",         # 4th row
                            "Dormant",         # 5th row
                            "UNREG"            # 6th row
                        ]
                    }
                    keyword_master_df = pd.DataFrame(keyword_master_data)
                    keyword_master_df.to_excel(writer, sheet_name="Keyword-Master", index=False)
        
                    # Create the "Keyword-Alias" sheet
                    keyword_alias_data = {
                        "Keyword": [
                            row["Keywords"],  # 1st row
                            row["Keywords"],  # 2nd row
                        ],
                        "Short Code": [
                            str(int(row["Shortcode"])),  # 1st row from file2 without .0
                            str(int(row["Shortcode"])),  # 2nd row without .0
                        ],
                        "Keyword Aliases": [
                            row["Keyword Alias1"],  # 1st row
                            row["Keyword Alias2"],  # 2nd row
                        ]
                    }
                    keyword_alias_df = pd.DataFrame(keyword_alias_data)
                    keyword_alias_df.to_excel(writer, sheet_name="Keyword-Alias", index=False)
        
                    # Create the "Ruleset-Header" sheet
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
                            row["Commercial Name"], 
                            row["Commercial Name"], 
                            row["Commercial Name"],
                            row["Commercial Name"]
                        ],
                        "Commercial Name English": [
                            row["Commercial Name"], 
                            row["Commercial Name"], 
                            row["Commercial Name"],
                            row["Commercial Name"]
                        ],
                        "Variant Type": ["00", "00", "00", "00"],
                        "SubVariant Type": ["PRE00", "ACT00", "ACT00", "0000"],
                        "SimCard Validity": [
                            row["SIM Action"], 
                            row["SIM Action"], 
                            row["SIM Action"],
                            row["SIM Action"]
                        ],
                        "LifeTime Validity": [
                            str(int(row["SIM Validity"])) if pd.notna(row["SIM Validity"]) else "",
                            str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                            str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                            str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else ""
                        ],
                        "MaxLife Time": ["360", "360", "360", "360"],
                        "UPCC Package Code": [
                            file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                            file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                            file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                            file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else ""
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
                            0,
                            0,
                            int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0
                        ],
                        "Reg Subaction": ["1", "1", "1", "1"]
                    }
        
                    ruleset_header_df = pd.DataFrame(ruleset_header_data)
                    ruleset_header_df.to_excel(writer, sheet_name="Ruleset-Header", index=False)
        
                    # Ensure MCC is treated as a string and split by commas
                    mcc_raw = str(row['MCC'])  # Convert MCC to string
                    mcc_values = mcc_raw.split(',')  # Split by commas
        
                    # Add 'm' prefix to each value and strip any surrounding whitespace
                    mcc_prefixed = ','.join([f"m{mcc.strip()}" for mcc in mcc_values])
        
                    # Split CC values, prefix each with 'c', and join them back with commas
                    cc_raw = str(row['Country Code'])  # Convert CC to string
                    cc_values = str(row['Country Code']).split(',')
                    cc_prefixed = ','.join([f"c{cc.strip()}" for cc in cc_values])
        
                    # Create DDM-Rule
                    ddm_rule_data ={
                        "Keyword": [row["Keywords"],row["Keywords"], "AKTIF_P26", "AKTIF", row["Keywords"], row["Keywords"]],
                        "Ruleset ShortName": [
                            f"{po_id_from_file1}:MRPRE00",
                            f"{po_id_from_file1}:MRPRE00",
                            f"{po_id_from_file1}:MRACT00",
                            f"{po_id_from_file1}:MRACT00",
                            f"{po_id_from_file1}:MR0000",
                            f"{po_id_from_file1}:MR0000"
                        ],
                        "ACTIVE_SUBS": [""] * 6,
                        "OpIndex":[3,4,1,1,1,2],
                        "SALES_AREA": [""] * 6,
                        "ZONE": [""] * 6,
                        "ORIGIN": [
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}",
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}",
                            "SDP",
                            "SDP",
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}",
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}"
                        ],
                        "RSC_ChildPO": [
                            "PO_ADO_DOR_AKTIF_P26", "PO_ADO_DOR_AKTIF_P26", "", "", "",""
                        ],
                        "RSC_LOCATION": ["DEFAULT", "DEFAULT", "", "", "DEFAULT", "DEFAULT"],
                        "RSC_DEFAULT_SALES_AREA": [""] * 6,
                        "SUBSCRIBER_TYPE": ["PREPAID,POSTPAID"] * 6,
                        "SM_REGION": [""] * 6,
                        "RSC_MAXMPP": [""] * 6,
                        "RSC_RESERVE_BALANCE": [""] * 6,
                        "DA_204": [""] * 6,
                        "UA_165": [""] * 6,
                        "ORDERTYPE": ["REGISTRATION"] * 6,
                        "GIFT": ["FALSE","FALSE","","","FALSE","FALSE"],
                        "RSC_CommercialName": [row["Commercial Name"]] * 6,
                        "ROAMING": [
                            "",
                            "",
                            f"IN|{mcc_prefixed},{cc_prefixed},{row['MCC_hex']}",
                            f"IN|m{row['MCC']},c{row['Country Code']},{row['MCC_hex']}",
                            f"IN|{row['MCC_hex']}",
                            f"IN|{row['MCC_hex']}"
                        ],
                        "ROAMINGFLAG": ["EQ|TRUE", "", "", "", "EQ|TRUE", ""],
                        "RSC_serviceKeyword": ["", "ActivateIntlRoaming", "", "", "", "ActivateIntlRoaming"],
                        "RSC_serviceName": ["", "ActivateIntlRoaming", "", "", "", "ActivateIntlRoaming"],
                        "RSC_serviceProvider": ["", "ICARE", "", "", "", "ICARE"],
                        "RSC_BYP_CONSENT_CHANNEL" : [
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}",
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}",
                            "",
                            "",
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}",
                            f"{row['Channel-SS']},{row['Channel-Trad-NonTrad']}"
                        ],
                        "RSC_RuleSetName": [
                            "GLOBAL_ELIG_ROAMING_PREACT1",
                            "GLOBAL_ELIG_ROAMING_PREACT1",
                            "GLOBAL_ELIG_ROAMING_PREACT2",
                            "GLOBAL_ELIG_ROAMING_PREACT2",
                            "GLOBAL_ELIG_ROAMING_NORMAL",
                            "GLOBAL_ELIG_ROAMING_NORMAL"],
                        "PREACT_SUBS": [
                            "",
                            "",
                            f"IN|{po_id_from_file1}:MRPRE00",
                            f"IN|{po_id_from_file1}:MRPRE00",
                            "",
                            ""
                        ]
                    }
        
                    ddm_rule_df = pd.DataFrame(ddm_rule_data)
                    ddm_rule_df.to_excel(writer, sheet_name="DDM-Rule", index=False)
        
                    # Create Rules-Price
                    rules_price_data ={
                       "Ruleset ShortName": [
                            f"{po_id_from_file1}:MRPRE00",
                            f"{po_id_from_file1}:MRPRE00",
                            f"{po_id_from_file1}:MRACT00",
                            f"{po_id_from_file1}:MRACT00",
                            f"{po_id_from_file1}:MR0000",
                            f"{po_id_from_file1}:MR0000"
                        ],
                        "Variable Name": ["REGISTRATION"] * 3 + ["DORMANT"] + ["REGISTRATION"] * 2,
                        "Channel":[
                            row["Channel Free"],
                            "DEFAULT",
                            "DEFAULT",
                            f"{po_id_from_file1}:MRPRE00",
                            row["Channel Free"],
                            "DEFAULT"
                        ],
                        "Price": [
                            0,
                            int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                            0,
                            "",
                            0,
                            int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                        ],
                        "SID": [
                            "12200001178102", 
                            "12200001178102", 
                            "12200001178102", 
                            "",
                            "12200001178102", 
                            "12200001178102" 
                        ]
                    }
        
                    rules_price_df = pd.DataFrame(rules_price_data)
                    rules_price_df.to_excel(writer, sheet_name="Rules-Price", index=False)

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
