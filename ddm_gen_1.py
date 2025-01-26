import pandas as pd
import streamlit as st
import io

# Streamlit interface for file uploads
st.title('Excel File Processing App')

file1 = st.file_uploader("Upload Roaming_SC_Completion.xlsx", type=["xlsx"])
file2 = st.file_uploader("Upload Product Spec Roaming.xlsx", type=["xlsx"])

# Function to process the uploaded files and provide download link
def process_files(file1, file2):
    if file1 is not None and file2 is not None:
        # Load input files
        file1_df = pd.read_excel(file1)
        file2_df = pd.read_excel(file2)

        # Validate required columns
        required_columns_file2 = ["Keywords", "Shortcode", "Unreg", "Keyword Alias1", "Keyword Alias2", "Commercial Name", "SIM Action", "SIM Validity", "Package Validity", "Renewal", "PricePre"]
        for col in required_columns_file2:
            if col not in file2_df.columns:
                st.error(f"Missing required column '{col}' in Product Spec Roaming.xlsx")
                return

        output_file_name = None  # Initialize variable for output file name
        
        for index, row in file2_df.iterrows():
            keyword = row["Keywords"]

            # Get PO ID from file1_df based on some criteria (e.g., matching keyword)
            matching_rows = file1_df.loc[file1_df['Keyword'] == keyword, 'POID']

            if not matching_rows.empty:
                po_id_from_file1 = matching_rows.iloc[0]
                output_file_name = f"{po_id_from_file1}.xlsx"

                # Create a Pandas ExcelWriter
                with io.BytesIO() as output:
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # Create "PO-Master" sheet
                        po_master_data = {"PO ID": [po_id_from_file1], "Family": ["ROAMINGSINGLECOUNTRY"], "Family Code": ["RSC"]}
                        po_master_df = pd.DataFrame(po_master_data)
                        po_master_df.to_excel(writer, sheet_name="PO-Master", index=False)

                        # Create "Keyword-Master" sheet
                        keyword_master_data = {
                            "Keyword": [row["Keywords"], row["Keywords"], row["Keywords"], "AKTIF_P26", "AKTIF", row["Unreg"]],
                            "Short Code": [str(int(row["Shortcode"])), "124", "929", "122", "122", "122"],
                            "Keyword Type": ["Master", "Master", "Master", "Dormant", "Dormant", "UNREG"]
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

                        # Create Rules-Renewal
                        rules_renewal_data= {
                            "Ruleset ShortName": [
                                f"{po_id_from_file1}:MRPRE00",
                                f"{po_id_from_file1}:MRACT00",
                                f"{po_id_from_file1}:MR0000"
                            ],
                            "PO ID": [po_id_from_file1] * 3,
                            "Flag Auto": [
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP"
                            ],
                            "Period": [
                                int(row["Dorman"]),
                                int(row["Package Validity"]),
                                int(row["Package Validity"])
                            ],
                            "Period UOM": ["DAY"] * 3,
                            "Flag Charge": ["FALSE"] * 3,
                            "Flag Suspend": ["FALSE"] * 3,
                            "Suspend Period": [""] *3,
                            "Suspend UOM": [""] * 3,
                            "Flag Option": ["FALSE"] * 3,
                            "Max Cycle": [1] *3,
                            "Progression Renewal": [""] * 3,
                            "Reminder Group Id": ["GROUP18"] * 3,
                            "Amount": [""] *3,
                            "Reg Subaction": [str(1)] * 3,
                            "Action Failure": ["DEFAULT"] * 3
                        }

                        rules_renewal_df= pd.DataFrame(rules_renewal_data)
                        rules_renewal_df.to_excel(writer, sheet_name="Rules-Renewal", index=False)

                        # Create Case-Type
                        case_type_data= {
                            "RulesetName": [
                                f"{po_id_from_file1}:MRPRE00",
                                f"{po_id_from_file1}:MRACT00",
                                f"{po_id_from_file1}:MR0000"
                            ],
                            "Case_Type": ["REGISTRATION,UNREG"] * 3
                        }

                        case_type_df=pd.DataFrame(case_type_data)
                        case_type_df.to_excel(writer, sheet_name="Case-Type", index=False)

                    # Move the file pointer to the beginning of the file so it can be downloaded
                    output.seek(0)

                    # Provide a download button for the user
                    st.download_button(
                        label=f"Download {output_file_name}",
                        data=output,
                        file_name=output_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.success(f"Output file '{output_file_name}' created successfully for keyword: {keyword}")
            else:
                st.warning(f"No matching POID found in file1_df for keyword: {keyword}")
    else:
        st.warning("Please upload both files to proceed.")

# Call the process function if both files are uploaded
if file1 is not None and file2 is not None:
    process_files(file1, file2)
