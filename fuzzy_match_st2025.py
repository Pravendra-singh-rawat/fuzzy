# import streamlit as st
# import pandas as pd
# from fuzzywuzzy import fuzz
# from fuzzywuzzy import process

# # Streamlit app
# def main():
#     st.title("Fuzzy Matching App")
#     st.write("Upload your Excel file and select the columns for matching.")

#     # File upload
#     uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    
#     if uploaded_file:
#         # Load Excel file
#         sheet_names = pd.ExcelFile(uploaded_file).sheet_names
#         sheet1_name = st.selectbox("Select first sheet (Old)", sheet_names)
#         sheet2_name = st.selectbox("Select second sheet (New)", sheet_names)
        
#         # Read sheets into DataFrames
#         df_sheet1 = pd.read_excel(uploaded_file, sheet_name=sheet1_name)
#         df_sheet2 = pd.read_excel(uploaded_file, sheet_name=sheet2_name)
        
#         st.write("### Preview of Sheet1 (Old)")
#         st.dataframe(df_sheet1.head())
#         st.write("### Preview of Sheet2 (New)")
#         st.dataframe(df_sheet2.head())
        
#         # Select columns for matching
#         col1 = st.selectbox("Select column to match from Sheet1 (Old)", df_sheet1.columns)
#         col2 = st.selectbox("Select column to match from Sheet2 (New)", df_sheet2.columns)
        
#         # Select column to check for code existence
#         code_col1 = st.selectbox("Select Center Code column from Sheet1 (Old)", df_sheet1.columns)
#         code_col2 = st.selectbox("Select Center Code column from Sheet2 (New)", df_sheet2.columns)
        
#         # Perform fuzzy matching
#         def find_matches(row, df2, col2):
#             matches = process.extractOne(row[col1], df2[col2], scorer=fuzz.token_set_ratio)
#             if matches and matches[1] >= 80:  # Set threshold for fuzzy matching
#                 return matches[0]  # Return matched name
#             else:
#                 return None

#         if st.button("Match and Process"):
#             # Apply fuzzy matching
#             df_sheet1["Matched Center Name"] = df_sheet1.apply(find_matches, df2=df_sheet2, col2=col2, axis=1)
            
#             # Check if Center_Code exists in both sheets
#             df_sheet1["Code Exists in Sheet2"] = df_sheet1[code_col1].isin(df_sheet2[code_col2])
            
#             # Display the results
#             st.write("### Matched Results")
#             st.dataframe(df_sheet1)
            
#             # Download the output
#             output_file = "output.xlsx"
#             df_sheet1.to_excel(output_file, index=False)
#             st.success(f"Results saved to {output_file}")
            
#             # Provide download link
#             with open(output_file, "rb") as file:
#                 st.download_button("Download Matched Results", file, file_name=output_file)

# if __name__ == "__main__":
#     main()







import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from time import sleep

# Streamlit app
def main():
    # Set the app title and layout
    st.set_page_config(page_title="Fuzzy Matching Tool", layout="wide", initial_sidebar_state="expanded")
    
    # Add a navigation bar
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Go to", ["Home", "About"])

    if page == "Home":
        render_home_page()
    elif page == "About":
        render_about_page()

def render_home_page():
    # App title
    st.title("🔍 Fuzzy Matching Tool for Excel Sheets")
    st.subheader("Match data between two sheets using fuzzy logic.")
    st.markdown("---")

    # File upload
    st.sidebar.header("Step 1: Upload Your Excel File")
    uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file:
        # Load Excel file
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        st.sidebar.header("Step 2: Select Sheets")
        sheet1_name = st.sidebar.selectbox("Choose first sheet (Old)", sheet_names)
        sheet2_name = st.sidebar.selectbox("Choose second sheet (New)", sheet_names)

        # Read sheets into DataFrames
        df_sheet1 = pd.read_excel(uploaded_file, sheet_name=sheet1_name)
        df_sheet2 = pd.read_excel(uploaded_file, sheet_name=sheet2_name)

        st.sidebar.header("Step 3: Column Selection")
        col1 = st.sidebar.selectbox("Select column to match from Sheet1 (Old)", df_sheet1.columns)
        col2 = st.sidebar.selectbox("Select column to match from Sheet2 (New)", df_sheet2.columns)
        code_col1 = st.sidebar.selectbox("Select Center Code column from Sheet1 (Old)", df_sheet1.columns)
        code_col2 = st.sidebar.selectbox("Select Center Code column from Sheet2 (New)", df_sheet2.columns)

        # Display previews with tabs
        st.header("📊 Data Previews")
        tabs = st.tabs(["Sheet1 (Old)", "Sheet2 (New)"])
        with tabs[0]:
            st.subheader(f"Preview of Sheet1 ({sheet1_name})")
            st.dataframe(df_sheet1.head(10))
        with tabs[1]:
            st.subheader(f"Preview of Sheet2 ({sheet2_name})")
            st.dataframe(df_sheet2.head(10))

        # Perform fuzzy matching
        def find_matches(row, df2, col2):
            matches = process.extractOne(row[col1], df2[col2], scorer=fuzz.token_set_ratio)
            if matches and matches[1] >= 80:  # Set threshold for fuzzy matching
                return matches[0]  # Return matched name
            else:
                return None

        # Button to start processing
        if st.button("🔄 Match and Process"):
            with st.spinner("Processing... Please wait."):
                progress_bar = st.progress(0)

                # Perform fuzzy matching and track progress
                results = []
                for i, row in df_sheet1.iterrows():
                    match = find_matches(row, df_sheet2, col2)
                    results.append(match)
                    progress_bar.progress((i + 1) / len(df_sheet1))
                    sleep(0.05)  # Simulate processing time

                # Add results to the dataframe
                df_sheet1["Matched Center Name"] = results
                df_sheet1["Code Exists in Sheet2"] = df_sheet1[code_col1].isin(df_sheet2[code_col2])

                st.success("✅ Matching completed successfully!")
                
                # Display results in an expander
                with st.expander("📋 Matched Results"):
                    st.dataframe(df_sheet1)
                
                # Save and download results
                output_file = "fuzzy_matching_output.xlsx"
                df_sheet1.to_excel(output_file, index=False)
                
                with open(output_file, "rb") as file:
                    st.download_button("📥 Download Matched Results", file, file_name=output_file)

    else:
        st.warning("⚠️ Please upload an Excel file to get started.")

    # Footer
    st.markdown("---")
    st.markdown("Built with ❤️ by Pravendra Singh")

def render_about_page():
    st.title("ℹ️ About the Fuzzy Matching Tool")
    st.markdown("""
    ## What does this tool do?
    This tool helps you match data between two sheets using fuzzy logic. It's especially useful for cases where names or other identifiers in two datasets are similar but not identical.

    ## How to use the tool
    1. **Upload your Excel file**: Ensure the file contains at least two sheets with the data you want to compare.
    2. **Select sheets and columns**:
       - Choose the sheets containing the data to be compared.
       - Select the column from each sheet that you want to match.
    3. **Match and Process**: Click the button to perform the fuzzy matching. The tool will:
       - Match names between the two columns based on similarity.
       - Check if codes from one sheet exist in the other.
    4. **Download results**: Once the process is complete, download the results as an Excel file.

    ## Features
    - **Fuzzy Matching**: Matches similar names with a customizable threshold.
    - **Code Check**: Verifies if codes from one sheet exist in the other.
    - **User-Friendly Interface**: Easy-to-follow steps and progress indicators.
    - **Export Results**: Download the output in Excel format.

    ## Who can use this tool?
    This tool is ideal for professionals handling messy data, such as:
    - Data analysts
    - Database administrators
    - Anyone working with Excel files and datasets
    """)
    st.markdown("---")
    st.markdown("Have questions or suggestions? Reach out to [Your Name](https://your-portfolio-link.com)")

if __name__ == "__main__":
    main()
