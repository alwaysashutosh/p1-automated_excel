import streamlit as st
import pandas as pd
from io import BytesIO
import base64

def generate_excel_download_link(df, summary_df):
    # Create a new Excel writer object
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write original data
        df.to_excel(writer, sheet_name='Data', index=False)
        
        # Write summary starting two rows below the data
        start_row = len(df) + 2
        worksheet = writer.sheets['Data']
        
        # Write summary header
        worksheet.cell(start_row, 2, 'Summary Statistics')
        
        # Write summary data
        for idx, row in summary_df.iterrows():
            worksheet.cell(start_row + idx + 1, 2, row['Metric'])
            worksheet.cell(start_row + idx + 1, 3, row['Value'])
    
    # Generate download link
    output.seek(0)
    excel_data = output.read()
    b64 = base64.b64encode(excel_data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="analyzed_data.xlsx">Download Analyzed Excel File</a>'

def main():
    st.title("Excel Data Analyzer")
    st.write("""
    ### Upload your Excel file with an 'Amount' column
    This app will calculate key statistics and provide a downloadable report.
    """)
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file)
            
            # Check for Amount column
            if 'Amount' not in df.columns:
                st.error("Error: 'Amount' column not found in the Excel file!")
                return
            
            # Show original data
            st.subheader("Original Data Preview")
            st.dataframe(df.head())
            
            # Calculate summary statistics
            summary = {
                'Metric': ['Sum', 'Average', 'Max', 'Min', 'Count'],
                'Value': [
                    df['Amount'].sum(),
                    df['Amount'].mean(),
                    df['Amount'].max(),
                    df['Amount'].min(),
                    df['Amount'].count()
                ]
            }
            summary_df = pd.DataFrame(summary)
            
            # Display summary statistics with formatting
            st.subheader("Summary Statistics")
            # Format the summary dataframe for display
            formatted_summary = summary_df.copy()
            formatted_summary['Value'] = formatted_summary['Value'].apply(
                lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x
            )
            st.dataframe(formatted_summary.set_index('Metric'))
            
            
            # Generate download link
            st.subheader("Download Analyzed Data")
            st.markdown(
                generate_excel_download_link(df, summary_df), 
                unsafe_allow_html=True
            )
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            
if __name__ == "__main__":
    main()