'''
Created on 10. maj 2024

@author: rn
'''


import streamlit as st
import pandas as pd
from P1Datahandling import startP1Datahandling
from io import BytesIO
#streamlit run C:\Users\rn\git\QC-tools\sourcecode\stream.py

def main(testFile):
    st.title("P1 Datahandling")

    pathPic = st.text_input("Enter the path for the picture storage", value=r"\\192.168.145.20\QC-billeder-P1")

    # Boolean input to delete default picture (not implemented as a real function here)
    delete_def_picture = st.checkbox("Delete Default Picture", value=False)

    # Numeric inputs for bounds
    low_bound = st.number_input("Low Bound", value=300, min_value=0)
    high_bound = st.number_input("High Bound", value=700, min_value=0)

    if testFile is None:
        # File uploader allows user to add a file
        uploaded_file = st.file_uploader("Choose a CSV file", type="xlsx")
    else:
        uploaded_file = testFile
    
    if uploaded_file is not None:
        # Read the CSV data into a dataframe
        data = pd.read_excel(uploaded_file)
        
        output = BytesIO()
        
        writer = pd.ExcelWriter(output, engine='openpyxl')
        
        df_results, writer, fig1, fig2 = startP1Datahandling(data, pathPic ,writer)
        
        writer.close()
        
        # Show the data as a table (you could also use st.write())
        st.dataframe(df_results)
        
        st.pyplot(fig1)
        #st.pyplot(fig2)
        
        # Rewind to beginning of stream
        output.seek(0)
        
        name = data["Roll ID"][0]
        # Provide download link for the excel
        st.download_button(
            label="Download Excel file",
            data=output,
            file_name=f"{name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Show statistics on the data
        #st.write("Basic Statistics:")
        #st.write(data.describe())

if __name__ == "__main__":
    #testFile = "C:/Users/rn/git/QC-tools/p1_view.xlsx"
    testFile = None
    main(testFile)