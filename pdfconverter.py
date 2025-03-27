import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog
import os

crop_area = (50, 100, 580, 750)  # Example coordinates

# Function to select multiple PDF files
def choose_pdf_files():
    Tk().withdraw()  # Hide the root Tkinter window
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    return file_paths

# Function to ask for output format (CSV or XLSX)
def choose_output_format():
    while True:
        output_format = input("Choose output format (csv/xlsx): ").strip().lower()
        if output_format in ['csv', 'xlsx']:
            return output_format
        else:
            print("Invalid option. Please choose either 'csv' or 'xlsx'.")

def extract_tables_from_cropped_area(pdf, crop_area):
    all_tables = []
    for i, page in enumerate(pdf.pages, start=1):
        # Crop the page to the specified area
        cropped_page = page.crop(crop_area)
        print(f"Extracting tables from cropped page {i}...")

        tables = cropped_page.extract_tables()
        if tables:
            for table in tables:
                if table and len(table) > 1:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)
        else:
            print(f"No tables found on cropped page {i}")
    return all_tables


# # Function to extract all tables from a PDF page using pdfplumber
# def extract_tables_from_page(pdf):
#     all_tables = []
#     for i, page in enumerate(pdf.pages, start=1):
#         print(f"Extracting tables from page {i}...")  # Debugging output

#         # Adjust table extraction strategy
#         table_settings = {
#             "vertical_strategy": "text",  # Using text-based vertical strategy
#             "horizontal_strategy": "text",  # Using text-based horizontal strategy
#             "intersection_tolerance": 5,  # Allow some tolerance for text intersections
#         }

#         tables = page.extract_tables(table_settings=table_settings)
        
#         if tables:
#             for table in tables:
#                 if table and len(table) > 1:
#                     df = pd.DataFrame(table[1:], columns=table[0])
#                     all_tables.append(df)
#         else:
#             print(f"No tables found on page {i}")  # Debugging when no table is found on a page
            
#     return all_tables

# Function to convert PDF to CSV or XLSX
def convert_pdf_to_format(pdf_path, output_format):
    # Open the selected PDF
    with pdfplumber.open(pdf_path) as pdf:
#        all_tables = extract_tables_from_page(pdf)
        all_tables = extract_tables_from_cropped_area(pdf, crop_area)       
        if not all_tables:
            print(f"No valid tables found in {pdf_path}. Skipping file.")
            return
        
        # Reset index for each table to avoid duplicate indices
        final_df = pd.concat([table.reset_index(drop=True) for table in all_tables], ignore_index=True)

        # Get the output filename (same as the PDF file but with chosen format)
        output_filename = os.path.splitext(os.path.basename(pdf_path))[0]
        
        # Save the DataFrame in the chosen format
        if output_format == 'csv':
            final_df.to_csv(f"{output_filename}.csv", index=False)
            print(f"File saved as {output_filename}.csv")
        elif output_format == 'xlsx':
            final_df.to_excel(f"{output_filename}.xlsx", index=False, engine='openpyxl')
            print(f"File saved as {output_filename}.xlsx")


def main():
    # Step 1: Choose PDF files
    pdf_paths = choose_pdf_files()
    if not pdf_paths:
        print("No PDF files selected. Exiting.")
        return

    # Step 2: Choose output format (CSV or XLSX)
    output_format = choose_output_format()

    # Step 3: Convert each PDF and save as the chosen format
    for pdf_path in pdf_paths:
        print(f"Processing {pdf_path}...")
        convert_pdf_to_format(pdf_path, output_format)

if __name__ == "__main__":
    main()
