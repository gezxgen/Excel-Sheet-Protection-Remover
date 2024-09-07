from os import rename, path, remove, listdir, walk
from sys import exit
from zipfile import ZipFile
import re


def main() -> None:
    # Input getting & checking
    input_dir: str = input("Enter the URL of the Excel file (...\\my_excel.xlsx): ")
    base_dir, ending = path.splitext(input_dir)
    if ending != ".xlsx":
        exit("Wrong file type entered.")

    # Change to zip
    zip_dir: str = base_dir + ".zip"
    rename(input_dir, zip_dir)

    # Extract zip
    # Create path to extracting files
    splits_dir = base_dir.split("\\")
    extract_dir = path.join("\\".join(splits_dir[:-1]), "extracted")

    # Extract into path
    with ZipFile(zip_dir, "r") as f:
        f.extractall(path=extract_dir)

    # Remove old zip file
    remove(zip_dir)

    # Path to Excel sheet (.xml) files
    sheets_dir: str = path.join(extract_dir, "xl", "worksheets")
    workbook_file: str = path.join(extract_dir, "xl", "workbook.xml")

    # Define regex patterns to remove sheet and workbook protection
    sheet_protection_pattern = re.compile(r'<sheetProtection[^>]*>.*?</sheetProtection>', re.DOTALL)
    workbook_protection_pattern = re.compile(r'<workbookProtection[^>]*>.*?</workbookProtection>', re.DOTALL)

    # Remove protection from sheet XML files
    xml_files = [f for f in listdir(sheets_dir) if f.endswith(".xml")]
    for xml_file in xml_files:
        file_path = path.join(sheets_dir, xml_file)

        # Read the content of the sheet XML file
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Remove the <sheetProtection> tag using regex
        modified_content = re.sub(sheet_protection_pattern, '', content)

        # Write the modified content back to the sheet XML file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(modified_content)

    # Remove protection from workbook XML file
    if path.exists(workbook_file):
        with open(workbook_file, 'r', encoding='utf-8') as f:
            content = f.read()

        # Remove the <workbookProtection> tag using regex
        modified_content = re.sub(workbook_protection_pattern, '', content)

        # Write the modified content back to the workbook XML file
        with open(workbook_file, 'w', encoding='utf-8') as f:
            f.write(modified_content)

    # Rename the extracted folder with "_removed" to indicate the protection is removed
    new_excel_path = base_dir + "_removed.xlsx"

    # Create a new zip file from the modified folder
    new_zip_dir = base_dir + "_removed.zip"

    with ZipFile(new_zip_dir, 'w') as new_zip:
        # Walk through the extracted directory and add files back to the zip
        for foldername, subfolders, filenames in walk(extract_dir):
            for filename in filenames:
                # Construct the full file path
                file_path = path.join(foldername, filename)
                # Get the relative path to preserve folder structure
                relative_path = path.relpath(file_path, extract_dir)
                new_zip.write(file_path, relative_path)

    # Rename the zip back to .xlsx
    rename(new_zip_dir, new_excel_path)

    # Print success message
    print(f"Excel sheets and workbook protection removed successfully. New file saved as: {new_excel_path}")


if __name__ == "__main__":
    main()
