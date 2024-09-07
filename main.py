from os import path, remove, listdir, mkdir, rename, walk
from sys import exit
from zipfile import ZipFile
import re


def main() -> None:
    # Input getting & checking
    input_dir: str = input("Enter the URL of the Excel file (...\\my_excel.xlsx): ")
    base_dir, ending = path.splitext(input_dir)
    if ending != ".xlsx":
        exit("Wrong file type entered.")

    # Extract file name and create new compressed name
    filename = base_dir.split("\\")[-1]
    new_filename = filename + "_removed.xlsx"

    # Change to zip
    zip_dir: str = base_dir + ".zip"
    rename(input_dir, zip_dir)

    # Extract zip
    # Create path to extracting files
    splits_dir = base_dir.split("\\")
    extract_dir = path.join("\\".join(splits_dir[:-1])) + "\\extracted"
    # Create extraction directory if it doesn't exist
    mkdir(extract_dir)  # Handle existing directory case

    # Extract into path
    with ZipFile(zip_dir, "r") as f:
        f.extractall(path=extract_dir)

    # Delete old zip file
    remove(zip_dir)

    # Remove protection in sheets
    sheets_dir = extract_dir + "\\xl\\worksheets\\"
    protection_pattern = re.compile(r'<sheetProtection[^>]*>', re.IGNORECASE)  # Simplified pattern

    for i in range(1, len(listdir(sheets_dir)) + 1):  # Loop through sheet files efficiently
        sheet_path = f"{sheets_dir}sheet{i}.xml"

        try:
            # Read the content of the sheet XML file
            with open(sheet_path, "r", encoding="utf-8") as f:
                content = f.read()

            # Remove <sheetProtection> tag and everything until the next >
            modified_content = protection_pattern.sub('', content)

            # Write the modified content back to the sheet XML file
            with open(sheet_path, "w", encoding="utf-8") as f:
                f.write(modified_content)

        except FileNotFoundError:
            pass  # No need to print anything here

    # Remove file-sharing from workbook.xml
    workbook_path = extract_dir + "\\workbook.xml"
    filesharing_pattern = re.compile(r'<fileSharing>', re.IGNORECASE)

    try:
        with open(workbook_path, "r", encoding="utf-8") as f:
            workbook_content = f.read()

        workbook_content = filesharing_pattern.sub('', workbook_content)

        with open(workbook_path, "w", encoding="utf-8") as f:
            f.write(workbook_content)

    except FileNotFoundError:
        pass  # Handle missing workbook.xml

    # Create a new compressed file
    with ZipFile(new_filename, "w") as zip_file:
        for root, _, files in walk(extract_dir):
            for file in files:
                zip_file.write(path.join(root, file), path.relpath(path.join(root, file), extract_dir))

    # Delete the temporary extracted directory
    remove(extract_dir)

    print("Protection and file-sharing removed")
    print("New file:", new_filename)  # Print the new filename


if __name__ == "__main__":
    main()
