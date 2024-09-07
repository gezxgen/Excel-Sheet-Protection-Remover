import customtkinter as ctk
from os import path, remove, listdir, mkdir, walk, rename
from zipfile import ZipFile
import re


class RemoveSheetProtectionGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Remove Sheet Protection")
        self.geometry("400x200")

        # Create widgets
        self.label = ctk.CTkLabel(self, text="Enter the path to the Excel file:")
        self.label.pack(pady=10)

        self.entry = ctk.CTkEntry(self)
        self.entry.pack(pady=10)

        self.button = ctk.CTkButton(self, text="Remove Protection", command=self.remove_protection)
        self.button.pack(pady=10)

    def remove_protection(self):
        input_dir = self.entry.get()
        base_dir, ending = path.splitext(input_dir)

        if ending != ".xlsx":
            ctk.CTkMessageBox(title="Error", message="Wrong file type entered.").show()
            return

        # Extract file name and create new compressed name
        filename = base_dir.split("\\")[-1]
        new_filename = filename + "_removed.xlsx"

        # Change to zip
        zip_dir = base_dir + ".zip"
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

        ctk.CTkMessageBox(title="Success", message="Protection and file-sharing removed. New file: " + new_filename).show()


if __name__ == "__main__":
    app = RemoveSheetProtectionGUI()
    app.mainloop()
