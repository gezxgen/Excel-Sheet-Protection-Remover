from re import sub, IGNORECASE
from os import rename, path
from sys import exit
from zipfile import ZipFile


def main() -> None:
    # input getting & checking
    input_dir: str = input("Enter the URL of the Excel file (...\\my_excel.xlsx): ")
    base_dir, ending = path.splitext(input_dir)
    if ending != ".xlsx":
        exit("Wrong file type entered.")

    # change to zip
    extract_dir: str = ""
    zip_dir: str = base_dir + ".zip"
    rename(input_dir, zip_dir)

    # extract zip
    # create path to extracting files
    splits_dir = base_dir.split("\\")
    for i in range(len(splits_dir)-1):
        extract_dir += splits_dir[i]
    # extract into path
    with ZipFile(zip_dir, "r") as f:
        f.extractall(path=extract_dir)
    # create folder name of extracted files
    f_name: str = sub(r"^[a-z]:", "", extract_dir, flags=IGNORECASE)
    f_name = sub(r"\\", "", f_name, flags=IGNORECASE)
    sheets_dir: str = extract_dir + "\\" + f_name + "\\xl\\worksheets"
    print(f"The Excel sheets are located at: {sheets_dir}")


if __name__ == "__main__":
    main()
