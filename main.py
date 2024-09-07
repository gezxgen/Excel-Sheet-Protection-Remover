from os import rename, path, remove
from sys import exit
from zipfile import ZipFile


def main() -> None:
    # input getting & checking
    input_dir: str = input("Enter the URL of the Excel file (...\\my_excel.xlsx): ")
    base_dir, ending = path.splitext(input_dir)
    if ending != ".xlsx":
        exit("Wrong file type entered.")

    # change to zip
    zip_dir: str = base_dir + ".zip"
    rename(input_dir, zip_dir)

    # extract zip
    # create path to extracting files
    splits_dir = base_dir.split("\\")
    extract_dir = path.join("\\".join(splits_dir[:-1]), "extracted")
    # extract into path
    with ZipFile(zip_dir, "r") as f:
        f.extractall(path=extract_dir)

    # delete old zip file
    # remove(extract_dir[:-9])

    sheets_dir: str = extract_dir + "\\xl\\worksheets"
    print(f"The Excel sheets are located at: {sheets_dir}")


if __name__ == "__main__":
    main()
