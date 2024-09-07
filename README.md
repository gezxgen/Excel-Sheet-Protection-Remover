## Removing Sheet Protection and File-Sharing from Excel Files

### Introduction

This Python script provides a graphical user interface (GUI) using CustomTkinter to remove sheet protection and file-sharing from Excel files. It allows users to input the path to an Excel file, processes the file, and creates a new file with the protection and file-sharing removed.

### Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Code Explanation](#code-explanation)
    - [Import Statements](#import-statements)
    - [GUI Class](#gui-class)
    - [Remove Protection Function](#remove-protection-function)
- [Summary](#summary)

### Installation

1. **Install required libraries:**
   ```bash
   pip install customtkinter
   ```

### Usage

1. **Run the script:**
   ```bash
   python remove_protection_gui.py
   ```
2. **Enter the path:** Input the path to the Excel file in the provided entry box.
3. **Click the button:** Click the "Remove Protection" button to start the process.

### Code Explanation

#### Import Statements

```python
import customtkinter as ctk
from os import path, remove, listdir, mkdir
from sys import exit
from zipfile import ZipFile
import re
```

- **customtkinter:** Used for creating the GUI.
- **os:** Provides functions for file and directory operations.
- **sys:** Used for exiting the program.
- **zipfile:** Used for working with ZIP files.
- **re:** Used for regular expressions to match and replace patterns.

#### GUI Class

```python
class RemoveSheetProtectionGUI(ctk.CTk):
    def __init__(self):
        # ... (rest of the GUI initialization code)
```

- **RemoveSheetProtectionGUI:** The main class for the GUI.
- **__init__(self):** Initializes the GUI components, including the label, entry box, and button.

#### Remove Protection Function

```python
def remove_protection(self):
    # ... (rest of the protection removal code)
```

- **remove_protection(self):** This function is called when the "Remove Protection" button is clicked. It performs the following steps:
    1. **Gets the input path:** Retrieves the path entered by the user.
    2. **Checks file type:** Verifies that the file has a `.xlsx` extension.
    3. **Extracts file name:** Extracts the filename from the path.
    4. **Creates new filename:** Creates a new filename with "_removed" appended.
    5. **Converts to ZIP:** Converts the Excel file to a ZIP file.
    6. **Extracts ZIP:** Extracts the ZIP file to a temporary directory.
    7. **Removes protection:** Removes sheet protection and file-sharing from the extracted files.
    8. **Creates new ZIP:** Creates a new ZIP file with the modified files.
    9. **Deletes temporary directory:** Removes the temporary directory.
    10. **Shows success message:** Displays a message indicating successful removal.

### Summary

This script provides a convenient GUI for removing sheet protection and file-sharing from Excel files. It automates the process, making it easier for users to manage their Excel files.
