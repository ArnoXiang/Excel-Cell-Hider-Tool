## README: Excel Cell Hider Tool

### Project Overview

This tool is a simple, standalone Python application designed to help users quickly prepare Excel files for processes like translation or external review. It allows users to easily clear (hide) specified cells in an Excel sheet while perfectly preserving all other data, formatting, formulas, and visual styles.

### Problem Solved

When preparing an Excel file (e.g., a questionnaire, survey, or data template) for translation, you often need to **hide specific cells** (like internal reference IDs, original source text, or confidential data) so that only the translatable content remains visible to the external party.

**This tool provides a user-friendly solution by:**

  * **Precise Targeting:** Clears *only* the specific cells you input (e.g., `B4`, `B5`), eliminating the risk of accidental data deletion.
  * **Format Preservation:** The tool uses `openpyxl` to load and save, ensuring all colors, fonts, column widths, and formulas remain intact.
  * **Safety First:** It always saves the modified content to a **new file** with the prefix `For Translation_` (e.g., `For Translation_OriginalFilename.xlsx`), leaving your original file untouched.

### Features

  * **Graphical User Interface (GUI):** Easy-to-use interface built with `tkinter`.
  * **File Selection:** Browse and select the input Excel file (`.xlsx`, `.xlsm`).
  * **Sheet Selection:** Specify the name of the worksheet to be processed.
  * **Custom Cell Input:** Input a list of cells to hide, separated by commas (e.g., `C3, D10, F2`).
  * **Non-Destructive Saving:** The output file is automatically named with the prefix `For Translation_`.

### Prerequisites

To run this tool, you need the following installed on your system:

1.  **Python 3:** (Generally version 3.6 or newer is recommended).

2.  **Required Libraries:**

    Open your terminal or command prompt and run the following command to install the necessary library:

    ```bash
    pip install openpyxl
    ```

    *(Note: `tkinter` is usually included with standard Python installations.)*

### How to Use

1.  **Save the Script:** Save the provided code as a Python file (e.g., `extract_demo.py`).

2.  **Run the Script:** Execute the file from your terminal:

    ```bash
    python extract_demo.py
    ```

3.  **Follow the Steps in the GUI:**

    | Step | Action |
    | :--- | :--- |
    | **Step 1** | Click **Browse** and select your original Excel file. |
    | **Step 2** | Enter the exact **Sheet Name** you need to modify (e.g., `Emotional Function (Recall)`). |
    | **Step 3** | Enter the cell addresses you want to clear, separated by commas (e.g., `B4, B5, C10`). |
    | **Step 4** | Click the green **Run & Save** button. |

4.  **Check Output:** A success message will appear, and the new file (e.g., `For Translation_MyFile.xlsx`) will be saved in the same directory as the original file. The specified cells will be empty, ready for translation.
