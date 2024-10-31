# FindReplaceEverywhere Macro

A VBA macro for Microsoft Word that performs find and replace operations across all Word documents in a specified folder.

## Features

- Works on both Windows and Mac
- Processes both .doc and .docx files
- Handles errors gracefully
- Shows operation progress
- Provides detailed completion summary

## Installation

1. Open Microsoft Word
2. Press Alt+F11 (Windows) or Option+F11 (Mac) to open the VBA editor
3. Right-click on your project in the Project Explorer
4. Select Import File
5. Navigate to `FindReplaceEverywhere.bas` and import it

## Usage

1. Open any Word document
2. Run the macro (View > Macros > FindReplaceEverywhere)
3. Enter the text to find
4. Enter the replacement text
5. Enter the folder path (or leave blank to use current document's folder)
6. Confirm the operation

## Notes

- Always backup your documents before running the macro
- The macro will skip temporary files (starting with ~$)
- Files that can't be accessed will be skipped with a warning
- A summary report will show the number of replacements made

## Limitations

- Cannot process password-protected documents
- May require permissions to access network drives
- Mac users need to use forward slashes (/) in paths
