# Word VBA Macros

This directory contains VBA (Visual Basic for Applications) macros for Microsoft Word automation.

## Available Macros

### FindReplaceEverywhere
Located in: `/find-replace`

A comprehensive find and replace macro that searches through all possible locations in a Word document including:
- Main document body
- Headers and footers
- Text boxes
- Shapes
- Tables
- Comments
- Footnotes and endnotes

#### Usage
1. Open the Visual Basic Editor in Word (Alt + F11)
2. Import the macro from the `find-replace` directory
3. Run the macro through the Macros dialog (Alt + F8) or assign it to a shortcut

## Contributing
To add new macros to this directory:
1. Create a new subdirectory with a descriptive name
2. Add your `.bas` or `.cls` files
3. Include clear documentation within the code
4. Update this README with details about your macro

## Requirements
- Microsoft Word 2010 or later
- VBA enabled in Word settings

## Security Note
Before using any macro, ensure you:
- Review the source code
- Only enable macros from trusted sources
- Save your work before running macros
