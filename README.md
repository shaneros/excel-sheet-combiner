# excel-sheet-combiner
Have an Excel workbook with many sheets that need to be combined into a single sheet? Use the GUI-enhanced `combine_sheets_gui.py` to automate the copying/pasting in just a few clicks.

![Image of GUI](/images/sheet-combiner-GUI.png)

Choose your Excel Workbook, click Combine Sheets, and you're done!

## Dependencies
You'll need Python 3 to run the scripts. We rely on the `pandas` library and its super-powerful Data Frames for all our manipulation.

```bash
pip install pandas
```

# Running the script
## Command-Line Only (not recommended)
Edit `combine_sheets.py` directly and change `input.xlsx` to whatever workbook you wish to manipulate. Change `output.xlsx` to a filename of your choice. This script must be run in the same directory as your source file.
```bash
py combine_sheets.py
```

## GUI-enabled
Run  `combine_sheets_gui.py` and follow the prompts in the window.
```bash
py combine_sheets_gui.py
```

## Packaged Executable
Alternatively, you can run the packaged executable in 'dist\' on a machine running Windows.

You can also use pyinstaller ([pyinstaller.org](https://www.pyinstaller.org/)) on a Mac or Linux machine to package the script for your own use.
```
pyinstaller --onefile combine_sheets_gui.py
```
