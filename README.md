# excel-sheet-combiner
Have an Excel workbook with many sheets that need to be combined into a single sheet? Use the GUI-enhanced "combine_sheets_gui.py" to automate the copying/pasting in just a few clicks.

## Dependencies
You'll need Python 3 to run the scripts. We rely on the 'pandas' library and its super-powerful Data Frames for all our manipulation.

```python
pip install pandas
```

## Packaged Executable
Alternatively, you can run the packaged executable in 'dist\' on a machine running Windows.

You can also use pyinstaller on a Mac or Linux machine to package the script for your own use.
```
pyinstaller --onefile combine_sheets_gui.py
```
