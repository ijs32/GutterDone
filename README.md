# Gutter Done
![image](src/logo.png)

To run: Use Windows 10/11 (Confirmed working on Windows 11, probably needs adjustments for Windows 10), full-screen Hydraflow Express on a 1920x1080 primary monitor, minimize all other open windows including Civil3D. Don’t touch the mouse during run.  

Because of Windows 11's odd print screen behavior, ensure you only have one monitor connected to your computer, otherwise the print screen will appear on your secondary monitor and GutterDone will not see it.

## STEPS SUMMARIZED:
1. Make sure you have a single monitor turned on.
2. Minimize all other windows.
3. Have all your defaults inputted
    - Location: On-Grade
    - Local Depression: 6.00
    - Slope (Sw): 0.092
    - Slope (Sx): 0.030
    - Width: 0.67
    - n-value: 0.013
    - Compute by: Known Q
4. Use default config if Excel hasn't changed, otherwise choose your updated config
5. Choose the Excel.
6. Specify the start row, this is the first non-header row containing data, your first inlet.
7. Leave mouse alone while running.

If GutterDone fails, it will still output a failed Excel, you can restart the program with this failed excel and specify the row it left off at.

#### Command to create .exe
```
pyinstaller --onefile --windowed --icon=src/logo.ico --name="GutterDone" --add-data "src/logo.png;." --add-data "src/config.json;." --add-data "src/images;images" src/gutter_done.py
```