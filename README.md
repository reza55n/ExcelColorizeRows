# ExcelColorizeRows

**Group your data visually!**

This script colorizes and thickens borders of Excel rows based on one or more columns.


## How to implement
1. Copy the code
2. Open your document
3. Press `Alt+F11` (`Fn+F11` on Mac) to open VBA Editor
4. Paste the code

Now you can save your document as `Excel Macro-Enabled Workbook (*.xlsm)` format to use it later.


## How to use
1. Sort your table based on intended columns
2. Determine columns in the script (See next title)
3. (Optional) Select entire sheet or table and from `Home` tab, click on `Borders` drop-down and select `All Borders`
4. Press `Alt+F8` (`Fn+F8` on Mac) to open Macro dialog box
5. Select `...Colorize` and click Run


## Configuration
You can make some configurations by pressing `Alt+F11` (`Fn+F11` on Mac) and editing respective part of the code:
- `HeaderRowsCount`: Number of rows at header to exclude from colorizing, default: `1`.
- **`Cols`: For example, write `Cols = [{5, 6}]` to colorize based on columns 5 and 6 and write `Cols = [{1}]` to colorize only based on first column.**
- `UseColor`, `UseBorder` (`True`/`False`): Restrict function of script, default: `True`.
- `fixed`, `random` (Integer, total less than 256): Configure colorizing.


## Example
![Colorize Sample](https://www.alvandsoft.com/cloud123/excel_colorize.png)  
(Sample data from [contextures.com](https://www.contextures.com/xlsampledata01.html))

In example above, rows are sorted based on columns 2 and 3 and then, colorizing took place on same columns.

Configuration: `Cols = [{2, 3}]`
