# Excel Matcher

This program will open excel and will compare one column from all of the rows between two sheets (A and B) and create two new sheets.

Each sheet will list every row from A with the rows from B under it that matched it.

You just need Matcher.py and config.ini to run this.

1) Place excel into the same directory as Matcher.py
2) Look over and edit config.ini as needed
3) Run Matcher.py (Alternatively do it line by line in Jupyter Notebook via Matcher.ipynb)

## Tips:

It's better to just set column to 0 so it'll compare on the first column. Move the columns that are important (in Excel or Libre Office) to the first 1-3 columns before running this script. It will make the output a lot easier to read.

## How output looks:

Row from Sheet A = XXXX

Matching Rows from Sheet B = YYYY

```
X X X X X
Y Y Y Y Y
Y Y Y Y Y


X X X X X
Y Y Y Y Y


X X X X X
Y Y Y Y Y
Y Y Y Y Y
Y Y Y Y Y
```
