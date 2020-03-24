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

## Todo:

### Allow reading from two separate excel sheets.

### Optimize the matching.
Right now it's N^2 * 2. Everything in sheet A is compared to sheet B and vice versa. 

Practically speaking this doesn't matter for anything under 300 row sheets on a 2016 laptop. If you had two excel sheets with 5000 rows then it may be a bit slow. 

We can maybe keep a log of things that matched and remove them as they are matched but the issue is that one row may match multiple rows and it's not clear exactly which one is the "real" one.

For example, comparing two sheets of checks based on the amount, 3 $20 checks in sheet A may match with 5 $20 checks in sheet B. You can't simply pull them out because it's not clear which one matches which.

We could include a secondary row to compare on maybe using regex or some kind of match % (like say check description) but this is niche and may not be used and we're now entering into scope creep for a feature that isn't always useful.
