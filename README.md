# xlPandas

Read and write Excel xlsx using pandas/openpyxl without destroying formatting.

Sometimes you have a nicely formatted worksheet, but you'd like to work with it
using pandas, or perhaps you want to write data to an Excel template.

Pandas can read and write excel files using `xlrd`, but treats them like csvs. 
xlPandas uses `openpyxl` to access data while preserving template formatting.

## Install

`pip install xlpandas`

## Example

``` python

import xlPandas as xpd

# Read excel file
df = xpd.read_file('template.xlsx', skiprows=2)
print(df.columns)

# Access openpyxl worksheet
sheet = df.to_sheet()
sheet.cell(1,1) = 'title'

# From openpyxl worksheet
df = xpd.xlDataFrame(sheet)

# Write file
df.to_file('out.xlsx')

```

