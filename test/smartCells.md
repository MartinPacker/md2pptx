template: Martin Template.pptx

## Here Is A Table - With Smart Colouring

``` run-python

chart_csv = RunPython.readCSV("chartdata2.csv")

# Make the table with the data
table1 = RunPython.makeTable(slide,renderingRectangle, chart_csv)

# Cycle through the cells, formatting them
redList = []
greenList = []
orangeList = []

for rowNumber, row in enumerate(chart_csv):
  for columnNumber, cell in enumerate(row):
    try:
      isFloat = True
      floatValue = float(cell)
      if columnNumber == 2:
        if floatValue >= 99:
          redList.append((rowNumber, columnNumber))
        elif floatValue >= 95:
          orangeList.append((rowNumber, columnNumber))
        else:
          greenList.append((rowNumber, columnNumber))
    except ValueError:
      isFloat = False

    # Align last two columns right
    if columnNumber > 0:
      RunPython.alignTableCellText(table1, rowNumber, columnNumber, PP_ALIGN.RIGHT)

# Set appropriate list of cells to red
RunPython.applyCellListFillRGB(table1, redList, 255, 0, 0)

# Set appropriate list of cells to orange
RunPython.applyCellListFillRGB(table1, orangeList, 255, 255, 0)

# Set appropriate list of cells to green
RunPython.applyCellListFillRGB(table1, greenList, 0, 255, 0)


```
