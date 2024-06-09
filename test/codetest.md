template: Martin Template.pptx
contentsplit: 1 2
contentsplitdirn: h

# Code Test

### Here Is  A Graph

* Here is a bullet
  * Here is a sub-bullet

``` run-python

# Read chart data from CSV file
chart_csv = RunPython.readCSV("chartdata.csv")

# Make chart data from the array. Second parameter defaults to True for "Series Is Column"
chart_data = RunPython.makeChartData(chart_csv, True)

chart1 = RunPython.makeChart(slide,
  XL_CHART_TYPE.COLUMN_CLUSTERED,
  renderingRectangle,
  chart_data,
  "Hello World",
  XL_LEGEND_POSITION.BOTTOM)


```

### Here Is  A Table
<!-- md2pptx: contentsplit: 2 1 -->
<!-- md2pptx: contentsplitdirn: v -->

``` run-python

# Read chart data from CSV file
chart_csv = RunPython.readCSV("chartdata.csv")

# Make the table with the data
table1 = RunPython.makeTable(slide,renderingRectangle, chart_csv)

# Set a cell background to yellow
RunPython.applyCellFillRGB(table1, 2, 3, 255, 255, 0)

# Set list of cells to green
greenList = [(0, 0), (2,1), (3,2)]
RunPython.applyCellListFillRGB(table1, greenList, 0, 255, 0)


```

* Here's a bullet below the table