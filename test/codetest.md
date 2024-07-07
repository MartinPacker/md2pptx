template: Martin Template.pptx
contentsplit: 1 2
contentsplitdirn: h
style.fontsize.christopher: 45px
style.fgcolor.christopher: FF0000
hidden: yes

<style>
.christopher{
  font-size: 45px
}
</style>

# Code Test

### Here Is A Slide With A Graph

``` run-python

# Read chart data from CSV file
chart_csv = RunPython.readCSV("chartdata.csv")

# Make chart data from the array. Second parameter defaults to True for "Series Is Column"
chart_data = RunPython.makeChartData(chart_csv, True)

chart1 = RunPython.makeChart(slide,
  XL_CHART_TYPE.COLUMN_CLUSTERED,
  renderingRectangle,
  chart_data,
  "My Important Chart",
  XL_LEGEND_POSITION.BOTTOM)        
```

## Here Is  A Table

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