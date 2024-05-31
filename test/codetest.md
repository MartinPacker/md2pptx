template: Martin Template.pptx
contentsplit: 1 2
contentsplitdirn: h

# Code Test

### I Ran Inline Code To Get This

* Here is a bullet
  * Here is a sub-bullet

``` run-python
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION

# Read chart data from CSV file
chart_csv = RunPython.readCSV("chartdata.csv")

chart_data = RunPython.makeChartData(chart_csv)

chart = RunPython.makeChart(slide,
  XL_CHART_TYPE.COLUMN_CLUSTERED,
  renderingRectangle,
  chart_data,
  "Hello World",
  XL_LEGEND_POSITION.RIGHT)


```