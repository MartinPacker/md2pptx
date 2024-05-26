template: Martin Template.pptx
contentsplit: 1 2
contentsplitdirn: h

# Code Test

### I Ran Inline Code To Get This

* Here is a bullet
  * Here is a sub-bullet

``` run-python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION

chart_data = CategoryChartData()

# Read chart data from CSV file
chart_csv = RunPython.readCSV("chartdata.csv")

chart_data.categories = chart_csv[0][1:]

for rowNumber, row in enumerate(chart_csv[1:]):
  chart_data.add_series(row[0],row[1:])

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    renderingRectangle.left,
    renderingRectangle.top,
    renderingRectangle.width, 
    renderingRectangle.height,
    chart_data
).chart

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False

chart.has_title = True
chart.chart_title.text_frame.text = "Hello World"
```