"""
runPython
"""

version = "0.1"

import csv
from pptx.chart.data import CategoryChartData

class RunPython:
    def __init__(
        self,
    ):
        pass



    # Execute the lines of code passed in
    def run(self,slide, renderingRectangle, codeLines, codeType):
        concatenatedCodeLines = "\n".join(codeLines)
        exec(concatenatedCodeLines)


    # Helper function for run-python
    def readCSV(filename):
        my_csv = []
        with open(filename, 'r') as csvfile:
            chart_reader = csv.reader(csvfile, quoting = csv.QUOTE_NONNUMERIC)
            for row in chart_reader:
                my_csv.append(row)

        return my_csv

    def makeChartData(chart_csv):

        chart_data = CategoryChartData()

        chart_data.categories = chart_csv[0][1:]

        for rowNumber, row in enumerate(chart_csv[1:]):
            chart_data.add_series(row[0],row[1:])

        return chart_data
  
    # Helper function to make a chart. The result can be further manipulated
    def makeChart(slide,
        chart_type,
        renderingRectangle,
        chart_data,
        title = None,
        legendPosition = None):    
        chart = slide.shapes.add_chart(
            chart_type,
            renderingRectangle.left,
            renderingRectangle.top,
            renderingRectangle.width, 
            renderingRectangle.height,
            chart_data
          ).chart
        
        if title is not None:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
          
        if legendPosition is not None:
            chart.has_legend = True
            chart.legend.position = legendPosition
            chart.legend.include_in_layout = False


        return chart


