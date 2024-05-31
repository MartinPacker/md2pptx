"""
runPython
"""

version = "0.1"

import csv

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


