"""
runPython
"""

myVersion = "0.1"

__version__ = myVersion

import csv

class RunPython:
  def __init__(
    self,
  ):
        pass

  
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

