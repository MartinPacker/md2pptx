"""
runPython
"""

version = "0.3"

import csv
from pptx.chart.data import CategoryChartData
from pptx.oxml.xmlchemy import OxmlElement, serialize_for_reading
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from colour import setColour, parseColour

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
    
    def filterRows(my_array, filterFunction):
        my_array2 = []
        for rowNumber, row in enumerate(my_array):
            if filterFunction(rowNumber, row):
                my_array2.append(row)
        
        return my_array2
        

    def transposeArray(chart_array):
        return list(map(list, zip(*chart_array)))
    
    def makeChartData(chart_array, seriesIsColumn = True, columns = None):

        chart_data = CategoryChartData()
        
        if columns is not None:
            chart_array2 = []
            for rowNumber, row in enumerate(chart_array):
                chart_row = []
                for column in columns:
                    chart_row.append(row[column])
                
                chart_array2.append(chart_row)
            chart_array = chart_array2
            print(pptx.enum.chart)


        if seriesIsColumn:
            # Transpose input data
            chart_array = RunPython.transposeArray(chart_array)
        
        # x values
        chart_data.categories = chart_array[0][1:]

        # Series
        for rowNumber, row in enumerate(chart_array[1:]):
            chart_data.add_series(row[0],row[1:])

        return chart_data
  
    # Helper function to make a chart. The result can be further manipulated
    def makeChart(slide,
        chart_type,
        renderingRectangle,
        chart_data,
        title = None,
        legendPosition = None):    
        c = slide.shapes.add_chart(
            chart_type,
            renderingRectangle.left,
            renderingRectangle.top,
            renderingRectangle.width, 
            renderingRectangle.height,
            chart_data
        )
          
        chart = c.chart
        
        if title is not None:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
          
        if legendPosition is not None:
            chart.has_legend = True
            chart.legend.position = legendPosition
            chart.legend.include_in_layout = False


        return c

    # Helper routine to make a table. The result can be further manipulated
    def makeTable(slide,
        renderingRectangle,
        table_array):

        height = len(table_array)
        width = len(table_array[0])

        t = slide.shapes.add_table(height,width,
            renderingRectangle.left,
            renderingRectangle.top,
            renderingRectangle.width,
            renderingRectangle.height)

        table = t.table
        
        for i in range(height):
            for j in range(width):
               c = table.cell(i, j)
               c.text = str(table_array[i][j])

        return t

    def applyCellFillRGB(table, row, column, red, green, blue):
        cff = table.table.cell(row, column).fill
        cff.solid()
        cff.fore_color.rgb = RGBColor(red, green, blue)

    def applyCellListFillRGB(table, cellList, red, green, blue):
        for row, column in cellList:
            RunPython.applyCellFillRGB(table, row, column, red, green, blue)
            
    def alignTableCellText(tableFrame, rowNumber, columnNumber, alignment, paragraphNumber = None):
        # Get the cell's text_frame
        tableCellFrame = tableFrame.table.cell(rowNumber, columnNumber).text_frame

        if paragraphNumber == None:
            # Iterate over the cell's paaragraph's, aligning right
            for p in tableCellFrame.paragraphs:
                p.alignment = alignment
        else:
            tableCellFrame.paragraphs[paragraphNumber].alignment = alignment

    def makeDrawnShape(slide, vertices, fill = False, text = None, textColor = None, fillColor = None, closed = True):
        ffBuilder = slide.shapes.build_freeform(*vertices[0], True)
    
        ffBuilder.add_line_segments(vertices[1:], close = closed)
    
        s = ffBuilder.convert_to_shape()
        
        if text is not None:
            s.text = text
            p = s.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            if textColor is None:
                setColour(p.font.color, parseColour('#000000'))
            else:
                setColour(p.font.color, parseColour(textColor))
    
        if fill:
            s.fill.solid()
            if fillColor is not None:
                setColour(s.fill.fore_color, parseColour(fillColor))
        
        return s
