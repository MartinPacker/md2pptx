"""
runPython
"""

version = "0.6"

import csv
from pptx.chart.data import CategoryChartData
from pptx.oxml.xmlchemy import OxmlElement, serialize_for_reading
from pptx.oxml import parse_xml
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from colour import setColour, parseColour
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import PP_PLACEHOLDER
from paragraph import *
from pptx.util import Inches, Pt
import globals

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

    def doChecklistChecks(placeholder, checklist, colourChecks = False):
        tf = placeholder.text_frame
        paras = tf.paragraphs

        for paraNumber, para in enumerate(paras):
            # Save original font size
            originalFontSize = para.font.size
            
            # Save original indentation level
            level = para.level
            
            # Remove the original pPr element
            para._element.remove(para._element.getchildren()[0])
    
            xml = ''
            
            # Note the level insertion
            xml += f'<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" marL="{Inches(0.25) * level}" indent="{Inches(0.33)}" lvl="{level}">'


            if checklist[paraNumber] == None:
                # Checkbox unchecked

                # Set the bullet to an empty square - and don't colour it
                xml += '<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="Wingdings" pitchFamily="2" charset="2"/>'
                xml += '<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="o"/>'
            elif checklist[paraNumber] == True:
                # Checkbox ticked

                # Maybe colour the mark
                if colourChecks:
                    xml += '<a:buClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    xml += '<a:srgbClr val="00FF00" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />'
                    xml += '</a:buClr>'

                # Set the bullet to a square with a tick
                xml += '<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="Wingdings 2" pitchFamily="2" charset="2"/>'
                xml += '<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="R"/>'
            else:
                # Checkbox crossed

                # Maybe colour the mark
                if colourChecks:
                    xml += '<a:buClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    xml += '<a:srgbClr val="FF0000" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />'
                    xml += '</a:buClr>'
                
                # Set the bullet to a square with a cross
                xml += '<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="Wingdings 2" pitchFamily="2" charset="2"/>'
                xml += '<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="T"/>'
                
     
            xml += '</a:pPr>'
    
            # Parse this XML
            parsed_xml = parse_xml(xml)
    
            # Insert the parsed XML fragment as a child of the pPr element
            para._element.insert(0, parsed_xml)
    
            # Restore original font size
            para.font.size = originalFontSize

        return placeholder

    def makeChecklist(placeholder, checklist, checkTextIndex = 0, checkMarkIndex = 1, levelIndex = 2, colourChecks = False):
        checkMarks = []
        
        for paraNumber, checklistItem in enumerate(checklist):
            checkMarks.append(checklistItem[checkMarkIndex])

            if paraNumber == 0:
               para = placeholder.text_frame.paragraphs[0]
            else:
               para = placeholder.text_frame.add_paragraph()

            addFormattedText(para, checklistItem[checkTextIndex])
            
            # The actual level setting in the surviving XML is done by doChecklistChecks
            # The below sets the level as a hint for doChecklistChecks to work with
            if len(checklistItem) > levelIndex:
                try:
                    level = int(checklistItem[levelIndex])
                    
                    para.level = level - 1
                except ValueError:
                    para.level = 0
            else: para.level = 0

        RunPython.doChecklistChecks(placeholder, checkMarks, colourChecks)
        
        return placeholder

    def makeTruthy(table_array, columnNumber = 1, trueString = "Yes", falseString = "No", unsetString = ""):
        for row in table_array:
            if row[columnNumber] == unsetString:
                row[columnNumber] = None
            elif row[columnNumber] == trueString:
                row[columnNumber] = True
            else:
                row[columnNumber] = False

        return table_array

    def ensureTextbox(slide, renderingRectangle, shapeIndex = None):
        if shapeIndex is not None:
            if len(slide.shapes) < shapeIndex + 1:
                # Need to create a text box as shape index too high
                newShape = slide.shapes.add_textbox(
                    renderingRectangle.left,
                    renderingRectangle.top,
                    renderingRectangle.width,
                    renderingRectangle.height
                )

                # Return new shape
                return newShape
            else:
                # shape Index is valid
                if (theShape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX) | (theShape.placeholder_format.type == PP_PLACEHOLDER.OBJECT):
                    # shape is a text box so return it
                    return slide.shapes[shapeIndex]
                else:
                    # Shape isn't a text box so create one
                    newShape = slide.shapes.add_textbox(renderingRectangle.left, renderingRectangle.top, renderingRectangle.width, renderingRectangle.height)

                    # Return new shape
                    return newShape

        # shapeIndex wasn't set so potentially search for the last text box
        if len(slide.shapes) < 2:
            # Need to create a text box as only shape is presumed to be a title
            newShape = slide.shapes.add_textbox(renderingRectangle.left, renderingRectangle.top, renderingRectangle.width, renderingRectangle.height)

            # Return new shape
            return newShape
        
        # Search for last text box
        for shapeIndex, theShape in reversed(list(enumerate(slide.shapes))):
            if (theShape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX) | (theShape.placeholder_format.type == PP_PLACEHOLDER.OBJECT):
                return theShape


        # Need to create a text box none found - other than perhaps at Index 0 (title)
        newShape = slide.shapes.add_textbox(renderingRectangle.left, renderingRectangle.top, renderingRectangle.width, renderingRectangle.height)

        # Return new shape's index as presumed to be the text box we want
        return newShape
        
    def checklistFromCSV(slide, renderingRectangle, filename, shapeIndex = None, colourChecks = False):
        # Read in CSV and turn second column into "truthy" values
        myChecklist = RunPython.makeTruthy(RunPython.readCSV(filename), 1)

        # Ensure we have a placeholder - whether first or second or specified
        textShape = RunPython.ensureTextbox(slide, renderingRectangle, shapeIndex)

        # Make the checklist in this placeholder from the imported file
        RunPython.makeChecklist(textShape, myChecklist, 0, 1, 2, colourChecks)
        
        return textShape

    def removeBullet(theShape, paragraphNumber):
        removeBullet(theShape.text_frame.paragraphs[paragraphNumber])
        
        return theShape
    
    def removeBullets(theShape):
        for p in theShape.text_frame.paragraphs:
            removeBullet(p)

        return theShape

    def removeSelectedBullets(theShape, removalArray):
        removeSelectedBullets(theShape.text_frame, removalArray)

        return theShape

    def getParagraphs(slide, wantedParagraphs = []):
        return getParagraphs(slide, wantedParagraphs)