from scriptforge import CreateScriptService
from com.sun.star.awt import FontWeight
from com.sun.star.table import BorderLineStyle
from uno import createUnoStruct, Enum
import math

bas = CreateScriptService("Basic")
ui = CreateScriptService("UI")

def create_writer_doc(args=None):
    name = bas.InputBox("What is your name?")
    doc = ui.createDocument("Writer")
    doc_component = doc.XComponent
    text = doc_component.getText()
    text.insertString(text.End, f"Hello {name}\n", False)

def create_normdist_table(args=None):
    # Get the Document service instance for the current component
    doc = CreateScriptService("Document", bas.ThisComponent)
    # If it is not a Calc document, do nothing
    if not doc.isCalc:
        return
    # Insert the data in the sheet
    doc.setValue("A1:B1", ["z-Value", "P(Z<z)"])
    z_values = [-3 + z * 0.5 for z in range(13)]
    z_range = doc.Offset("A2", height=len(z_values))
    doc.setValue(z_range, z_values)
    # Insert the formulas
    formula_range = doc.Offset("B2", height=len(z_values))
    base_formula = "=NORM.S.DIST(A2;1)"
    doc.setFormula(formula_range, base_formula)
    # Format the header of the table
    range_obj = doc.XCellRange("A1:B1")
    range_obj.CellBackColor = bas.RGB(200, 200, 200)
    range_obj.CharWeight = FontWeight.BOLD
    # Format the remainder of the table
    table_range = doc.Offset("A1", width=2, height=len(z_values)+1)
    range_obj = doc.XCellRange(table_range)
    range_obj.CharFontName = "Arial"
    justify_center = Enum("com.sun.star.table.CellHoriJustify", "CENTER")
    range_obj.HoriJustify = justify_center
    # Struct that defines the line format
    line_format = createUnoStruct("com.sun.star.table.BorderLine2")
    line_format.LineStyle = BorderLineStyle.SOLID
    line_format.LineWidth = 10
    range_obj.TopBorder = line_format
    range_obj.BottomBorder = line_format
    range_obj.LeftBorder = line_format
    range_obj.RightBorder = line_format

def plot_function(args=None):
    # Plot X and Y values for X² function
    data = [[x, math.pow(x, 2)] for x in range(-5, 6)]
    doc = CreateScriptService("Calc", bas.ThisComponent)
    doc.setValue("A1:B1", [["X", "Y"]])
    data_range = doc.Offset("A2", width=2, height=len(data))
    doc.setValue(data_range, data)
    # Select the entire table
    table_range = doc.Region("A1")
    # Insert the chart
    cur_sheet = doc.SheetName(doc.CurrentSelection)
    chart = doc.CreateChart("X-Squared", cur_sheet, table_range)
    chart.ChartType = "XY"
    chart.Legend = False
    chart.Title = "X-Squared Plot"
    chart.XTitle = "X values"
    chart.YTitle = "Y values"
    # Set the line tipe to "smooth" (using cubic splines)
    diagram = chart.XDiagram
    diagram.SplineType = Enum("com.sun.star.chart2.CurveStyle", "CUBIC_SPLINES")
    # Place the Y-axis at -6 for better visualization
    y_axis = diagram.getYAxis()
    y_axis.CrossoverPosition = Enum("com.sun.star.chart.ChartAxisPosition", "VALUE")
    y_axis.CrossoverValue = -6

g_exportedScripts = (create_normdist_table, create_writer_doc, plot_function)

