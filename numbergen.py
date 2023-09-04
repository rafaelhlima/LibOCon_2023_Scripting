from scriptforge import CreateScriptService
import random as rnd

bas = CreateScriptService("Basic")

def open_dialog(args=None):
    dlg = CreateScriptService("Dialog", "GlobalScope", "NumberGenerator", "NumberGenDlg")
    dlg.Execute()

def btn_cancel_click(event=None):
    # Get the control that was clicked
    control = CreateScriptService("DialogEvent", event)
    # Get the parent dialog and terminate it
    dlg = control.Parent
    dlg.EndExecute(bas.IDCANCEL)

def btn_generate_click(event=None):
    # Get the control that was clicked
    control = CreateScriptService("DialogEvent", event)
    dlg = control.Parent
    # Get the parameters from the dialog
    start_cell = dlg.Controls("editStartCell").Value
    num_cols = int(dlg.Controls("editNumColumns").Value)
    num_rows = int(dlg.Controls("editNumRows").Value)
    mean = float(dlg.Controls("editMean").Value)
    std_dev = float(dlg.Controls("editStdDev").Value)
    # Generate the numbers
    numbers = [[rnd.gauss(mean, std_dev) for _ in range(num_cols)] for _ in range(num_rows)]
    # Get the document and add the data
    doc = CreateScriptService("Document", bas.ThisComponent)
    num_range = doc.Offset(start_cell, width=num_cols, height=num_rows)
    doc.setValue(num_range, numbers)

g_exportedScripts = (open_dialog, btn_cancel_click, btn_generate_click)
