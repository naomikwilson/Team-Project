
vba_code = ("""
Sub Benchmarking()
'Benchmarking Macro
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "='Operating Statistics'!R[8]C[-2]"
    Range("C6:O6").Select
    Selection.FillRight
    Sheets("Operating Statistics").Select
    Range("A15").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("COMPCO AND BENCHMARKING").Select
    Range("C7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Columns.AutoFit
    Range("D7").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=INDEX('Operating Statistics'!R14C[-2]:R27C[-2],MATCH('COMPCO AND BENCHMARKING'!RC3,'Operating Statistics'!R14C1:R27C1,0))"

   
    Range("D7:D17").Select
    Selection.FillDown
    Range("D7:O17").Select
    Selection.FillRight
    Range("C22").Select
    ActiveCell.FormulaR1C1 = "Percentile"
    Range("C23").Select
    ActiveCell.FormulaR1C1 = "0%"
    Range("C24").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+5%"
    Range("C24:C43").Select
    Selection.FillDown
    Sheets("Operating Statistics").Select
    Range("A14").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    
    Selection.Copy
    Sheets("COMPCO AND BENCHMARKING").Select
    Range("C6").Select
    Selection.End(xlDown).Select
    Range("C17").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Macro2()

End Sub



""")


def extract_macro_names(vba_code):
    """
    Extracts macro name from str which is used to execute code
    """
    macro_names = []
    for line in vba_code.splitlines():
        if line.startswith("Sub "):
            macro_names.append(line[4:].strip("()"))
    return macro_names


print(extract_macro_names(vba_code))
