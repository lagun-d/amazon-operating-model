Attribute VB_Name = "Module1"
Sub RefreshReport()
    ActiveWorkbook.RefreshAll
    Sheets("Forecast").Range("B3:E6,B8:E11,B14:E14").NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Sheets("Forecast").Range("B12:E12,B15:E15").NumberFormat = "0.0%"
End Sub
Sub ScenarioSummary()
    Sheets("Assumptions").Range("B1").Value = "Base"
    Sheets("Forecast").Range("E6,E11:E12").Copy
    Sheets("Forecast").Range("B19").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    Sheets("Assumptions").Range("B1").Value = "Optimistic"
    Sheets("Forecast").Range("E6,E11:E12").Copy
    Sheets("Forecast").Range("C19").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    Sheets("Assumptions").Range("B1").Value = "Pessimistic"
    Sheets("Forecast").Range("E6,E11:E12").Copy
    Sheets("Forecast").Range("D19").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    Sheets("Assumptions").Range("B1").Value = "Base"
    Application.CutCopyMode = False
End Sub
