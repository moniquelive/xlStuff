Option Explicit

' # VBA
' ref: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference
' funcs: https://docs.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications

' # Excel
' examples (left menu): https://docs.microsoft.com/en-us/office/vba/api/overview/excel

Sub Workbook()
    ' ### Workbook
    ' ref: https://docs.microsoft.com/en-us/office/vba/api/excel.workbook
    ' singletons: ActiveWorkbook, ThisWorkbook
    ' all: Workbooks("file.xls")
    'ThisWorkbook.Worksheets("sheet 1").Activate
    Dim wbBook As Workbook
    
    Set wbBook = ActiveWorkbook
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
End Sub

Sub Worksheet()
    ' ### Worksheet
    ' creating: https://docs.microsoft.com/en-us/office/vba/excel/concepts/workbooks-and-worksheets/create-or-replace-a-worksheet
    With Worksheets("sheet1")
        .Cells.ClearContents ' clear all cells
        
        Dim aRange As Range ' a saved range
        Set aRange = .Range("a1")
        aRange = "Hello World!"
        aRange.Font.Bold = True
        aRange.Activate ' Jump to A1

        .Range("a2:e3") = "Many"
    End With
End Sub
