Private Sub Worksheet_Change(ByVal Target As Range)
    ' ### Range
    ' ref: https://docs.microsoft.com/en-us/office/vba/api/excel.range(object)
    ' methods & properties on the left menu
    
    Set myRange = ActiveSheet.Range("$e$6")

    Select Case Target.Address
    Case myRange.Address
        myRange.Activate
        ' #### Item(dx, dy) = relative address
        myRange.Item(1, 2) = "Chamando API..."
    Case Else
        ' Default case
    End Select
End Sub
