Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$B$1" Then
        Debug.Print (Target.Address & ": " & Target.Text)
        ActiveSheet.Range("$B$1").Activate
        ActiveSheet.Range("B3:B6").Cells.ClearContents

        Dim objRequest As Object
        Dim strUrl As String
        Dim blnAsync As Boolean
        Dim strResponse As String
        Dim json As Object
        
        Set objRequest = CreateObject("MSXML2.XMLHTTP")
        strUrl = "https://jsonplaceholder.typicode.com/posts/" & Target.Text
        blnAsync = True
        
        With objRequest
            .Open "GET", strUrl, blnAsync
            .SetRequestHeader "Content-Type", "application/json"
            .send
            
            While objRequest.readyState <> 4
                DoEvents
            Wend
                
            strResponse = .responseText
        End With

        Set json = JsonConverter.ParseJson(strResponse)
        'Debug.Print json("title")
        ActiveSheet.Range("$B$3").Value = json("id")
        ActiveSheet.Range("$B$4").Value = json("title")
        ActiveSheet.Range("$B$5").Value = json("body")
        ActiveSheet.Range("$B$6").Value = json("userId")
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Debug.Print (Target.Address)
End Sub
