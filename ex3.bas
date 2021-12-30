Sub restAPICall()
Dim objRequest As Object
Dim strUrl As String
Dim blnAsync As Boolean
Dim strResponse As String
Dim json As Object

Set objRequest = CreateObject("MSXML2.XMLHTTP")
strUrl = "https://jsonplaceholder.typicode.com/posts/1"
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
    

Debug.Print strResponse

Set json = JsonConverter.ParseJson(strResponse)
Debug.Print json("title")

End Sub
