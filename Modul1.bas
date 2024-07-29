Attribute VB_Name = "Modul11"
Function GPT(query As String, apiKey As String) As String
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.ServerXMLHTTP")

    Dim url As String
    url = "https://api.openai.com/v1/chat/completions"

    ' Encode query to handle special characters and line breaks
    query = Replace(query, vbCrLf, "\n")
    query = Replace(query, vbCr, "\n")
    query = Replace(query, vbLf, "\n")
    query = Replace(query, """", "\""")
    
    Dim requestBody As String
    requestBody = "{""model"": ""gpt-4o-mini"", ""messages"": [{""role"": ""user"", ""content"": """ & query & """}]}"

    ' Send a POST request
    httpObject.Open "POST", url, False
    httpObject.setRequestHeader "Content-Type", "application/json"
    httpObject.setRequestHeader "Authorization", "Bearer " & apiKey
    On Error GoTo ErrorHandler
    httpObject.send (requestBody)
    
    Dim response As String
    response = httpObject.responseText
    
    ' Extract the response (you may need to adjust the parsing based on the response structure)
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    GPT = json("choices")(1)("message")("content")
    
    Set httpObject = Nothing
    Exit Function

ErrorHandler:
    GPT = "Error: " & Err.Description
    Set httpObject = Nothing
End Function

