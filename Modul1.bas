Attribute VB_Name = "Modul1"
Function GPT(query As String) As String
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.ServerXMLHTTP")
    
    Dim url As String
    url = "https://api.openai.com/v1/chat/completions"
    
    Dim apiKey As String
    apiKey = "Your_OpenAI_API_Key"  ' Replace with your actual OpenAI API key
    
    Dim requestBody As String
    requestBody = "{""model"": ""gpt-4o-mini"", ""messages"": [{""role"": ""user"", ""content"": """ & query & """}]}"

    ' Send a POST request
    httpObject.Open "POST", url, False
    httpObject.setRequestHeader "Content-Type", "application/json"
    httpObject.setRequestHeader "Authorization", "Bearer " & apiKey
    httpObject.send (requestBody)
    
    Dim response As String
    response = httpObject.responseText
    
    ' Extract the response (you may need to adjust the parsing based on the response structure)
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    GPT = json("choices")(1)("message")("content")
    
    Set httpObject = Nothing
End Function

