Attribute VB_Name = "InternetConnectivity"
Function IsInternetReachable() As Boolean
    On Error GoTo NoInternet
    Dim http As Object
    Set http = CreateObject("WinHTTP.WinHTTPRequest.5.1")
    
    http.Open "GET", "http://clients3.google.com/generate_204", False
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.Send
    
    IsInternetReachable = (http.Status = 204)
    Exit Function

NoInternet:
    IsInternetReachable = False
End Function
