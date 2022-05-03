# vba-json-update
a simple vba script to parse json text and update in excel

```vba

Private Sub btnusdupdate_Click()

'get all parameters string from table
Dim W As Worksheet
Set W = ActiveSheet
Dim Last As Integer
Last = W.Range("A100").End(xlUp).Row 'update when table chanegd
Dim Symbols As String
Dim i As Integer
For i = 2 To Last
    Symbols = Symbols & W.Range("A" & i).Value & "," 'update when table changed
Next i
'remove last , in ticker query
Symbols = Left(Symbols, Len(Symbols) - 1)

'start api query

Dim URL As String
URL = "https://yfapi.net/v6/finance/quote?symbols=" & Symbols 'replace this when needed

Dim request As New WinHttpRequest

request.Open "Get", URL
request.SetRequestHeader "X-API-KEY", "yourapikey" 'replace with your own api key
request.SetRequestHeader "Content-Type", "application/json"
request.Send

'stop if request status is not ok
If request.Status <> 200 Then
MsgBox request.ResponseText
Exit Sub
End If

'use jsonconverter to parse json text - remember to load microsoft scripting runtime lib

Dim response As Object
Set response = JsonConverter.ParseJson(request.ResponseText) 'require to doanload module from github

'put json response into a collection, use 2 brackets for nested structure

Dim results As Collection
Set results = response("quoteResponse")("result")

Dim s As Integer

s = 2

Dim tickerinfo As Dictionary
For Each tickerinfo In results

'update cell value

W.Cells(s, 2).Value = tickerinfo("regularMarketPrice") 'update when table chanegd
W.Cells(s, 3).Value = tickerinfo("fiftyTwoWeekHigh") 'update when table chanegd
W.Cells(s, 4).Value = tickerinfo("fiftyTwoWeekLow") 'update when table chanegd

s = s + 1

Next tickerinfo


End Sub


```
