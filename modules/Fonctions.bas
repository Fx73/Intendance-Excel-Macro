Attribute VB_Name = "Fonctions"
Function CreateSheet(sname As String, position As Integer)
    CreateSheet = 0
    For Each s In ActiveWorkbook.Worksheets
        If (s.Name = sname) Then Exit Function
    Next
    
    Dim ws As Worksheet

    If Sheets.Count >= position Then
        Set ws = ThisWorkbook.Sheets.Add(before:=Sheets(position))
    Else
        Set ws = ThisWorkbook.Sheets.Add(after:=Sheets(Sheets.Count))
    End If
    ws.Name = sname
    CreateSheet = 1

End Function


Function GetPingResult(Host)
   'declaring variables
   Dim objPing As Object
   Dim objStatus As Object
   Dim Result As String
   'ping the host
   Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
       ExecQuery("Select * from Win32_PingStatus Where Address = '" & Host & "'")
   'report the results
   For Each objStatus In objPing
      Select Case objStatus.StatusCode
         Case 0: GetPingResult = "Connected"
         Case 11001: GetPingResult = "Buffer too small"
         Case 11002: GetPingResult = "Destination net unreachable"
         Case 11003: GetPingResult = "Destination host unreachable"
         Case 11004: GetPingResult = "Destination protocol unreachable"
         Case 11005: GetPingResult = "Destination port unreachable"
         Case 11006: GetPingResult = "No resources"
         Case 11007: GetPingResult = "Bad option"
         Case 11008: GetPingResult = "Hardware error"
         Case 11009: GetPingResult = "Packet too big"
         Case 11010: GetPingResult = "Request timed out"
         Case 11011: GetPingResult = "Bad request"
         Case 11012: GetPingResult = "Bad route"
         Case 11013: GetPingResult = "Time-To-Live (TTL) expired transit"
         Case 11014: GetPingResult = "Time-To-Live (TTL) expired reassembly"
         Case 11015: GetPingResult = "Parameter problem"
         Case 11016: GetPingResult = "Source quench"
         Case 11017: GetPingResult = "Option too big"
         Case 11018: GetPingResult = "Bad destination"
         Case 11032: GetPingResult = "Negotiating IPSEC"
         Case 11050: GetPingResult = "General failure"
         Case Else: GetPingResult = "Unknown host"
      End Select
   Next
   'reset object ping variable
   Set objPing = Nothing
End Function

Sub GetDataFromGoogle(wsn As String, addressKey As String)
Dim googlePrefix As String: googlePrefix = "https://spreadsheets.google.com/tq?tqx=out:html&tq=&key="
Dim i As Integer
  With Worksheets(wsn)
    With .QueryTables.Add(Connection:="URL;" & googlePrefix & addressKey, Destination:=.Range("$A$1"))
        .PreserveFormatting = False
        .WebFormatting = xlWebFormattingNone
        .Refresh BackgroundQuery:=False
    End With
    DoEvents
  End With
    For i = 1 To ThisWorkbook.Connections.Count
        If ThisWorkbook.Connections.Count = 0 Then Exit Sub
        ThisWorkbook.Connections.item(i).Delete
    i = i - 1
    Next i
End Sub


Function GetSMTPGmailServerConfig() As Object
Dim Cdo_Config As New CDO.Configuration
Dim Cdo_Fields As Object

Set Cdo_Fields = Cdo_Config.Fields
With Cdo_Fields
.item(cdoSendUsingMethod) = cdoSendUsingPort
.item(cdoSMTPServer) = "smtp.gmail.com"
.item(cdoSMTPServerPort) = 465
.item(cdoSendUserName) = "AIErrorHandler@gmail.com"
.item(cdoSendPassword) = "ErrorH1234"
.item(cdoSMTPAuthenticate) = cdoBasic
.item(cdoSMTPUseSSL) = True
.Update
End With

Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Public Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

Public Function NextRow(a As Range, ws As String) As Integer
    Dim c As Integer: c = a.Column
    Dim r As Integer: r = a.row + 1
    While Worksheets(ws).Cells(r, c) = "" And Worksheets(ws).Cells(r, c + 1) <> ""
        r = r + 1
    Wend
    NextRow = r - 1
End Function

Public Sub CopySaveAt(path As String)
    Dim savepath As String: savepath = ThisWorkbook.path & "\" & ActiveWorkbook.Name
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs path
    ActiveWorkbook.SaveAs savepath
    Application.DisplayAlerts = True
End Sub


Function ReadBinaryFile(FileName)
    Const adTypeBinary = 1
    
    'Create Stream object
    Dim BinaryStream
    Set BinaryStream = CreateObject("ADODB.Stream")
    
    'Specify stream type - we want To get binary data.
    BinaryStream.Type = adTypeBinary
    
    'Open the stream
    BinaryStream.Open
    
    'Load the file data from disk To stream object
    BinaryStream.LoadFromFile FileName
    
    'Open the stream And get binary data from the object
    ReadBinaryFile = BinaryStream.Read
End Function
