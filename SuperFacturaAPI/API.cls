VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public user As String, password As String
Private jsonOptions As String
Private savePDF As String
Private saveXML As String
Private oZLIB As New cZLIB
Public URL As String

Private Const CP_UTF8 = 65001

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

Public Function UTF8_URLEncode(str As String) As String
    UTF8_URLEncode = URLUtility.URLEncode(StrConv(Utf8BytesFromString(str), vbUnicode))
End Function


Public Sub SetSavePDF(outputFile As String)
    savePDF = outputFile
    SetOption "getPDF", "1"
End Sub

Public Sub SetSaveXML(outputFile As String)
    saveXML = outputFile
    SetOption "getXML", "1"
End Sub

Public Sub SetOption(key As String, val As String)
    If jsonOptions <> "" Then jsonOptions = jsonOptions + ","
    jsonOptions = jsonOptions + EscapeArgument(key) + ":" + EscapeArgument(val)
End Sub

Private Function SendRequest(jsonData As String, jsonOptions As String) As String
    Dim req As New WinHttpRequest
    
    ' Ignorar validación SSL ya que generalmente los entornos con VB6 no
    ' soportan los certificados SSL.
    ' Si su aplicación se ejecutará en Windows 10 o superior puede comentar esta linea.
    req.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 13056
    
    req.SetTimeouts 50000, 50000, 50000, 50000

    If URL = "" Then URL = "https://superfactura.cl"

    req.Open "POST", URL & "/?a=json", False

    req.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    req.SetRequestHeader "User-Agent", "SuperFacturaAPI-VB6-v1.0"

    'On Local Error Resume Next
    
    Dim postString As String
    postString = "user=" & user & "&pass=" & password & "&content=" & UTF8_URLEncode(jsonData) & "&options=" & UTF8_URLEncode(jsonOptions)
    
    req.Send postString
    
    Dim rawOutput As String
    rawOutput = StrConv(req.ResponseBody, vbUnicode)

    Dim jsonString As String

    Dim err As Long

    err = oZLIB.UncompressString(rawOutput, Z_AUTO)
    If err Then
        MsgBox "zlib error"
        Exit Function
    End If
    
    SendRequest = rawOutput ' Return decompressed output
End Function

Private Function encodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument
    
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    encodeBase64 = objNode.Text
 
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Private Function decodeBase64(ByVal strData As String) As String
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    decodeBase64 = StrConv(objNode.nodeTypedValue, vbUnicode)
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Public Function SendDTE(jsonData As String, ambiente As String) As APIResult
    SetOption "ambiente", ambiente

    Set SendDTE = New APIResult

    Dim output As String
    output = SendRequest(jsonData, "{" + jsonOptions + "}")
    
    Dim obj As Object
    Set obj = JSON.parse(output)

    If obj.Item("ack") <> "ok" Then
        SendDTE.error = IIf(obj.Item("response").Item("title") <> "", obj.Item("response").Item("title") & " - ", "") & obj.Item("response").Item("message")
        Exit Function
    End If

    Dim appRes As Object
    Set appRes = obj.Item("response")

    Dim folio As Long
    folio = Conversion.CLng(appRes.Item("folio"))
    
    If appRes.Item("ok") = "1" Then
        SendDTE.ok = True
        SendDTE.folio = folio
        SendDTE.escpos = appRes.Item("escpos")

        If savePDF <> "" Then
            WriteFile savePDF & ".pdf", decodeBase64(appRes.Item("pdf"))

            If appRes.Item("pdfCedible") <> "" Then
                WriteFile savePDF & "-cedible.pdf", decodeBase64(appRes.Item("pdfCedible"))
            End If
        End If

        If saveXML <> "" Then
            WriteFile saveXML & ".xml", appRes.Item("xml")
        End If
    Else
        SendDTE.error = "RPC Error"
    End If
End Function

Sub WriteFile(filename As String, Data As String)
    Dim f As Integer
    f = FreeFile()
    Open filename For Binary As #f
    Put #f, , Data
    Close #f
End Sub

Private Function EscapeArgument(arg As String) As String
    arg = Replace(arg, "\", "\\")
    arg = Replace(arg, "\""", "\\""")
    EscapeArgument = """" & arg & """"
End Function


