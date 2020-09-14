VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public folio As Long ' Entrega el folio asignado al DTE.
Public ok As Boolean    ' Indica que el DTE se gener� correctamente.
Public error As String
Public escpos As String

Function PrintEscPos(port As String) As Boolean
    Dim data As String
    data = decodeBase64(escpos)
    Open port For Binary Access Write As #1
    Put #1, 1, data
    Close #1
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

