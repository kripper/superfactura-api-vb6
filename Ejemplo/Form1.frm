VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ejemplo API SuperFactura"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDownloadPDF 
      Caption         =   "Descargar PDF"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox tbDocumentID 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "F123"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox tbData 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton cmdEmitirFactura 
      Caption         =   "Emitir Factura"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "DocumentID :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4575
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEmitirFactura_Click()
    Dim api As New api
    api.user = "usuario@cliente.cl"
    api.password = "mynewpassword"

    Dim ruta As String
    ruta = "C:\Documents and Settings\Kripper\Desktop"
    
    Dim data As String
    data = tbData.Text
    
    ' Generar/obtener ID interno. Debe ser único. Se puede usar el ID de su Base de Datos.
    ' No confundir con el folio.
    Dim idInterno As String
    idInterno = Form1.tbDocumentID.Text

    ' Enviar documentID (importante para evitar documentos duplicados en caso de falla de red y reenvío):
    ' Si se envía un ID ya utilizado, se retornará el mismo documento, en vez de crear uno nuevo.
    api.SetOption "documentID", idInterno

    If chkDownloadPDF.Value Then
        ' Indicar que queremos guardar los PDF (original y copia cedible)
        api.SetSavePDF ruta & "\dte-" & idInterno
    End If

    ' Indicar que queremos guardar el archivo XML
    ' api.SetSaveXML "C:\Documents and Settings\Kripper\Desktop\dte-" & idInterno
    
    ' Enviar DTE y obtener resultados
    Dim res As apiResult
    Set res = api.SendDTE(data, "cer")
    
    If res.ok Then
        MsgBox "Se creó el DTE con folio " & res.folio
    Else
        MsgBox "ERROR: " & res.Error
    End If
End Sub

Private Sub Form_Load()
    ' Sugerencia: Evite acciones innecesarias para depurar problemas que requieran reiniciar la ejecución.
    ' cmdEmitirFactura_Click
    ' Unload Me
End Sub

