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
   Begin VB.CheckBox chkLocalServer 
      Caption         =   "Conectar a Servidor Local"
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Imprimir en LPT1"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CheckBox chkDownloadPDF 
      Caption         =   "Descargar PDF"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
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
    
    If chkLocalServer.Value Then
        ' Para conectarse al Servidor Local de SuperFactura y poder emitir documentos en forma offline.
        ' Ver: https://blog.superfactura.cl/servicio-offline-para-puntos-de-venta/
        api.url = "http://localhost:9080"
    End If

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
    
    If chkPrint.Value Then
        ' Obtener formato Esc/Pos para impresoras térmicas
        ' Ver: https://blog.superfactura.cl/impresion-con-impresoras-termicas/
        api.SetOption "getEscPos", True
		
        ' Indicar el modelo de impresora
        api.SetOption "modelo", "default"
    End If

    If chkDownloadPDF.Value Then
        ' Indicar que queremos guardar los PDF (original y copia cedible)
        api.SetSavePDF ruta & "\dte-" & idInterno
    End If

    ' Indicar que queremos guardar el archivo XML. No es necesario y conviene evitar.
    ' api.SetSaveXML "C:\Documents and Settings\Kripper\Desktop\dte-" & idInterno
    
    ' Enviar DTE y obtener resultados
    Dim res As apiResult
    Set res = api.SendDTE(data, "cer")
    
    If res.ok Then
        MsgBox "Se creó el DTE con folio " & res.folio
        
        If chkPrint.Value Then
            ' Imprimir a impresora térmica
            res.PrintEscPos "LPT1:"
        End If
    Else
        ' IMPORTANTE: Este mensaje de error se debe mostrar al usuario para que pueda recibir soporte.
        MsgBox "API ERROR: " & res.Error
    End If
End Sub

