VERSION 5.00
Begin VB.Form frm_VTA_GuiasxFacturar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guias del Pedido"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   4920
      Width           =   7935
      Begin VB.OptionButton optFactura 
         Caption         =   "Factura"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "Fac"
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optBoleta 
         Caption         =   "Boleta"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Tag             =   "BOL"
         Top             =   240
         Width           =   975
      End
      Begin vbp_Ventas.ctlDataCombo cboVenta 
         Height          =   315
         Left            =   3720
         TabIndex        =   20
         Top             =   180
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         MatchEntry      =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   7935
      Begin VB.OptionButton optFacturadas 
         Caption         =   "Facturadas"
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Tag             =   "FAT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optTransito 
         Caption         =   "Pendientes por Facturar"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Tag             =   "TRA"
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame fraTipoFactura 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   7935
      Begin VB.OptionButton optMultiplesFacturas 
         Caption         =   "Multiples Documentos"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optUnaFactura 
         Caption         =   "Un Documento"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblTotSist 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   6600
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total =>"
         Height          =   195
         Left            =   5880
         TabIndex        =   12
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7935
      Begin VB.Label lblPedido 
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F. Pedido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblFechaPedido 
         Caption         =   "xx/xx/xxxx"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label lblCliente 
         Caption         =   "xxxxxxxxxx"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   330
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdReimprimir 
      Caption         =   "Reimprimir"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "Generar Documento"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla grdDetalle 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4471
   End
End
Attribute VB_Name = "frm_VTA_GuiasxFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPedido As New clsPedido
Dim rs As oraDynaset
Dim StrGuias As String
Dim strNumDoc As String
Dim strTipoDoc As String
Dim ConVta As oraDynaset


Public Sub Carga(ByVal NumeroPedido As String, ByVal strEstadoGuia As String)
On Error GoTo handle
    lblPedido.Caption = NumeroPedido
    Call CargaGuias(NumeroPedido, strEstadoGuia)
    Set rs = objPedido.Cabecera_Pedido(objUsuario.CodigoLocal, NumeroPedido)
    lblCliente.Caption = "" & rs("CLIENTE").Value
    lblFechaPedido.Caption = "" & rs("FCH_EMISION").Value
    cmdFacturar.Enabled = False
    
    
    Set cboVenta.RowSource = objPedido.ListaCondVenta
    cboVenta.ListField = "DES"
    cboVenta.BoundColumn = "COD"
       
    Set ConVta = objPedido.Dev_CondVenta(rs("RUC").Value)
   
    If IsNull(ConVta("CONVTA").Value) Then
       MsgBox "No se encontró la Condición Venta del Cliente ... Comunicarse con Creditos y Cobranzas ", vbQuestion, Caption
       cboVenta.BoundText = "*"
    Else
       cboVenta.BoundText = ConVta("CONVTA").Value
    End If
    
    Me.Show vbModal
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub CargaGuias(ByVal NumeroPedido As String, ByVal strEstadoGuia As String)
 Set grdDetalle.DataSource = objPedido.listaGuias(NumeroPedido, strEstadoGuia, objUsuario.CodigoLocal)
 SeteaGrilla
 CalcularTotal
End Sub



Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant

    arrCampos = Array("NUM_GUIA", "FCH_EMISION", "DIR_ENTREGA", "TOT_ITEM", "EST_GUIA", "MTO_TOTAL")
    arrCaption = Array("Codigo", "Emision", "Direccion Entrega", "# Items", "Estado", "Total")
    arrAncho = Array(1000, 1000, 3200, 800, 700, 1000)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgRight)
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDetalle.Columns("MTO_TOTAL").NumberFormat = "####0.00"
End Sub

Private Sub cmdFacturar_Click()
    Dim strMensaje As String
    If grdDetalle.ApproxCount > 0 Then
        ConcatenarGuias
        If optFactura.Value Then
            strTipoDoc = "FAC"
        Else
            strTipoDoc = "BOL"
        End If
        strMensaje = objPedido.GeneraFactura(lblPedido.Caption, objUsuario.CodigoLocal, objUsuario.Codigo, StrGuias, objUsuario.NombrePC, IIf(optUnaFactura.Value = True, 1, 0), strTipoDoc, strNumDoc, cboVenta.BoundText)
        If strMensaje = "" Then
           MsgBox "Se grabo satisfactoriamente la(s) siguientes facturas: " + strNumDoc, vbInformation, App.ProductName
           ImprimirFacturas
           Call CargaGuias(lblPedido.Caption, optTransito.Tag)
           Exit Sub
        Else
            MsgBox strMensaje, vbCritical, App.ProductName
        End If
    End If
End Sub

Private Sub ImprimirFacturas()
    Dim strVectorFacturas() As String
    Dim strVectorTipo() As String
    Dim i As Integer
            
    strVectorFacturas = Split(strNumDoc, "|")
    strVectorTipo = Split(strTipoDoc, "|")
    intNumImpresion = 0
    For i = LBound(strVectorFacturas) To UBound(strVectorFacturas) - 1
        psub_Imprime_Doc_Cliente objUsuario.CodigoEmpresa, strVectorFacturas(i), strVectorTipo(i)
    Next
End Sub

Private Sub ConcatenarGuias()
   On Error GoTo CtrlErr
   Dim i As Integer
   StrGuias = ""
   grdDetalle.DataSource.MoveFirst
   While i < grdDetalle.DataSource.RecordCount
       StrGuias = StrGuias + grdDetalle.Columns(0).Value + "|"
       grdDetalle.DataSource.MoveNext
       i = i + 1
   Wend
   
   Exit Sub
CtrlErr:
     MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub

Private Sub CalcularTotal()
   On Error GoTo CtrlErr
   Dim i As Integer
   Dim TotGuias As Double
   
   TotGuias = 0
   grdDetalle.DataSource.MoveFirst
   While i < grdDetalle.DataSource.RecordCount
       TotGuias = TotGuias + grdDetalle.Columns(5).Value
       grdDetalle.DataSource.MoveNext
       i = i + 1
   Wend
   
  LblTotSist = Format(TotGuias, "#,###,##0.00")
   
   Exit Sub
CtrlErr:
     MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub


Private Sub cmdReimprimir_Click()
    If grdDetalle.ApproxCount > 0 Then
        psub_Imprime_Guia_Cliente grdDetalle.Columns(0).Value
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub optFacturadas_Click()
If optFacturadas.Value = True Then
    ActivarBoton
    Call CargaGuias(lblPedido.Caption, optFacturadas.Tag)
 End If
End Sub

Private Sub optTransito_Click()
 If optTransito.Value = True Then
      ActivarBoton
      Call CargaGuias(lblPedido.Caption, optTransito.Tag)
 End If
End Sub

Private Sub optUnaFactura_Click()
   ActivarBoton
End Sub

Private Sub optMultiplesFacturas_Click()
   ActivarBoton
End Sub

Private Sub ActivarBoton()
    If optTransito.Value = True Then
        cmdFacturar.Enabled = optUnaFactura.Value = True Or optMultiplesFacturas.Value = True
    Else
        cmdFacturar.Enabled = False
    End If
End Sub



