VERSION 5.00
Begin VB.Form frm_VTA_Previa 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Previa"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "frm_VTA_Previa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Método de entrega"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   0
      TabIndex        =   37
      Top             =   1920
      Width           =   8175
      Begin vbp_Ventas.ctlGrillaArray grdMetodosEntrega 
         Height          =   855
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1508
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin vbp_Ventas.ctlDataCombo CboTeleoperadores 
      Height          =   315
      Left            =   1320
      TabIndex        =   34
      Top             =   4980
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.PictureBox PicFpago 
      Height          =   1185
      Left            =   80
      ScaleHeight     =   1125
      ScaleWidth      =   6330
      TabIndex        =   31
      Top             =   3555
      Visible         =   0   'False
      Width           =   6390
      Begin VB.Label LblFpago 
         Caption         =   "NO TIENE FORMA DE PAGO POR SER UN CONVENIO 100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   6315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Relacion de Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1845
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   8205
      Begin vbp_Ventas.ctlGrillaArray grdProductos 
         Height          =   1545
         Left            =   60
         TabIndex        =   27
         Top             =   210
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   2725
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   20
      Top             =   5400
      Width           =   8175
      Begin VB.CheckBox Check3 
         Caption         =   "Pedido Urgente"
         Height          =   375
         Left            =   3660
         TabIndex        =   33
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtObsCliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   5220
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   240
         Width           =   2790
      End
      Begin VB.TextBox txtObsLocal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   495
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4620
         TabIndex        =   22
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000004&
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   8175
      TabIndex        =   1
      Top             =   6180
      Width           =   8235
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   7020
         Picture         =   "frm_VTA_Previa.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblRedondeo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   4170
         TabIndex        =   19
         Top             =   1095
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000018&
         Caption         =   "Vuelto Redondeado ====== >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Top             =   1350
         Width           =   3540
      End
      Begin VB.Label lblTotalPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   3360
         TabIndex        =   17
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Redondeo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   1095
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblVuelto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   4170
         TabIndex        =   15
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vuelto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   540
      End
      Begin VB.Label lblPagado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   4170
         TabIndex        =   13
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4170
         TabIndex        =   11
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblTotalaPagar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Co-Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblcopago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   4170
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblPctCopago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   2445
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   5640
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total en Dolares"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   1875
         Width           =   1335
      End
      Begin VB.Label lblTotalDolares 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   2865
         TabIndex        =   4
         Top             =   1875
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   1635
         Width           =   1290
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   2865
         TabIndex        =   2
         Top             =   1635
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   4740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraFormaPago 
      BackColor       =   &H80000000&
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   0
      TabIndex        =   24
      Top             =   3360
      Width           =   6555
      Begin vbp_Ventas.ctlGrillaArray grdFormaPago 
         Height          =   1170
         Left            =   75
         TabIndex        =   25
         Top             =   210
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   2064
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Label Label18 
      Caption         =   "Zoom:"
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   615
   End
   Begin VB.Label LblTeleoperadores 
      Caption         =   "Asignación a Teleoperadora"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblDocumento 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1290
      Left            =   6600
      TabIndex        =   28
      Top             =   3480
      Width           =   1470
   End
End
Attribute VB_Name = "frm_VTA_Previa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCliente As New clsCliente
Dim objCovenio As New clsConvenio
Dim objLocal As New clsLocal


Dim odynCliente As oraDynaset
Public flgContinua As Boolean
Public pstrCodigoUsuario As String
Public pblnAsignoMeta As Boolean
Dim strEncontro As String
Public bolCancelCliente As Boolean

Private Sub Check3_Click()
    If Check3.Value = 1 Then objVenta.FlgUrgente = "1" Else objVenta.FlgUrgente = "0"
End Sub


Private Sub cmdGrabar_Click()
    'Validación de Linea Credito en Convenio por Delivery'
    'Hecho el 11/07/2007 por CRUEDA'
    
  objVenta.ObservacionLocal = txtObsLocal.Text
    
    
    If objVenta.Totales(2) < 0 Then
       MsgBox "El Total a pagar es menor al Monto Total", vbCritical, App.ProductName
       Exit Sub
    End If
    
    If objUsuario.EsDelivery Then
      If objVenta.ptmModalidad = Venta_Convenio Then
        If objVenta.flgBeneficiarios = "1" And objVenta.FlgValidaLinCre = "1" Then
            If val(objVenta.Totales(0)) > val(objVenta.LineaCred) Then
                MsgBox "Total de Compra al Credito Supera Linea de Credito del Cliente" & Chr(13) & _
                      objVenta.NombreBeneficiario & ", " & "Linea Credito =>" & objVenta.LineaCred & Chr(13) & _
                    "y su Compra es por =>" & " " & objVenta.Totales(0), vbCritical, Caption: Exit Sub
            End If
        End If
      End If
    End If
    If pblnAsignoMeta = True Then
      If CboTeleoperadores.BoundText = "*" Then MsgBox "Seleccione una Teleoperadora para la asignación", vbCritical, Caption: Exit Sub
       pstrCodigoUsuario = Trim(CboTeleoperadores.BoundText)
    End If
    Dim flgCapacidadDisp As String
    If gstrIndRAv3 = "2" Then
    flgCapacidadDisp = objVenta.validaCapacidad
    If flgCapacidadDisp <> "1" Then
'        MsgBox "La capacidad elegida ya no está" & vbNewLine & _
'               "disponible. Modifícala desde la" & vbNewLine & _
'               "opción Documentos en 'Datos de" & vbNewLine & _
'               "entrega (F11)'.", vbExclamation, App.ProductName
        
        frm_VTA_MetodosSegmentos.Parametro = objLocal.GetCodPosu(mdiPrincipal.ctlCliente1.LocalDespacho)
        frm_VTA_MetodosSegmentos.Tipo = 3
        frm_VTA_MetodosSegmentos.permiteCerrar = "1"
        frm_VTA_MetodosSegmentos.Show vbModal
        Me.grdMetodosEntrega.Rebind
        Exit Sub
    End If
    End If
    flgContinua = True
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    flgContinua = False
    Set objCliente = Nothing
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
On Error GoTo handle
    flgContinua = False
    FormatoGrilla
    lblPctCopago.Caption = objVenta.PctBeneficiario
    
    '*************************************************************************'
    ' Usuario que asigna meta en delivery '
    ' Fecha 19/10/2007
    ' Por Cristhian Rueda
    '*************************************************************************'
    strEncontro = Trim(objUsuario.AsignaMetaDLV(objUsuario.CodigoAplicacion, _
                                                objUsuario.CodigoMenuAsigna, _
                                                objUsuario.Codigo))
    If strEncontro = "1" Then
                                
            LblTeleoperadores.Visible = True: CboTeleoperadores.Visible = True
            
            pblnAsignoMeta = True
            Set CboTeleoperadores.RowSource = objUsuario.ListaUsuarioDLV
            CboTeleoperadores.ListField = "DES"
            CboTeleoperadores.BoundColumn = "COD"
            CboTeleoperadores.BoundText = "*"
      Else
            pblnAsignoMeta = False
    End If
    '*************************************************************************'
    If objVenta.ptmModalidad = Venta_Convenio Then
        If val(frmPedido.lblPctCopago.Caption) = 0 Then
            PicFpago.Visible = True
            Label4.Visible = False: lblPctCopago.Visible = False
            Label8.Visible = False: LblCoPago.Visible = False
        Else
            PicFpago.Visible = False
            Label4.Visible = True: lblPctCopago.Visible = True
            Label8.Visible = True: LblCoPago.Visible = True
        End If
    Else
        PicFpago.Visible = False
        Label4.Visible = False: lblPctCopago.Visible = False
        Label8.Visible = False: LblCoPago.Visible = False
    End If
    
    Call objVenta.OrdenarArregloFormaPago
    frm_VTA_FormaPago.GrdListaFP.Rebind
    
    grdFormaPago.Array1 = objVenta.FormaPago
    grdProductos.Array1 = objVenta.Producto
    grdMetodosEntrega.Array1 = objVenta.MetodoSegmento
    lblTotal.Caption = Format(objVenta.Totales(0), "###,##0.00") '+ dlbg
    lblRedondeo.Caption = objVenta.Totales(1)
    lblTotalPagar.Caption = objVenta.Totales(2) '- dlbg
    LblCoPago.Caption = objVenta.Totales(8)
    lblPagado.Caption = objVenta.Totales(6)
    lblVuelto.Caption = objVenta.Totales(7) '- dlbg
    lblTipoCambio.Caption = objVenta.Totales(10)
    lblTotalDolares.Caption = objVenta.Totales(11)
    '***** *****'
        
    txtObsCliente.Text = objVenta.ObservacionClienteDLV
    txtObsLocal.Text = objVenta.ObsNotaLocal
        
    '**
    If val(objVenta.FlgUrgente) = 1 Then Check3.Value = 1 Else Check3.Value = 0
'    If frm_VTA_Documento.blnTipoDoc = True Then
'        LblDocumento.Caption = frm_VTA_Documento.strDlvDocumento
'    Else
'        LblDocumento.Caption = "Usted Efectuara el Pago con" & "  " & IIf(Mid(objUsuario.TipDocDefault, 1, 3) = "BOL", "BOLETA", "FACTURA")
'    End If
    
    If Mid(frmPedido.lblSiguiente.Caption, 1, 3) = "BOL" Then
        LblDocumento.Caption = "Usted emitirá una BOLETA"
    Else
        LblDocumento.Caption = "Usted emitirá una FACTURA"
    End If
   
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Sub FormatoGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Código", "Descripción", "F", "Cantidad", "Precio", "T", "FlagFraccion", "Regalo", "TipDcto", "ImpDcto", "PrcAnt", "CodAutoriza", "CodUsuario", "Pct_Comi", "CtdProductoOrig", "FlgFraccionOrig", "Dato1", "Dato2", "Dato3", "Dato4", "Dato5", "FlgReceta")
    arrAncho = Array(0, 5294, 200, 600, 900, 190, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 0, 0, 0, 0, 0, 100)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgCenter, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral)
    
    grdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdProductos.Columns(1).FetchStyle = True
    grdProductos.Columns(7).FetchStyle = True
    grdProductos.Columns(1).WrapText = True
    grdProductos.Columns(5).Visible = False
    grdProductos.Columns(6).Visible = False
    grdProductos.Columns(7).Visible = False
    grdProductos.Columns(8).Visible = False
    grdProductos.Columns(9).Visible = False
    grdProductos.Columns(10).Visible = False
    grdProductos.Columns(11).Visible = False
    grdProductos.Columns(12).Visible = False
    grdProductos.Columns(13).Visible = False
    grdProductos.Columns(14).Visible = False
    grdProductos.Columns(15).Visible = False
    grdProductos.Columns(16).Visible = False
    grdProductos.Columns(17).Visible = False
    grdProductos.Columns(18).Visible = False
    grdProductos.Columns(19).Visible = False
    grdProductos.Columns(20).Visible = False
    grdProductos.Columns(21).Visible = False
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("localCode", "Tipo", "Día", "Valor", "Valor Fin", "Fecha", "Tiempo de Dlv", "Amount", "Segmento", "starHour", "endHour")
    arrAncho = Array(600, 700, 1300, 1100, 1100, 2300, 1500, 600, 2000, 1100, 1100)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    
    grdMetodosEntrega.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdMetodosEntrega.Columns("localCode").Visible = False
    grdMetodosEntrega.Columns("Valor").Visible = False
    grdMetodosEntrega.Columns("Valor Fin").Visible = False
    grdMetodosEntrega.Columns("Amount").Visible = False
    grdMetodosEntrega.Columns("starHour").Visible = False
    grdMetodosEntrega.Columns("endHour").Visible = False
        
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Código", "Forma Pago", "Código", "Pago", "Soles", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "Dolares")
    arrAncho = Array(0, 0, 0, 1600, 1100, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1100)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight)
    
    grdFormaPago.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdFormaPago.Columns(0).Visible = False: grdFormaPago.Columns(1).Visible = False
    grdFormaPago.Columns(2).Visible = False: grdFormaPago.Columns(5).Visible = False
    grdFormaPago.Columns(6).Visible = False: grdFormaPago.Columns(7).Visible = False
    grdFormaPago.Columns(8).Visible = False: grdFormaPago.Columns(9).Visible = False
    grdFormaPago.Columns(10).Visible = False: grdFormaPago.Columns(11).Visible = False
    grdFormaPago.Columns(12).Visible = False: grdFormaPago.Columns(13).Visible = False
    grdFormaPago.Columns(14).Visible = False: grdFormaPago.Columns(15).Visible = False
    grdFormaPago.Columns(16).Visible = False: grdFormaPago.Columns(17).Visible = False
    grdFormaPago.Columns(18).Visible = False: grdFormaPago.Columns(19).Visible = False
    grdFormaPago.Columns(20).Visible = False: grdFormaPago.Columns(21).Visible = False
    grdFormaPago.Columns(22).Visible = False: grdFormaPago.Columns(23).Visible = False
    grdFormaPago.Columns(24).Visible = False: grdFormaPago.Columns(25).Visible = False
    grdFormaPago.Columns(26).Visible = False: grdFormaPago.Columns(27).Visible = False
    grdFormaPago.Columns(28).Visible = False: grdFormaPago.Columns(29).Visible = False
    grdFormaPago.Columns(30).Visible = False: grdFormaPago.Columns(31).Visible = False
End Sub

Private Sub Frame3_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

