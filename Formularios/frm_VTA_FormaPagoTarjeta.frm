VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_FormaPagoTarjeta 
   BorderStyle     =   0  'None
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrilla grdTarjetas 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2143
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   3000
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   6735
      Begin vbp_Ventas.ctlDataCombo ctlCboTipoTarjeta 
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   2750
         _ExtentX        =   4842
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox txtNumDNI 
         Height          =   315
         Left            =   5500
         TabIndex        =   3
         Top             =   1545
         Visible         =   0   'False
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   556
         Tipo            =   3
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbp_Ventas.ctlTextBox txtNomTitular 
         Height          =   315
         Left            =   5500
         TabIndex        =   2
         Top             =   1110
         Visible         =   0   'False
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   556
         MaxLength       =   100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox ChkRetEfec 
         Alignment       =   1  'Right Justify
         Caption         =   "Retiro Efectivo"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2320
         Width           =   1335
      End
      Begin vbp_Ventas.ctlDataCombo ctLCboTipoCuota 
         Height          =   315
         Left            =   5500
         TabIndex        =   6
         Top             =   2415
         Visible         =   0   'False
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin MSMask.MaskEdBox mskVencimiento 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   640
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin vbp_Ventas.ctlTextBox txtNroCuota 
         Height          =   315
         Left            =   5000
         TabIndex        =   5
         Top             =   2415
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbp_Ventas.ctlTextBox txtNroAut 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   1110
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbp_Ventas.ctlTextBox txtImporte 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   1545
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbp_Ventas.ctlTextBox TxtRetiro 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   2415
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbp_Ventas.ctlTextBox txtNroTar 
         Height          =   315
         Left            =   5500
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   556
         Tipo            =   3
         Enabled         =   0   'False
         MaxLength       =   21
         TABAuto         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Retiro efectivo (F2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3000
         TabIndex        =   23
         Top             =   2415
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Importe S/. : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1545
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Vencimiento : "
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   640
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Tarjeta : "
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Autorización : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label LblNomTraj 
         BackColor       =   &H00DBFBFA&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5500
         TabIndex        =   17
         Top             =   640
         Visible         =   0   'False
         Width           =   1000
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_FormaPagoTarjeta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FormaPagoTarjeta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Supr] - Eliminar Tarjeta del  Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   24
      Top             =   360
      Width           =   3045
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F1] -Tarjetas disponibles del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   11
      Left            =   6090
      TabIndex        =   15
      Top             =   7200
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   12
      Left            =   4380
      TabIndex        =   14
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoTarjeta.frx":0B14
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Tarjeta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   4
      Left            =   420
      TabIndex        =   13
      Top             =   60
      Width           =   2535
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objTarjeta As New clsFormaPago
Dim odynR1 As oraDynaset
Dim strCodTarj As String
Dim strNomTarj As String
Dim strTarjeta As String
Dim dblTotal As Double
Dim strVencMes As String
Dim strVencAño As String
Dim FchVenc As String
Dim strMoneda As String
Public pdblRetEfect As String
Dim strValorTarj As String
''nuevas variables
Public pstrDato As String
Public pstrDatoDes As String
Public bolCancelar As Boolean
Public pblnTarjNew As Boolean

Dim intLongTarj As Integer    ' => Optiene la longitud de la tarjeta que esta en BTLPROD.PKG_CONSTANTES.CONS_LONG_TARJ
Dim intLongLeidoTjt As Integer
Private intCount As Integer
Private intentos As Integer
'Public PNombreTitular As String
'Public PNumDni As String

Private Sub ChkRetEfec_Click()
On Error GoTo Handle
    If mskVencimiento.Text = "" Then Exit Sub
    If (txtNroCuota.Text = "") Then Exit Sub
    If txtNroAut.Text = "" Then Exit Sub
    If (txtImporte.Text = "" Or txtImporte.Text <= 0) Then Exit Sub
    If ChkRetEfec.Value = "1" Then
        TxtRetiro.Visible = True
      Else
        TxtRetiro.Text = ""
        TxtRetiro.Visible = False
    End If
    '''SendKeys "{TAB}"
    'frm_Retiro_Efectivo.Show vbModal
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub



Private Sub ctLCboTipoCuota_Change()
    If ctLCboTipoCuota.BoundText = "*" Then Exit Sub
End Sub

Private Sub Form_Activate()
    ''Carga_Inicial
     On Error GoTo Handle
     
    'If grdTarjetas.ApproxCount = 0 Then txtNroTar.SetFocus
    If grdTarjetas.ApproxCount > 0 Then grdTarjetas.SetFocus
    If objUsuario.EsDelivery Then
        Label3.Visible = False
'        Label9.Visible = True
'        Label10.Visible = True
        txtNomTitular.Visible = False
        txtNumDNI.Visible = False
        mskVencimiento.Visible = False
'        Label4.top = 2445
        Label6.top = 1140
        Label5.top = 1545
        ChkRetEfec.top = 2010
        Label8.top = 2325
        txtNroCuota.top = 2475
        ctLCboTipoCuota.top = 1110
        txtNroAut.top = 1110
        txtImporte.top = 1545
        TxtRetiro.top = 1980 '3720
      Else
        Label3.Visible = False
        mskVencimiento.Visible = False
'        Label9.Visible = False
'        Label10.Visible = False
        txtNomTitular.Visible = False
        txtNumDNI.Visible = False
'        Label4.top = 1140
        Label6.top = 1140
        Label5.top = 1545
        ChkRetEfec.top = 2010
        Label8.top = 2325
        txtNroCuota.top = 2475
        ctLCboTipoCuota.top = 1110
        txtNroAut.top = 1110
        txtImporte.top = 1545
        TxtRetiro.top = 1980 '2415
    End If
    
   If gintFidelizado = 1 Then
    'frmPedido.optCredito.Value = True
    If Format(frmPedido.lblTotal, "0.00") <> "0.00" Then
        txtImporte.Text = Format(frmPedido.lblTotal, "0.00")
        'txtImporte.Enabled = False
    End If
    txtNroAut.SetFocus
End If
    Exit Sub
Handle:
            MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Me.top = 0
    Me.left = 0
    setteaFormulario Me
    Carga_Inicial
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Carga_Inicial()
    Set odynR1 = objTarjeta.ListaHijo(pstrDato)
    strMoneda = "" & odynR1("COD_MONEDA").Value
    Set ctLCboTipoCuota.RowSource = objTarjeta.TipoCuota
    ctLCboTipoCuota.ListField = "DES"
    ctLCboTipoCuota.BoundColumn = "COD"
    ctLCboTipoCuota.BoundText = "1"
    ChkRetEfec.Value = 0
    If objVenta.CodModalidadVenta = Venta_Convenio Then
        txtImporte.Text = Format(frmPedido.lblcopago, "0.00")
    Else
        txtImporte.Text = Format(frmPedido.lblTotal, "0.00")
    End If
    txtImporte.Enabled = True
    txtNroCuota.Text = "0"
    LblNomTraj.Caption = ""
    txtNroTar.Text = ""
    Format_Grilla
    If objUsuario.EsDelivery Then
        Set grdTarjetas.DataSource = objTarjeta.ListaTJTxdlv(objVenta.CodigoCliente)
        If grdTarjetas.ApproxCount > 0 Then pblnTarjNew = False Else pblnTarjNew = True
        mskVencimiento.Text = "__/____"
        ChkRetEfec.Visible = False: TxtRetiro.Visible = False: Label8.Visible = False
        Label6.Visible = objUsuario.flgDeliveryProv = "1"
        txtNroAut.Visible = objUsuario.flgDeliveryProv = "1"
      Else
        ConsultaTarjetasDisponibles
        Label6.Visible = True
        txtNroAut.Visible = True
    End If
    
    intLongTarj = objVenta.LongTarj
    Set ctlCboTipoTarjeta.RowSource = objTarjeta.ListaTarjetasDlv
    ctlCboTipoTarjeta.ListField = "DES_HIJO"
    ctlCboTipoTarjeta.BoundColumn = "COD_HIJO"
    ctlCboTipoTarjeta.BoundText = "*"
    txtNroCuota.Text = "1"
    intentos = 0
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Control
    'If ctLCboTipoCuota.BoundText = "*" Then MsgBox "Seleccione tipo de cuota", vbCritical, Caption: Exit Sub
    'txtNroCuota_Change
    intentos = intentos + 1
    If intentos > 1 Then Exit Sub
    If ctlCboTipoTarjeta.BoundText = "*" Then
        txtNroTar.Text = ""
        LblNomTraj.Caption = ""
        MsgBox "Seleccione el tipo de Tarjeta", vbCritical, App.ProductName: ctlCboTipoTarjeta.SetFocus: Exit Sub
        intentos = 0
        Exit Sub
    Else
        'I.ECASTILLO 07.10.2020
        'If ctlCboTipoTarjeta.BoundText = "003" Then txtNroTar.Text = "4100000000000000"
        'If ctlCboTipoTarjeta.BoundText = "004" Then txtNroTar.Text = "5100000000000000"
        txtNroTar.Text = "" & gclsOracle.FN_Valor("BTLPROD.PKG_FORMA_PAGO.FN_GET_NUM_TARJETA", ctlCboTipoTarjeta.BoundText)
        'F.ECASTILLO 07.10.2020
        txtNroTar_KeyPress 13
    End If
    
    Dim strMensajeBin As String
    strMensajeBin = "" & objTarjeta.TarjetaBloqueada(txtNroTar.Text)
    If Not strMensajeBin = "" Then
       MsgBox strMensajeBin, vbCritical, App.ProductName
       intentos = 0
       Exit Sub
    End If
    
    
    If ValidaTarjeta = False Then Exit Sub
    If txtNroCuota.Text < 0 Then MsgBox "El Nº de cuota no puede ser cero", vbCritical, App.ProductName: txtNroCuota.SetFocus: intentos = 0: Exit Sub
    If objUsuario.EsDelivery = False Then
        If txtNroAut.Text = "" Then MsgBox "Ingrese Codigo de Autorizacion", vbCritical, App.ProductName: txtNroAut.SetFocus: intentos = 0: Exit Sub
    End If
    If (txtImporte.Text = "" Or Format(txtImporte.Text, "0.00") = "0.00") And objVenta.CodModalidadVenta <> Venta_Convenio Then MsgBox "El Importe no puede ser cero", vbCritical, App.ProductName: intentos = 0: txtImporte.SetFocus: Exit Sub
    
    If objUsuario.EsDelivery Then
' 21/11/2019
' EL NUEVO REQUERIMIENTO INDICA QUE LA FECHA DE TARJETA SEA AUTOMATICO
'        strVencMes = Mid(mskVencimiento.Text, 1, 2)
'        strVencAño = Mid(mskVencimiento.Text, 4, 4)
'
'        If strVencMes = "00" Then MsgBox "Error en el ingreso del mes", vbCritical, App.ProductName: mskVencimiento.SetFocus: Exit Sub
'        If Len(strVencMes & strVencAño) < 6 Then MsgBox "Error al digitar la fecha de vencimiento", vbCritical, App.ProductName: mskVencimiento.SetFocus: Exit Sub
'        If InStr(strVencMes, "_") > 0 Then MsgBox "Error en el dia de la fecha", vbCritical, App.ProductName: mskVencimiento.SetFocus: Exit Sub
'        If InStr(strVencAño, "_") > 0 Then MsgBox "Error en el año de la fecha", vbCritical, App.ProductName: mskVencimiento.SetFocus: Exit Sub
'
'        If strVencMes > 12 Then
'            MsgBox "El mes ingresado no es valido", vbCritical, App.ProductName
'            mskVencimiento.SetFocus
'            Exit Sub
'        End If
'        'cambio pherrera 050109 no valida bien, cuando la fecha de vencimiento era 2008 en 2009 no valida por que el mes de venc era 12
'        'If (strVencMes < Format(objUsuario.sysdate, "mm")) And (strVencAño <= Format(objUsuario.sysdate, "yyyy")) Then
'        If ((strVencMes < Format(objUsuario.sysdate, "mm")) And (strVencAño <= Format(objUsuario.sysdate, "yyyy"))) Or (strVencAño < Format(objUsuario.sysdate, "yyyy")) Then
'            MsgBox "La tarjeta Ingresada esta vencida", vbCritical, App.ProductName
'            mskVencimiento.SetFocus
'            Exit Sub
'        End If
'        FchVenc = DateSerial(year(Format("01/" & mskVencimiento.Text, "dd/mm/yyyy")), month(Format("01/" & mskVencimiento.Text, "dd/mm/yyyy")) + 1, 0)
        FchVenc = Format(DateAdd("m", 1, Now), "dd/MM/yyyy")
    End If
    
    objVenta.NomTitular = txtNomTitular.Text
    objVenta.NumDNI = txtNumDNI.Text

    objVenta.AgregaFormaPago pstrDato, _
                             pstrDatoDes, _
                             strCodTarj, _
                             strNomTarj, _
                             dblTotal, _
                             strCodTarj, _
                             strMoneda, "", _
                             "", "", _
                             "", 0, _
                             txtNroTar.Text, _
                             txtNroCuota.Text, _
                             FchVenc, _
                             ctLCboTipoCuota.BoundText, _
                             "", "", _
                             "", "", _
                             txtNroAut.Text, "", _
                             "", "", _
                             "", IIf(Trim(objVenta.NumDNI) = "", "", Trim(objVenta.NumDNI)), _
                             "", "", _
                             "", IIf(Trim(TxtRetiro.Text) = "", "0", Trim(TxtRetiro.Text)), _
                             "", "", "", IIf(Trim(objVenta.NomTitular) = "", "", Trim(objVenta.NomTitular))
                             
                             'txtNroTar.Text , _

    frmPedido.Cal_Promo
    Unload Me
    '***************************************'
    'Arma el arreglo cada ez que se modifica'
      frm_VTA_FormaPago.Show
      'frm_VTA_FormaPago.GrdListaFP.Array = objVenta.FormaPago
      frm_VTA_FormaPago.GrdListaFP.Rebind
    '***************************************'
    frmPedido.Cal_Montos
Exit Sub
Control:
    intentos = 0
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub cmdCancelar_Click()
Cancelar
'frmPedido.flgF6 = 0
frmPedido.OptionsFocus
End Sub

Public Sub Cancelar()
On Error GoTo CtrlErr
If Val(txtImporte.Text) = 0 And LblNomTraj.Caption <> "" Then
    objVenta.RemoverFormaPago pstrDato, strCodTarj, txtNroTar.Text
    frm_VTA_FormaPago.GrdListaFP.Refresh
    ''''''''frm_VTA_FormaPago.GrdListaFP.Delete
    frmPedido.Cal_Promo
    frmPedido.Cal_Montos
End If


Unload Me
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
'Private Sub ctLCboTipoCuota_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      SendKeys "{TAB}"
'   End If
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyF1
            grdTarjetas.SetFocus
            pblnTarjNew = True
        Case vbKeyF2
            If ChkRetEfec.Value = 0 Then
                ChkRetEfec.Value = 1
                If TxtRetiro.Visible = True Then TxtRetiro.SetFocus
            Else
                ChkRetEfec.Value = 0
            End If
            
            
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyEscape
            cmdCancelar_Click
        'Case vbKeyF8
    End Select
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName

End Sub

Private Sub grdTarjetas_Click()
    pblnTarjNew = True
    'grdTarjetas.SetFocus
End Sub

Private Sub grdTarjetas_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strMensaje As String
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
    If KeyCode = vbKeyDelete Then
        If grdTarjetas.ApproxCount = 0 Then Exit Sub
        If MsgBox("Desea eliminar tarjeta ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
            strMensaje = objTarjeta.GrabaTarjetasXDLV(objVenta.CodigoCliente, _
                                    grdTarjetas.Columns("NÚMERO").Value, _
                                    grdTarjetas.Columns("COD_TIPO_TARJETA").Value, _
                                    grdTarjetas.Columns("F. VENC.").Value, _
                                    "0")
            If strMensaje = "" Then
              MsgBox "Se realizo la anulación satisfactoriamente ", vbExclamation, App.ProductName
            Else
              MsgBox strMensaje, vbCritical, App.ProductName
            End If
            Carga_Inicial
        
        
        End If
        
    End If
End Sub

Private Sub grdTarjetas_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdTarjetas.ApproxCount > 0 Then
      If pblnTarjNew = True Then
        LblNomTraj.Caption = grdTarjetas.DataSource(1).Value
        txtNroTar.Text = grdTarjetas.Columns("NÚMERO").Value
        mskVencimiento.Text = Format(grdTarjetas.Columns(2), "mm/yyyy")
        txtNomTitular.Text = grdTarjetas.Columns("NOM_TITULAR").Value
        txtNumDNI.Text = grdTarjetas.Columns("NUM_DNI").Value
        '''txtNroCuota.SetFocus
      End If
    End If
End Sub

Private Sub mskVencimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub mskVencimiento_Validate(Cancel As Boolean)
'    Cancel = Not fbln_Valida_Fecha("MM/yyyy", "Error en fecha de vencimiento", mskVencimiento.Text)
'    If Cancel Then
'        MsgBox "Error en fecha de vencimiento", vbExclamation, Caption
'        mskVencimiento.SetFocus
'    End If
End Sub

Private Sub txtImporte_Change()
    dblTotal = Val(txtImporte.Text)
End Sub

Private Sub txtImporte_GotFocus()
    cmdAceptar.Default = True
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    txtImporte.Tipo = Real
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtImporte_LostFocus()
    cmdAceptar.Default = False
End Sub

Private Sub txtNomTitular_KeyPress(KeyAscii As Integer)
    txtNomTitular.Tipo = Mayusculas
End Sub

Private Sub txtNroCuota_Change()
'    If txtNroTar.Text = "" Then MsgBox "Ingrese antes la tarjeta", vbCritical, Caption: Exit Sub
'    If Not IsNumeric(txtNroCuota.Text) Then MsgBox "Error Dato debe ser numerico", vbCritical, Caption: txtNroCuota.selection: Exit Sub
'    If (txtNroCuota.Text = "" Or txtNroCuota.Text = 0) Then MsgBox "La cuota tiene que ser mayor a cero", vbCritical, Caption: Exit Sub
End Sub

Private Sub txtNroTar_GotFocus()
    intCount = 0
    
End Sub

''''Private Sub txtNroCuota_KeyPress(KeyAscii As Integer)
''''    txtNroAut.Tipo = Entero
''''    If KeyAscii = 13 Then
''''        'ctLCboTipoCuota.TabIndex = 3
''''        SendKeys "{TAB}"
''''    End If
''''End Sub

Public Sub txtNroTar_KeyPress(KeyAscii As Integer)
    On Error GoTo CtrlErr
    If intCount > 1 And KeyAscii <> 13 Then KeyAscii = 0
    'txtNroTar.Tipo = Entero
    'Debug.Print Chr(KeyAscii) & " " & KeyAscii
    If intentos > 1 Then Exit Sub
    If Chr(KeyAscii) = "&" Then intCount = intCount + 1
    If KeyAscii = 13 Then
        Dim strMen As String
        strMen = "" & objTarjeta.TarjetaBloqueada(txtNroTar.Text)
        If Not strMen = "" Then
            MsgBox strMen, vbCritical, App.ProductName
            intentos = 0
            Exit Sub
        End If
'oK:
        If ValidaTarjeta Then
                 gintFidelizado = 0
                 If objVenta.ValidaTarjetaCMR(txtNroTar.Text) > 0 Then
                        frmPedido.flgF6 = 0
                        AgregaTarjetaCMR (txtNroTar.Text)
                 End If
                 
                
                objVenta.AgregaFormaPago pstrDato, _
                                         pstrDatoDes, _
                                         strCodTarj, _
                                         strNomTarj, _
                                         dblTotal, _
                                         strCodTarj, _
                                         strMoneda, "", _
                                         "", "", _
                                         "", 0, _
                                         txtNroTar.Text, _
                                         txtNroCuota.Text, _
                                         FchVenc, _
                                         ctLCboTipoCuota.BoundText, _
                                         "", "", _
                                         "", "", _
                                         txtNroAut.Text, "", _
                                         "", "", _
                                         "", "", _
                                         "", "", _
                                         "", IIf(Trim(TxtRetiro.Text) = "", "0", Trim(TxtRetiro.Text))
            
                   frmPedido.Cal_Promo
                   frmPedido.Cal_Montos
                   
                   
                   
            txtNroTar.TABAuto = True
            
            If gintFidelizado = 1 Then '
                If Format(frmPedido.lblTotal, "0.00") <> "0.00" Then
                    txtImporte.Text = Format(frmPedido.lblTotal, "0.00")
                    'txtImporte.Enabled = False
                End If
                txtNroAut.SetFocus
            Else
                'txtImporte.Enabled = True
                'txtImporte.Text = Format(frmPedido.lblTotal, "0.00")
'                If objVenta.CodModalidadVenta = Venta_Convenio Then
'                    txtImporte.Text = Format(dblTotal, "0.00")
'                Else
'                    txtImporte.Text = Format(frmPedido.lblTotal, "0.00")
'                End If
'                txtImporte.Text = IIf(Format(dblTotal, "0.00") <> "0.00", Format(dblTotal, "0.00"), Format(frmPedido.lblTotal, "0.00"))
            End If
            
        End If
        
    End If
    
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtNroTar_LostFocus()
txtNroTar.TABAuto = False
End Sub

Private Sub TxtRetiro_GotFocus()
    cmdAceptar.Default = True
End Sub

Private Sub TxtRetiro_KeyPress(KeyAscii As Integer)
    TxtRetiro.Tipo = Real
End Sub
Sub ConsultaTarjetasDisponibles()
If Not objVenta.CodigoCliente = "" Then
    Dim objCliente As New clsCliente
        Set grdTarjetas.DataSource = objCliente.ListaTarjetas(objVenta.CodigoCliente)
        grdTarjetas.Rebind
    Set objCliente = Nothing
End If
End Sub

Private Function ValidaTarjeta() As Boolean
If txtNroTar.Text = "" Or ctlCboTipoTarjeta.BoundText = "*" Then
                'MsgBox "EL número de tarjeta es incorrecto", vbCritical, App.ProductName
                MsgBox "Seleccione tipo de tarjeta", vbCritical, App.ProductName
                'txtNroTar.selection
                LblNomTraj.Caption = ""
                ValidaTarjeta = False
                intentos = 0
                Exit Function
End If
         If objTarjeta.ValidaVoucher(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.NombrePC, objUsuario.Codigo, Trim(txtNroTar.Text), Trim(txtNroAut.Text)) > 0 Then
                MsgBox "El voucher ya ha sido registrado", vbCritical, App.ProductName
                ValidaTarjeta = False
                intentos = 0
                Exit Function
         End If



        strValorTarj = Trim(txtNroTar.Text)
        'strValorTarj = Mid(txtNroTar.Text, 1, intLongTarj)
        txtNroTar.Text = Mid(strValorTarj, 1, intLongTarj)
        
        strCodTarj = objTarjeta.ValidaTarjeta(txtNroTar.Text, "1")
        strNomTarj = objTarjeta.ValidaTarjeta(txtNroTar.Text, "2")
        strTarjeta = objTarjeta.ValidaTarjeta(txtNroTar.Text, "3")
        'If strNomTarj = "*" Then MsgBox "Numero Tarjeta Ingresada no Valida", vbCritical, App.ProductName: txtNroTar.selection: LblNomTraj.Caption = "": ValidaTarjeta = False: Exit Function
        If strNomTarj = "*" Then
            If objTarjeta.ValidaMOD10(txtNroTar.Text, "0") = 1 Then
                If MsgBox("El número de tarjeta no tiene ninguna coincidencía, ¿Desea registrarlo?", vbYesNo + vbInformation, App.ProductName) = vbNo Then
                    txtNroTar.selection
                    LblNomTraj.Caption = ""
                    ValidaTarjeta = False
                    intentos = 0
                    Exit Function
                Else
                    frm_VTA_Agrega_BIN.pbNumeroTarjeta = ""
                    frm_VTA_Agrega_BIN.pbNumeroTarjeta = txtNroTar.Text
                    frm_VTA_Agrega_BIN.Show vbModal
                    If frm_VTA_Agrega_BIN.pbNumeroTarjeta = "" Then
                        txtNroTar.selection
                        LblNomTraj.Caption = ""
                        ValidaTarjeta = False
                        intentos = 0
                        Exit Function
                    End If
                    ValidaTarjeta
                End If
            Else
                MsgBox "EL número de tarjeta es incorrecto", vbCritical, App.ProductName
                txtNroTar.selection
                LblNomTraj.Caption = ""
                ValidaTarjeta = False
                intentos = 0
                Exit Function
            End If
            
        End If
        
        LblNomTraj.Caption = strNomTarj
        
        If strValorTarj = "" Then intentos = 0: Exit Function
          If objUsuario.EsDelivery Then
            If Len(strValorTarj) > intLongTarj Then
              mskVencimiento.Text = Mid(strValorTarj, 19, 2) & "/" & "20" & Mid(strValorTarj, 17, 2)
            End If
          End If
        
        If IIf(strTarjeta = "", "0", "1") Then
            ChkRetEfec.Enabled = True
          Else
            ChkRetEfec.Enabled = False
        End If
        
        ValidaTarjeta = True
End Function

Private Sub TxtRetiro_LostFocus()
    cmdAceptar.Default = False
End Sub


Sub Format_Grilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
           arrCampos = Array("NUMERO", "DESCRIPCION", "F. VENC.", "COD_TIPO_TARJETA", "NOM_TITULAR", "NUM_DNI")
           arrCaption = Array("Número", "Descripción", "F. Venc.", "Tipo", "Titular", "DNI")
           arrAncho = Array(2000, 2500, 1500, 0, 2000, 1500)
           arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgLeft, dbgRight, dbgCenter)
           grdTarjetas.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
           grdTarjetas.Columns(0).EditMask = "####-####-####-####"
           grdTarjetas.Columns(0).NumberFormat = "Edit Mask"

           grdTarjetas.Columns(3).Visible = False
           
End Sub

Function AgregaTarjetaCMR(ByVal Texto As String)

Dim rs As oraDynaset
Dim objTarCMR As New clsClienteD

Set rs = objTarCMR.ListaClienteCMR(Texto)
frmPedido.pstrDniCli = ""
If rs.RecordCount > 0 Then
    If frmPedido.pstrDniCli <> "" Then
       If frmPedido.pstrDniCli <> rs("NUM_DOCUMENTO_ID") Then
           MsgBox "El DNI Fidelizado es Diferente al DNI que pertenece la Tarjeta, se remeplazará por este", vbCritical
       End If
    End If
   frmPedido_Busca_Cli.ctlTxtDNI.Text = rs("NUM_DOCUMENTO_ID")
   frmPedido_Busca_Cli.ctlTxtDNI_KeyPress (13)
   gintFidelizado = 1
Else
    If MsgBox("¿ Desea Inscribirse al Club de Descuentos CMR.. ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        gintFidelizado = 1
        frmPedido_Busca_Cli.Show vbModal
        If frmPedido.pstrDniCli <> "" Then
            Dim strMensaje As String
            strMensaje = objTarCMR.GrabarTarjetaCMR(objVenta.CodigoCliente, Texto, objUsuario.Codigo)
            If strMensaje = "" Then
            Else
                intentos = 0
                MsgBox strMensaje, vbCritical, App.ProductName
            End If
        End If
    Else
        gintFidelizado = 0
    End If
End If
 frmPedido.optCredito.Value = True
End Function




