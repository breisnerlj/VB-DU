VERSION 5.00
Begin VB.Form frm_VTA_PreviaTomaPedido 
   Caption         =   "Verifica Cliente"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCliente 
      Caption         =   "&Cliente"
      Height          =   615
      Left            =   5280
      Picture         =   "frm_VTA_PreviaTomaPedido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdTomaPedido 
      Caption         =   "&Pedido"
      Height          =   615
      Left            =   6720
      Picture         =   "frm_VTA_PreviaTomaPedido.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Frame FrameUbicacion 
      Caption         =   "Ubicación"
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
      Height          =   5775
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   8175
      Begin VB.TextBox txtUbicacion 
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   6375
      End
      Begin VB.CommandButton BtnBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   6720
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbZoom 
         Height          =   315
         Left            =   720
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmbMapType 
         Height          =   315
         Left            =   3720
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin vbp_Ventas.ucGMap ucGMap1 
         Height          =   3975
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7011
      End
      Begin VB.Label sdsd 
         Caption         =   "Zoom:"
         Height          =   255
         Left            =   195
         TabIndex        =   24
         Top             =   765
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   765
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dar doble clic sobre el mapa para establecer la ubicación de entrega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   4995
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
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
      Height          =   2265
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin vbp_Ventas.ctlTextBox txtRUC 
         Height          =   315
         Left            =   6540
         TabIndex        =   2
         Top             =   1860
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         ColorDefault    =   -2147483634
         ColorDefault    =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtRazonSocial 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1860
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         ColorDefault    =   -2147483634
         ColorDefault    =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtDistrito 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   888
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
         ColorDefault    =   -2147483634
         ColorDefault    =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtUrbanizacion 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   1536
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
         ColorDefault    =   -2147483634
         ColorDefault    =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox ctlTextBox1 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   564
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
         ColorDefault    =   -2147483634
         ColorDefault    =   -2147483634
         TABAuto         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox ctlTextBox2 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
         ColorDefault    =   -2147483634
         ColorDefault    =   -2147483634
         TABAuto         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox ctlTextBox3 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1212
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
         ColorDefault    =   -2147483634
         ColorDefault    =   -2147483634
         TABAuto         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   624
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1242
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   948
         Width           =   480
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Urbanización"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   1596
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   945
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "RUC"
         Height          =   195
         Left            =   5700
         TabIndex        =   9
         Top             =   1920
         Width           =   345
      End
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   8880
      Width           =   4695
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   8520
      Width           =   4695
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   8160
      Width           =   4815
   End
End
Attribute VB_Name = "frm_VTA_PreviaTomaPedido"
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
Dim maxZoom As Long
Dim flgZoom As Boolean
Dim flgNoInfo As Boolean
Public flgMapa As Boolean
Public strDireccion_old As String
Private objWS As New clsWebService


Private Sub BtnBuscar_Click()
    Debug.Print "CD"
    ucGMap1.countCalls "BtnBuscar_Click"
    If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido.BtnBuscar.click"
    txtSearch_KeyDown 13, 0
End Sub

Private Sub cmdCliente_Click()
On Error GoTo CtrErr
    Dim strAdic As String
    Dim valExisteComa As Integer
    
    Dim strSuf As String
    Dim valI As Integer
    Dim ind As Integer
    Dim strSufijoDir As String
    Dim indExistePunto As Integer
    Dim indExisteSufijo As Integer
    Dim strSufijo As String
    Dim strDireccion As String
    Dim dataUbigeo As oraDynaset
    bolCancelCliente = True
    ctlTextBox2.Text = ""
    ctlTextBox1.Text = ""
    ctlTextBox3.Text = ""
    txtSearch.Text = ""
    frm_VTA_Cliente.Telefono = objVenta.DesAuxCliTlf
    frm_VTA_Cliente.ctlCliente1.XTipoFuncion = "Editar"
    frm_VTA_Cliente.ctlCliente1.CodDireccionCli = objVenta.CodDireccionCli
    '''frm_VTA_Cliente.ctlCliente1.ConsultaCliente objVenta.CodigoClienteDLV
    frm_VTA_Cliente.strCodigo = objVenta.CodigoClienteDLV
    frm_VTA_Cliente.CargarValores
    frm_VTA_Cliente.Show vbModal
    objVenta.CodigoCliente = IIf(objVenta.CodigoConvenio = "", objVenta.CodigoClienteDLV, objVenta.CodigoBeneficiario) ''parche arturo 09/06/2009
    objVenta.bk_codBeneficiario = IIf(objVenta.CodigoConvenio = "", "", objVenta.CodigoBeneficiario) 'ECASTILLO 17.12.2020
    If bolCancelCliente = False Then
        Set odynCliente = objCliente.Lista(objVenta.CodigoClienteDLV)
        ctlTextBox2.Text = "" & odynCliente("XNOMBRE").Value
        ctlTextBox1.Text = "" & odynCliente("DES_DIRECCION_SOCIAL").Value
        ctlTextBox3.Text = "" & odynCliente("DES_REFERENCIA").Value
    '    objVenta.Latitud = "" & odynCliente("LATITUD").Value
      '  objVenta.Longitud = "" & odynCliente("LONGITUD").Value
        If gstrIndRAv4 = "1" Then
            objVenta.bk_Ubigeo = "" & odynCliente("UBIGEO").Value
            Set dataUbigeo = objCliente.getUbigeoDesc(objVenta.bk_Ubigeo)
            objVenta.dc_departamentBK = ""
            If Not dataUbigeo.EOF Then
                objVenta.dc_departamentBK = "" & dataUbigeo("DES_DEPARTAMENTO").Value
            End If
        End If
        'objVenta.dc_departamentBK = "" & odynCliente("DES_DEPARTAMENTO").Value
        strAdic = "": valExisteComa = 0
        valExisteComa = InStr(1, odynCliente("DES_DIRECCION_SOCIAL"), ",", vbTextCompare)
        strAdic = IIf(valExisteComa > 0, "", ", " + txtDistrito.Text)
        txtSearch.Text = odynCliente("DES_DIRECCION_SOCIAL") + strAdic
        indExistePunto = InStr(1, objVenta.bk_SufijoDir, ".")
        If indExistePunto > 0 Then indExisteSufijo = InStr(1, UCase(txtSearch.Text), UCase(objVenta.bk_SufijoDir))
        If indExisteSufijo > 0 Then
            strDireccion = txtSearch.Text
        Else
            If Len(objVenta.bk_SufijoDir) > 0 Then strSufijo = Mid(objVenta.bk_SufijoDir, 1, indExistePunto - 1)
            indExisteSufijo = InStr(1, UCase(txtSearch.Text), UCase(strSufijo) & " ")
            If indExisteSufijo > 0 Then
                strDireccion = txtSearch.Text
            Else
                strDireccion = objVenta.bk_SufijoDir + " " + txtSearch.Text
            End If
        End If
        txtSearch.Text = strDireccion
        
        'valI = InStr(1, txtSearch.Text, strSufijoDir)
        'If valI = 0 Then txtSearch.Text = strSufijoDir + " " + txtSearch.Text
        'txtSearch.Text = odynCliente("DES_DIRECCION_SOCIAL") + " " + txtDistrito.Text
    End If
    Exit Sub
CtrErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdTomaPedido_Click()
    If strDireccion_old <> txtSearch.Text Then
        MsgBox "Se ha cambiado la dirección sin realizar la busqueda, favor de hacer click en 'Buscar'", vbCritical, App.ProductName
        Exit Sub
    End If
    If objVenta.CodigoTipoVenta = 1 Or objVenta.CodigoTipoVenta = 2 Then
        If ucGMap1.Markers.Count = 0 Or objVenta.Latitud = "" Or objVenta.Longitud = "" Then
            MsgBox "Debe seleccionar un punto de entrega", vbCritical, Caption:
            ucGMap1.countCalls "cmdTomaPedido_Click"
            If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido.cmdTomaPedido.click; Lat,Lng nulos; sistema fuerza busqueda"
            txtSearch_KeyDown 13, 0
            Exit Sub
        End If
    End If
    'flgContinua = True
    'I.ECASTILLO 27.10.2020
    '17.12.2020 | se comenta función, volver a usar
    objVenta.isLocalDcCappa = "" & objLocal.GetIndCDCAP(mdiPrincipal.ctlCliente1.LocalDespacho)
    'F.ECASTILLO 27.10.2020
    
    'I.ECASTILLO 17.12.2020 | 2da etapa reserva | 06.01.2021
    Dim flg_ruteoA_cnv
    flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto
    End If
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    Dim sCia As String
    Dim rsCia As oraDynaset
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, mdiPrincipal.ctlCliente1.LocalDespacho)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    Dim flgFunLocal As String
    flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "1" And flgFunLocal = "1" Then
        GoTo CallServiceReserva
    End If
    objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, mdiPrincipal.ctlCliente1.LocalDespacho, "RESERVA_STOCK_2DA")
    If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
cnvNoRuteaAuto:
        flgContinua = True
    Else
CallServiceReserva:
'        If objVenta.isLocalDcCappa = "0" Then
            If Len(Trim(objVenta.Latitud)) = 0 Or _
                (ucGMap1.GPoint = "20.703879,-40.993700") Then
                If MsgBox("No tiene geolocalización valida desea continuar?", _
                        vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbNo Then
                        Exit Sub
                End If
                frm_VTA_MetodosSegmentos.Parametro = objVenta.bk_Ubigeo
                frm_VTA_MetodosSegmentos.Tipo = 1
            Else
                frm_VTA_MetodosSegmentos.Parametro = objVenta.Latitud & "," & objVenta.Longitud
                frm_VTA_MetodosSegmentos.Tipo = 2
            End If
            frm_VTA_MetodosSegmentos.permiteCerrar = "0"
            frm_VTA_MetodosSegmentos.Show vbModal
'        Else
'            flgContinua = True
'        End If
    End If
    'F.ECASTILLO 12.2020
    If gstrIndRAv4 = "1" Then
        objVenta.dc_street = txtUbicacion.Text 'txtSearch.Text
    Else
        objVenta.dc_street = txtSearch.Text
    End If
        
    objVenta.DireccionCliente = objVenta.dc_street 'txtSearch.Text
    objVenta.DireccionClienteDLV = objVenta.dc_street 'txtSearch.Text
    objVenta.Out_Direccion = objVenta.DireccionClienteDLV
    strDireccion_old = ""
    objVenta.bk_codCliente = objVenta.CodigoClienteDLV
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    mdiPrincipal.ctlCliente1.Limpiar
    flgContinua = False
    Unload Me
End Select
End Sub

Private Sub Form_Load()
On Error GoTo handle
    If gstrIndRAv4 = "1" Then
        txtUbicacion.top = 240
    End If
    Me.KeyPreview = True
    flgContinua = False
    maxZoom = 20 '1048576
    flgZoom = False
    flgNoInfo = True
    Frame1.Caption = "Cliente - " & objVenta.CodigoClienteDLV
    ctlTextBox2.Text = objVenta.NombreClienteDLV
    ctlTextBox1.Text = objVenta.DireccionClienteDLV
    ctlTextBox3.Text = objVenta.DesReferenciaCli
    txtUrbanizacion.Text = objVenta.DesUrbanizacionDLV
    txtDistrito.Text = objVenta.DesDistritoDLV
    'objVenta.dc_departamentBK = ""
    
'    txtObsCliente.Text = objVenta.ObservacionClienteDLV
'    txtObsLocal.Text = objVenta.ObsNotaLocal

    If objVenta.Out_Tipo = "1" Then
        TxtRazonSocial.Text = objVenta.Out_NombreCliente
        TxtRuc.Text = objVenta.Out_NumeroId
    Else
        TxtRazonSocial.Text = ""
        TxtRuc.Text = ""
    End If
    
    '******MAPA******'
    flgMapa = False
    'txtSearch.Text = objVenta.DireccionClienteDLV + " " + txtDistrito.Text
    Dim strAdic As String
    Dim valExisteComa As Integer
    Dim strSuf As String
    Dim valI As Integer
    Dim indExistePunto As Integer
    Dim indExisteSufijo As Integer
    Dim strSufijo As String
    Dim strDireccion As String
    valExisteComa = InStr(1, objVenta.DireccionClienteDLV, ",", vbTextCompare)
    strAdic = IIf(valExisteComa > 0, "", ", " + txtDistrito.Text)
    txtSearch.Text = objVenta.DireccionClienteDLV + strAdic
    indExistePunto = InStr(1, objVenta.bk_SufijoDir, ".")
    If indExistePunto > 0 Then indExisteSufijo = InStr(1, UCase(txtSearch.Text), UCase(objVenta.bk_SufijoDir))
    If indExisteSufijo > 0 Then
        strDireccion = txtSearch.Text
    Else
        If Len(objVenta.bk_SufijoDir) > 0 Then strSufijo = Mid(objVenta.bk_SufijoDir, 1, indExistePunto - 1)
        indExisteSufijo = InStr(1, UCase(txtSearch.Text), UCase(strSufijo) & " ")
        If indExisteSufijo > 0 Then
            strDireccion = txtSearch.Text
        Else
            strDireccion = objVenta.bk_SufijoDir + " " + txtSearch.Text
        End If
    End If
    txtSearch.Text = strDireccion
    'valI = InStr(1, txtSearch.Text, strSufijoDir)
    'If valI = 0 Then txtSearch.Text = objVenta.bk_SufijoDir + " " + txtSearch.Text
    'ucGMap1.SetCenterToTextLocation txtSearch.Text
    If gstrIndRAv4 = "1" Then
        objVenta.DireccionCliente = txtUbicacion.Text  'txtSearch.Text
    Else
        objVenta.DireccionCliente = txtSearch.Text
    End If
    objVenta.DireccionClienteDLV = objVenta.DireccionClienteDLV
    objVenta.Out_Direccion = objVenta.DireccionClienteDLV
    
    Dim i&
    For i = 0 To 100 Step 10: cmbZoom.AddItem i & "%": Next
    cmbZoom.ListIndex = 9
  
    For i = 0 To 3: cmbMapType.AddItem ucGMap1.GetMapType(i): Next
    cmbMapType.ListIndex = 0

    Dim MarkerChar As String
    MarkerChar = Chr$(65)
    ucGMap1.AddMarker objVenta.Latitud & "," & objVenta.Longitud, vbGreen, MarkerChar
    'I.ECASTILLO 17.12.2020 | validar que lat tenga 1 solo guion, sino es erroneo
    Dim countG As Long
    Dim nVecesG As Long
    For countG = 1 To Len(objVenta.Latitud)
        If Mid(objVenta.Latitud, countG, 1) = "-" Then nVecesG = nVecesG + 1
    Next countG
    'F.ECASTILLO 17.12.2020
    Dim flgBuscaInMap
    flgBuscaInMap = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGGMSTR") '1 => ACTIVO, 0 => INACTIVO
    If flgBuscaInMap = "1" Then
        ucGMap1.countCalls "Load.txtSearch_KeyDown"
        If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido.Load txtSearch.enter"
        txtSearch_KeyDown 13, 0
    Else
        If objVenta.Latitud <> "" And objVenta.Longitud <> "" And nVecesG <= 1 Then
            'ucGMap1.DefaultCoord True, objVenta.Latitud & "," & objVenta.Longitud
            ucGMap1.GPoint = objVenta.Latitud & "," & objVenta.Longitud 'Setea coords existentes
            ucGMap1.countCalls "Load.Refresh"
            If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido.Load ucGMap1.Refresh; setea coords previas"
            ucGMap1.Refresh
        Else
            ucGMap1.countCalls "Load.txtSearch_KeyDown"
            If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido.Load txtSearch.enter"
            txtSearch_KeyDown 13, 0
        End If
    End If
    If objVenta.Latitud = "" Or objVenta.Longitud = "" Then
        cmdTomaPedido.Enabled = False
    Else
        cmdTomaPedido.Enabled = True
    End If
    flgZoom = True
    flgNoInfo = False
'    ucGMap1.Refresh
    '***************'
    strDireccion_old = txtSearch.Text
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtSearch_Change()
    'al cambiar el valor, copiar el valor sin distrito en txtUbicacion
    'txtUbicacion debe almacenarse en una variable para guardar en bd,
    'y enviar a digital en lugar de txtSearch
    Dim strDistrito As String
    Dim strDireccion2 As String
    Dim indExisteDistrito As Integer
    Dim indExisteComa As Integer
    If objVenta.dc_district <> txtDistrito.Text And Len(objVenta.dc_district) > 0 Then
        strDistrito = objVenta.dc_district
    Else
        strDistrito = txtDistrito.Text
    End If
    strDistrito = UCase(strDistrito)
'    indExisteComa = InStr(1, UCase(txtSearch.Text), ", " & strDistrito)
'    If indExisteComa > 0 Then
'        strDireccion2 = Mid(UCase(txtSearch.Text), 1, indExisteComa - 1)
'    Else
'        strDireccion2 = UCase(txtSearch.Text)
'    End If
    strDireccion2 = UCase(txtSearch.Text)
'    indExisteDistrito = InStr(1, strDireccion2, strDistrito)
'    If indExisteDistrito > 0 Then strDireccion2 = Mid(strDireccion2, 1, indExisteDistrito - 1)
    
    Debug.Print strDireccion2
    txtUbicacion.Text = strDireccion2
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo HANDLERERROR
    If KeyCode <> 13 Then Exit Sub
    If Me.Visible Then cmbZoom.SetFocus
    'DoEvents
    Dim oldSearch As String
    txtSearch.Text = Replace(txtSearch.Text, vbCrLf, "")
    oldSearch = txtSearch.Text
    If Len(txtSearch.Text) = 0 Then Exit Sub
    Dim coords() As String
    coords = Split(txtSearch.Text, ",")
    If UBound(coords) + 1 >= 2 Then
        If IsNumeric(coords(0)) And IsNumeric(coords(1)) Then
            flgNoInfo = False
            If Trim$(txtSearch.Text) <> "0,0" Then
                ucGMap1.flgBuscaCoords = False
                ucGMap1.GPoint = txtSearch.Text
                ucGMap1.countCalls "txtSearch_KeyDown.coords"
                If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido txtSearch.enter ucgMap1_DblClick; setea coords"
                ucGMap1_DblClick ucGMap1.GPoint
                strDireccion_old = txtSearch.Text
                objVenta.dc_street = txtSearch.Text
            End If
        Else
            GoTo buscaUbicacion
        End If
    Else
buscaUbicacion:
        flgNoInfo = True
        ucGMap1.flgBuscaCoords = True
        ucGMap1.countCalls "txtSearch_KeyDown.dir"
        If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido txtSearch.enter ucgMap1.SetCenterToTextLocation"
        ucGMap1.SetCenterToTextLocation txtSearch.Text
        If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido txtSearch.enter ucGMap1_DblClick"
        ucGMap1_DblClick ucGMap1.GPoint
        txtSearch.Text = oldSearch
        objVenta.DireccionCliente = txtSearch.Text
        objVenta.DireccionClienteDLV = txtSearch.Text
        objVenta.Out_Direccion = txtSearch.Text
        flgNoInfo = False
    End If
    'ucGMap1.Refresh
    strDireccion_old = txtSearch.Text
    objVenta.dc_street = txtSearch.Text
    Exit Sub
HANDLERERROR:
    ucGMap1.GPoint = ucGMap1.respaldo
    'Err.Raise Err.Number, "mapa.show", Err.Description
    Err.Clear
End Sub
Private Sub cmbZoom_Click()
    Dim zoom As Long
    zoom = Replace(cmbZoom.Text, "%", "") 'Mid(cmbZoom.Text, 1, 2)
    ucGMap1.GZoom = CInt((maxZoom * zoom) / 100)
  'ucGMap1.SetCenterToTextLocation txtSearch.Text
    If flgZoom Then ucGMap1.countCalls "cmbZoom_Click": txtSearch_KeyDown 13, 0  'ucGMap1.Refresh
End Sub

Private Sub cmbMapType_Click()
  ucGMap1.MapType = cmbMapType.ListIndex
End Sub

Private Sub txtUbicacion_LostFocus()
    Debug.Print "AB"
    If InStr(1, txtSearch.Text, txtUbicacion.Text) > 0 Then
    Else
        Debug.Print txtSearch.Text
        txtSearch.Text = txtUbicacion.Text
        Debug.Print txtSearch.Text
    End If
End Sub

Private Sub ucGMap1_MouseUp(ByVal GMouseCoordLatLng As String)
'    cmbZoom.SetFocus
    If ucGMap1.xyIFDist = 1 Then
        ucGMap1.countCalls "ucGMap1_MouseUp"
        If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido ucgMap1_MouseUp ucGMap1.Refresh"
        ucGMap1.Refresh
    End If
End Sub
Private Sub ucGMap1_DblClick(ByVal GMouseCoordLatLng As String)
    Dim MarkerChar As String
    Dim PuntosCord() As String
    Dim Direccion As String
    If flgMapa = False Then Exit Sub
    PuntosCord = Split(GMouseCoordLatLng, ",")
    If UBound(PuntosCord) + 1 >= 2 Then
        If IsNumeric(PuntosCord(0)) And IsNumeric(PuntosCord(1)) Then
            ucGMap1.GPoint = GMouseCoordLatLng
            GoTo continuar
        Else
            Exit Sub
        End If
    End If
continuar:

    
    While ucGMap1.Markers.Count > 0
        ucGMap1.Markers.Remove (1)
    Wend
    MarkerChar = Chr$(65)
    ucGMap1.AddMarker GMouseCoordLatLng, vbGreen, MarkerChar
      
    'If ucGMap1.Markers.Count > 1 Then
    '    ucGMap1.Markers.Remove (1)
    'End If
    If flgNoInfo = False Then
        'txtSearch.Text = objVenta.dc_street 'ucGMap1.GetInfoFromLatLng(GMouseCoordLatLng)
        If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido ucgMap1_DblClick ucGMap1.GetInfoFromLatLng"
        txtSearch.Text = ucGMap1.GetInfoFromLatLng(GMouseCoordLatLng)
        objVenta.DireccionCliente = txtSearch.Text
        objVenta.DireccionClienteDLV = txtSearch.Text
        objVenta.Out_Direccion = txtSearch.Text
    End If
    ucGMap1.countCalls "ucGMap1_DblClick"
    If gstrFlagLogBD3 = "1" Then objWS.grabaLogDelivery "frm_VTA_PreviaTomaPedido ucgMap1_DblClick ucGMap1.Refresh"
    ucGMap1.Refresh
    strDireccion_old = txtSearch.Text
    objVenta.dc_street = txtSearch.Text
End Sub
