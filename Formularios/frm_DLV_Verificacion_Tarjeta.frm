VERSION 5.00
Begin VB.Form frm_DLV_Verificacion_Tarjeta 
   Caption         =   "Verificación de Tarjeta"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   Icon            =   "frm_DLV_Verificacion_Tarjeta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   11565
   Begin vbp_Ventas.ctlTextBox txtObservaciones 
      Height          =   495
      Left            =   60
      TabIndex        =   41
      Top             =   4020
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin VB.CommandButton CmdDetalle 
      Caption         =   "&Detalle"
      Height          =   375
      Left            =   60
      TabIndex        =   40
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CommandButton CmdCliente 
      Caption         =   "&Cliente"
      Height          =   375
      Left            =   7800
      TabIndex        =   34
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   1380
      TabIndex        =   33
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10395
      TabIndex        =   2
      Top             =   4620
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   4620
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Frame Frame4 
         Caption         =   "Numero  Autorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   8040
         TabIndex        =   31
         Top             =   1080
         Width           =   2895
         Begin vbp_Ventas.ctlTextBox TxtNumAutor 
            Height          =   615
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1085
            Alignment       =   2
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   1845
         Width           =   525
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cuotas"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1515
         Width           =   795
      End
      Begin VB.Label LblImporte 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   28
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label LblNroCuota 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label LblDesCuota 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2160
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Venc"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1155
         Width           =   870
      End
      Begin VB.Label LblFchVenc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   24
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nº Tarjeta"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   795
         Width           =   720
      End
      Begin VB.Label LblNumTarj 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   22
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Tarjeta"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   435
         Width           =   855
      End
      Begin VB.Label LblTipoTarj 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   20
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "F1   Tarjetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   0
      TabIndex        =   18
      Top             =   2460
      Width           =   11535
      Begin vbp_Ventas.ctlGrillaArray grdTarjetas 
         Height          =   1155
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2037
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   11535
      Begin vbp_Ventas.ctlDataCombo ctlCboMotivo 
         Height          =   315
         Left            =   1440
         TabIndex        =   36
         Top             =   1905
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label LblLocalSapAsig 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   42
         Top             =   885
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
         Height          =   195
         Left            =   3960
         TabIndex        =   39
         Top             =   1613
         Width           =   630
      End
      Begin VB.Label LblFono 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   4920
         TabIndex        =   38
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Motivo Rechazo"
         Height          =   435
         Left            =   240
         TabIndex        =   37
         Top             =   1845
         Width           =   1050
      End
      Begin VB.Label LblCodCliente 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   6720
         TabIndex        =   35
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1613
         Width           =   480
      End
      Begin VB.Label LblDistrito 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   16
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Direcc Entrega"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1268
         Width           =   1065
      End
      Begin VB.Label LblDireccEntrega 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   14
         Top             =   1215
         Width           =   5055
      End
      Begin VB.Label LblDoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4920
         TabIndex        =   13
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   938
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido"
         Height          =   195
         Left            =   6720
         TabIndex        =   11
         Top             =   293
         Width           =   720
      End
      Begin VB.Label LblPedido 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   7680
         TabIndex        =   10
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label LblTeleoperadora 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   555
         Width           =   5055
      End
      Begin VB.Label LblLocalAsig 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   885
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   293
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Teleoperadora"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   608
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Local asignado"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   938
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frm_DLV_Verificacion_Tarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProforma As New clsProforma
Dim objTarjeta As New clsTarjeta
Public objFormaPago As clsFormaPago
Public pPedido As String
Public pFono As String
Dim strCodDireccionCli As String

Private Sub Form_Activate()
    '**************************************************************************************************************'
    '***** Vuelve a cargar el arreglo en el active por lo que se abre la misma ventana mas de una vez       *******'
    '***** y cunado exite mas de una ventana abierta 100pre se queda en el obeto el ultimo valor levantado  *******'
    '***** no siendo este especificamente el que se quiera verificar                                        *******'
    '**************************************************************************************************************'
    '*****                      CAMBIO REALIZADO EL 15/10/2007 Por Cristhian Rueda                          *******'
    '**************************************************************************************************************'
    If grdTarjetas.Columns(3).Text <> "" Then Exit Sub
    CARGA_Valores LblPedido.Caption, _
                  LblCodCliente.Caption, _
                  LblCliente.Caption, _
                  LblDireccEntrega.Caption, _
                  LblDistrito.Caption, _
                  LblLocalAsig.Caption, _
                  LblDoc.Caption, _
                  LblFono.Caption, _
                  "", _
                  objFormaPago, _
                  LblLocalSapAsig.Caption
    '**************************************************************************************************************'
End Sub

Private Sub Form_Load()
    SeteaGrilla
End Sub

Private Sub cmdAceptar_Click()
    '-----------------------------------------------------------------------------------------'
    '-- Actualiza el Nº autorización del las tarjetas que tenga dicha proforma
    '-- Hecho 04/07/2007 Por CRUEDA
    '-----------------------------------------------------------------------------------------'
       
    Dim i As Integer
    Dim strCadNumAut As String
    Dim strCadCodHijoFP As String
    Dim strCadFlgVerifPos As String
    
    strCadNumAut = "": strCadCodHijoFP = "": strCadFlgVerifPos = ""
    
      If objFormaPago.TarjetaVerif.UpperBound(1) = -1 Then Exit Sub
      For i = 0 To objFormaPago.TarjetaVerif.UpperBound(1)
       If objFormaPago.TarjetaVerif(i, 3) = "" Then
         MsgBox "A alguna tarjeta le falta el Nº Autorización", vbCritical, Caption: Exit Sub
       Else
         strCadNumAut = strCadNumAut & objFormaPago.TarjetaVerif(i, 3) & "|"
         strCadCodHijoFP = strCadCodHijoFP & objFormaPago.TarjetaVerif(i, 0) & "|"
         strCadFlgVerifPos = strCadFlgVerifPos & Mid(objFormaPago.TarjetaVerif(i, 10), 1, 1) & "|"
         
'         '---------------------------------'
'         Dim xi As Integer
'         For xi = 0 To objFormaPago.TarjetaVerif.UpperBound(1)
'            strCadNumAut = strCadNumAut & grdTarjetas.Columns(3).Text & "|"
'            If Len(strCadNumAut) = "1" Then
'                MsgBox "A alguna tarjeta le falta el Nº Autorización", vbCritical, Caption: Exit Sub
'            End If
'            strCadCodHijoFP = strCadCodHijoFP & grdTarjetas.Columns(0).Value & "|"
'         Next xi
         '---------------------------------'
         
       End If
      Next i
    '-----------------------------------------------------------------------------------------'
    '-----------------------------------------------------------------------------------------'
    Dim strMensaje As String
    strMensaje = objFormaPago.VerfirificaTarjeta(objUsuario.CodigoEmpresa, _
                                                  objUsuario.CodigoLocal, _
                                                  LblPedido.Caption, _
                                                  objUsuario.Codigo, _
                                                  strCadNumAut, _
                                                  strCadCodHijoFP, _
                                                  ctlCboMotivo.BoundText, _
                                                  strCadFlgVerifPos _
                                                  )
    If strMensaje = "" Then
      MsgBox "Se grabo satisfactoriamente ", vbExclamation, App.ProductName
    Else
      MsgBox strMensaje, vbCritical, App.ProductName
    End If
    
    frm_DLV_Verificacion.psubActualiza
    Unload Me
End Sub

Private Sub cmdCliente_Click()
     Dim frm_vCliente As New frm_DLV_Verificacion_Cliente
     frm_vCliente.Show
     frm_DLV_Verificacion_Cliente.pBlnCliente = False
     frm_vCliente.pBlnCliente = False
     frm_vCliente.ctlCliente1.Cargar
     frm_vCliente.SSTab1.Tab = 0
     frm_vCliente.Carga_Datos_Pedido
     frm_vCliente.ctlCliente1.Verificar
     frm_vCliente.ctlCliente1.CodDireccionCli = strCodDireccionCli
     frm_vCliente.ctlCliente1.ConsultaCliente Trim(LblCodCliente.Caption)
     If frm_DLV_Verificacion_Cliente.pBlnCliente = False Then
        frm_vCliente.LblPedido.Caption = Trim(LblPedido.Caption)
        frm_vCliente.lblTelefono.Caption = Trim(LblFono.Caption)
     End If
End Sub

Private Sub cmdDetalle_Click()
    On Error GoTo CtrlErr
    
    If grdTarjetas.ApproxCount = 0 Then Exit Sub
    frm_VTA_DetallePedido.NumeroPedido = Trim(LblPedido.Caption)
    frm_VTA_DetallePedido.CodigoLocal = Trim(LblLocalAsig.Caption)
    frm_VTA_DetallePedido.ReCargaDetPedido
    frm_VTA_DetallePedido.Show vbModal
    Exit Sub
    
CtrlErr:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
    If KeyCode = vbKeyF1 Then grdTarjetas.SetFocus
End Sub

Private Sub grdTarjetas_DblClick()
    If grdTarjetas.ApproxCount <= 0 Then Exit Sub
    Set frm_DLV_Verificacion_Trajeta_Chk.frmVTarjeta = Me
    frm_DLV_Verificacion_Trajeta_Chk.Show vbModal
End Sub

Private Sub grdTarjetas_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
                grdTarjetas_DblClick
    End Select
End Sub

Private Sub grdTarjetas_RegistroSeleccionado(ByVal DatoColumna0 As String)
     If grdTarjetas.ApproxCount <= 0 Then Exit Sub
     LblTipoTarj.Caption = grdTarjetas.Columns(0).Value
     LblNumTarj.Caption = Format(grdTarjetas.Columns(2).Value, "####-####-####-####")
     LblFchVenc.Caption = grdTarjetas.Columns(4).Value
     LblNroCuota.Caption = grdTarjetas.Columns(5).Value
     LblDesCuota.Caption = grdTarjetas.Columns(6).Value
     LblImporte.Caption = grdTarjetas.Columns(7).Value
     TxtNumAutor.Text = grdTarjetas.Columns(3).Value
End Sub




Private Sub TxtNumAutor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
    Dim CtrlGraba As String
    CtrlGraba = objFormaPago.GrabaNumAut(objUsuario.CodigoEmpresa, _
                                         objUsuario.CodigoLocal, _
                                         LblPedido.Caption, _
                                         LblTipoTarj.Caption, _
                                         LblNumTarj.Caption, _
                                         TxtNumAutor.Text)

        If CtrlGraba = "" Then
            MsgBox "Se Grabo el Numero de Autorización", vbInformation, Caption
            llenagrilla
        Else
            MsgBox CtrlGraba, vbCritical, Caption
        End If
    End If
End Sub

Private Sub llenagrilla()
 objFormaPago.TarjetaVerif.ReDim 0, -1, 0, 11
 grdTarjetas.Array1 = objFormaPago.TarjetaVerif
 
 objFormaPago.LoadDetTarVerif LblPedido.Caption, _
                              LblCodCliente.Caption
                                                       
 grdTarjetas.Array1 = objFormaPago.TarjetaVerif
    
End Sub

Public Sub CARGA(ByVal vstrNumProforma As String, _
                 ByVal vstrCodigoCli As String, _
                 ByVal vstrNombres As String, _
                 ByVal vstrDirecc As String, _
                 ByVal vstrDist As String, _
                 ByVal vstrlocalRef As String, _
                 ByVal vstrTipoDoc As String, _
                 ByVal vstrTelefono As String, _
                 ByVal vstrMotivoRech As String, _
                 ByVal vstrObservaciones As String, _
                 ByVal vstrCodDireccionCli As String, _
                 ByRef frmOwn As Form, _
                 ByRef robjFP As clsFormaPago, _
                 Optional ByVal vstrlocalSapRef As String _
                 )

    Set objFormaPago = robjFP
    LblTeleoperadora.Caption = objUsuario.Nombre
    LblPedido.Caption = vstrNumProforma
    LblCodCliente.Caption = vstrCodigoCli
    LblCliente.Caption = vstrNombres
    LblDireccEntrega.Caption = vstrDirecc
    LblDistrito.Caption = vstrDist
    LblLocalAsig.Caption = vstrlocalRef
    LblLocalSapAsig.Caption = vstrlocalSapRef
    LblDoc.Caption = vstrTipoDoc
    LblFono.Caption = vstrTelefono
    txtObservaciones.Text = vstrObservaciones
    strCodDireccionCli = vstrCodDireccionCli

    llenagrilla
    
    Set ctlCboMotivo.RowSource = objTarjeta.ListaMotivoRechazo
    ctlCboMotivo.ListField = "DES"
    ctlCboMotivo.BoundColumn = "COD"
    If vstrMotivoRech = "" Then
        ctlCboMotivo.BoundText = "000"
      Else
        ctlCboMotivo.BoundText = vstrMotivoRech
    End If
    Me.Show , frmOwn
End Sub

Public Sub CARGA_Valores(ByVal vstrNumProforma As String, _
                         ByVal vstrCodigoCli As String, _
                         ByVal vstrNombres As String, _
                         ByVal vstrDirecc As String, _
                         ByVal vstrDist As String, _
                         ByVal vstrlocalRef As String, _
                         ByVal vstrTipoDoc As String, _
                         ByVal vstrTelefono As String, _
                         ByVal vstrMotivoRech As String, _
                         ByRef robjFP As clsFormaPago, _
                         Optional ByVal vstrlocalSapRef As String _
                         )

    Set objFormaPago = robjFP
    LblTeleoperadora.Caption = objUsuario.Nombre
    LblPedido.Caption = vstrNumProforma
    LblCodCliente.Caption = vstrCodigoCli
    LblCliente.Caption = vstrNombres
    LblDireccEntrega.Caption = vstrDirecc
    LblDistrito.Caption = vstrDist
    LblLocalAsig.Caption = vstrlocalRef
    LblLocalSapAsig.Caption = vstrlocalSapRef
    LblDoc.Caption = vstrTipoDoc
    LblFono.Caption = vstrTelefono
    llenagrilla
    
    Set ctlCboMotivo.RowSource = objTarjeta.ListaMotivoRechazo
    ctlCboMotivo.ListField = "DES"
    ctlCboMotivo.BoundColumn = "COD"
    If vstrMotivoRech = "" Then
        ctlCboMotivo.BoundText = "000"
      Else
        ctlCboMotivo.BoundText = vstrMotivoRech
    End If
End Sub


Private Sub cmdAnular_Click()
Dim strAnula As String
Dim Bookmark As Variant
    If MsgBox("Se procederá a Anular la Proforma pagada con Tarjeta", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
        strAnula = objProforma.Anula(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Trim(LblPedido.Caption), objUsuario.Codigo)
        If strAnula = "" Then
            Unload Me
        Else
            MsgBox strAnula, vbCritical + vbOKOnly, App.ProductName
            
        End If
    End If
End Sub

Private Sub SeteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant
  
    arrCampos = Array("", "", "", _
                      "", "", _
                      "", "", _
                      "", "", _
                      "", "")

    arrCaption = Array("Codigo", "Tarjeta", "Nº Tarjeta", _
                       "Nº Autorización", "Fch Venc", _
                       "NºCuot", "Pago Dife", _
                       "Importe", "POS", _
                       "", "Verificación")
                       
    arrAncho = Array(700, 2100, 2220, _
                     1200, 1200, _
                     650, 950, _
                     800, 2000, _
                     600, 1400)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft)
                          
    arrFoco = Array(False, False, False, _
                    True, False, _
                    False, False, _
                    False, False, _
                    False, False)
                              
    grdTarjetas.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    grdTarjetas.StylesFondo = True
    grdTarjetas.StylesSize = 9
    grdTarjetas.Columns(9).Visible = False
    'grdTarjetas.Columns(10).Visible = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub xNombre_Click()
End Sub
