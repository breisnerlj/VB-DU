VERSION 5.00
Begin VB.Form frm_DLV_Ruteo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de local y motorizado"
   ClientHeight    =   10200
   ClientLeft      =   2385
   ClientTop       =   510
   ClientWidth     =   12045
   Icon            =   "frm_DLV_Ruteo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAsignaSinValidar 
      Caption         =   "Asignar sin validar transferencia"
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
      Left            =   120
      TabIndex        =   52
      Top             =   9600
      Width           =   3375
   End
   Begin VB.CommandButton cmdPedido 
      Caption         =   "Detalle Pedido"
      Height          =   615
      Left            =   10680
      Picture         =   "frm_DLV_Ruteo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   9480
      Width           =   1200
   End
   Begin vbp_Ventas.ctlTextBox txtObsMotorizado 
      Height          =   975
      Left            =   5760
      TabIndex        =   50
      Top             =   7440
      Width           =   4400
      _ExtentX        =   7752
      _ExtentY        =   1720
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
   Begin VB.CommandButton cmdActObserv 
      Caption         =   "Act &Observ."
      Height          =   615
      Left            =   10800
      Picture         =   "frm_DLV_Ruteo.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTransferencias 
      Caption         =   "&Transferencia"
      Height          =   615
      Left            =   9240
      Picture         =   "frm_DLV_Ruteo.frx":0E1E
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   7035
      Picture         =   "frm_DLV_Ruteo.frx":13A8
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   5460
      Picture         =   "frm_DLV_Ruteo.frx":1932
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9480
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox12 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   7440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
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
   Begin VB.CheckBox Check2 
      Caption         =   "Asignar Motorizado"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   795
      Left            =   0
      TabIndex        =   21
      Top             =   6360
      Width           =   12000
      Begin vbp_Ventas.ctlDataCombo cboRuta2 
         Height          =   315
         Left            =   2895
         TabIndex        =   24
         Top             =   420
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboLocalAsig2 
         Height          =   315
         Left            =   5355
         TabIndex        =   25
         Top             =   420
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboMotorizado 
         Height          =   315
         Left            =   7995
         TabIndex        =   27
         Top             =   420
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboCiaMot 
         Height          =   315
         Left            =   200
         TabIndex        =   55
         Top             =   420
         Width           =   2600
         _ExtentX        =   4577
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin VB.Label Label18 
         Caption         =   "Cia:"
         Height          =   255
         Left            =   200
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Motorizado"
         Height          =   255
         Left            =   7995
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Local"
         Height          =   255
         Left            =   5355
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Zona:"
         Height          =   255
         Left            =   2895
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tranferencias"
      Height          =   1335
      Left            =   0
      TabIndex        =   19
      Top             =   4980
      Width           =   12000
      Begin vbp_Ventas.ctlGrillaArray grdTransferencia 
         Height          =   975
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   1720
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Asignación de local"
      Height          =   2175
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   12000
      Begin vbp_Ventas.ctlGrillaArray grdPedidoStock 
         Height          =   1335
         Left            =   120
         TabIndex        =   60
         Top             =   720
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2355
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.CommandButton cmdStockEnLinea 
         Caption         =   "&Stock en Linea"
         Height          =   615
         Left            =   10440
         Picture         =   "frm_DLV_Ruteo.frx":1EBC
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1460
         Width           =   1335
      End
      Begin vbp_Ventas.ctlDataCombo cboRuta1 
         Height          =   315
         Left            =   4080
         TabIndex        =   35
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.CommandButton cmdTranferencia 
         Caption         =   "&Transferencia"
         Height          =   615
         Left            =   10440
         Picture         =   "frm_DLV_Ruteo.frx":22B9
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   740
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Solo locales con stock"
         Height          =   375
         Left            =   9960
         TabIndex        =   17
         Top             =   330
         Width           =   2000
      End
      Begin vbp_Ventas.ctlDataCombo cboLocalAsig 
         Height          =   315
         Left            =   6840
         TabIndex        =   18
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboCia 
         Height          =   315
         Left            =   480
         TabIndex        =   53
         Top             =   360
         Width           =   2900
         _ExtentX        =   5106
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label28 
         Caption         =   "Cia:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Local"
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   390
         Width           =   450
      End
      Begin VB.Label Label12 
         Caption         =   "Zona:"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   390
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dirección de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   12000
      Begin vbp_Ventas.ctlTextBox txtUrbanizacion 
         Height          =   315
         Left            =   7560
         TabIndex        =   48
         Top             =   600
         Width           =   4300
         _ExtentX        =   7594
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtDireccionEntrega 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtReferencia 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   960
         Width           =   10659
         _ExtentX        =   18812
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtUbigeo 
         Height          =   315
         Left            =   1200
         TabIndex        =   38
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Bloqueado       =   -1  'True
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Urbanización"
         Height          =   195
         Left            =   6480
         TabIndex        =   47
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1020
         Width           =   780
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
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin vbp_Ventas.ctlTextBox txtNumPedidoIkf 
         Height          =   375
         Left            =   9600
         TabIndex        =   59
         Top             =   200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ColorDefault    =   -2147483643
         ColorDefault    =   -2147483643
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtNumPedido 
         Height          =   375
         Left            =   7200
         TabIndex        =   4
         Top             =   200
         Width           =   2300
         _ExtentX        =   4048
         _ExtentY        =   661
         ColorDefault    =   -2147483643
         ColorDefault    =   -2147483643
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtDocIdent 
         Height          =   315
         Left            =   4600
         TabIndex        =   5
         Top             =   600
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtTeleoperadora 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   2300
         _ExtentX        =   4048
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtNombreCliente 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   5455
         _ExtentX        =   9631
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Bloqueado       =   -1  'True
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
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   8400
         TabIndex        =   33
         Top             =   1200
         Width           =   975
      End
      Begin vbp_Ventas.ctlDataCombo cboLocal 
         Height          =   315
         Left            =   9600
         TabIndex        =   40
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdGrabarLocal 
         Caption         =   "Asignar Local"
         Height          =   375
         Left            =   9600
         TabIndex        =   41
         Top             =   570
         Visible         =   0   'False
         Width           =   1335
      End
      Begin vbp_Ventas.ctlDataCombo cboCiaAsig 
         Height          =   315
         Left            =   7200
         TabIndex        =   57
         Top             =   600
         Width           =   2300
         _ExtentX        =   4048
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loc. Asig."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6405
         TabIndex        =   39
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teleoperadora"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         Height          =   195
         Left            =   3700
         TabIndex        =   2
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   480
      End
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   8100
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   10695
      TabIndex        =   43
      Top             =   8520
      Visible         =   0   'False
      Width           =   10695
      Begin vbp_Ventas.ctlGrillaArray grdEntregaTerceros 
         Height          =   495
         Left            =   60
         TabIndex        =   45
         Top             =   315
         Width           =   11000
         _ExtentX        =   19394
         _ExtentY        =   873
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "PEDIDO ES CON ENTREGA A TERCEROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   60
         TabIndex        =   44
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observación a Motorizado"
      Height          =   195
      Left            =   5760
      TabIndex        =   49
      Top             =   7200
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observación de Ruteo :"
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   7860
      Width           =   1695
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Observación de Local :"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   7200
      Width           =   1650
   End
End
Attribute VB_Name = "frm_DLV_Ruteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strLocal As String
Public strNumProforma As String
Public strNumPedRef As String
Public strLocalPedido As String
Public strFlgTransf As String
Public bolEsLlamadoCab As Boolean
Public bolModoConsulta As Boolean
'Public ErrReservarStock As Boolean
Dim rsPedido  As oraDynaset
Dim objPedido As New clsProforma
Dim objZona As New clsZona
Dim strZonaLocalAsignado As String
Dim CodCliente As String
Dim CodDireccionCli As String
Dim intErr As Byte
Dim strCodEstado As String
Public strCia As String
Dim esLoad As Boolean
Dim objLocal As New clsLocal
Dim objProducto As New clsProducto
Dim objWS As New clsWebService
Dim arrInfo As New XArrayDB
Dim productosValidos As String

Private Sub cboCia_Change()
    Dim sCodEmp As String
 
    If esLoad Then Exit Sub ' es la carga inicial
 
    If grdTransferencia.ApproxCount > 0 Then
        MsgBox "Existen transferencias, no puede cambiar de Empresa", vbCritical + vbOKOnly, App.ProductName
        Exit Sub
    End If
 
    If cboCia.BoundText = "99" Then
        cmdStockEnLinea.Visible = True
    Else
        cmdStockEnLinea.Visible = False
    End If
 
    sCodEmp = cboCia.BoundText
    'MsgBox "cboCia Click, Sel :" + cboCia.Text + " Bound Text: " + sCodEmp
    'Set rsLocal = objlocal.Lista(cboCia.BoundText) ', IIf(objUsuario.flgDeliveryProv = 1, objUsuario.CodigoLocal, ""))
    cboRuta1.BoundText = ""
    cboLocalAsig.BoundText = ""
    'Set cboLocalAsig.RowSource = objZona.ListaLocal(objUsuario.CodigoEmpresa, cboRuta1.BoundText)
    Set cboLocalAsig.RowSource = objZona.ListaLocal(sCodEmp, cboRuta1.BoundText, "", sCodEmp)
    cboLocalAsig.BoundColumn = "COD_LOCAL"
    cboLocalAsig.ListField = "local_dex"
 
End Sub

Private Sub cboCiaMot_Change()
 Dim sCodEmp As String
 Dim objLocal As New clsLocal
 Dim rsLocal As oraDynaset
    
 sCodEmp = cboCiaMot.BoundText
 'MsgBox "Sel :" + cboCia.Text + " Bound Text: " + sCodEmp
 'Set rsLocal = objlocal.Lista(cboCia.BoundText) ', IIf(objUsuario.flgDeliveryProv = 1, objUsuario.CodigoLocal, ""))
 
 Set rsLocal = objLocal.Lista(sCodEmp)
 Set cboLocalAsig2.RowSource = rsLocal ' objLocal.Lista(objUsuario.CodigoEmpresa, "")
 cboLocalAsig2.ListField = "local_dex2"
 cboLocalAsig2.BoundColumn = "COD_LOCAL"
 cboLocalAsig2.BoundText = objUsuario.CodigoLocal
 Set rsLocal = Nothing
 Set objLocal = Nothing
End Sub

Private Sub cboLocalAsig_Change()
    buscaPedidoStock
End Sub

Private Sub cboLocalAsig_Click(Area As Integer)
    'buscaPedidoStock
End Sub

Private Sub cboLocalAsig2_Click(Area As Integer)
If cboLocalAsig2.BoundText = "" Then Exit Sub
    Dim objMotorizado As New clsMotorizado
    'MsgBox "cboLocalAsig2.BoundText :" + cboLocalAsig2.BoundText 'LOCAL PARA MOTORIZADO JCT
    Set cboMotorizado.RowSource = objMotorizado.ListaDisponible(cboLocalAsig2.BoundText)
    cboMotorizado.BoundColumn = "COD_MOTORIZADO"
    cboMotorizado.ListField = "NOMBRE"
    Set objMotorizado = Nothing

End Sub




Private Sub cboRuta1_Change()
    Dim s As String
    If esLoad Then Exit Sub
     
    'If Area = 0 Then Exit Sub
    
    cboLocalAsig.BoundText = ""
    s = cboCia.BoundText ' codigo cia:10(BTL),99(Mi Farma)
    'MsgBox "S, cboRuta1.BoundText: " + s + "," + cboRuta1.BoundText + "End"
    
    ' para el caso de mostrar, solo locales con stock llamar directo a la funcion desde DB
'    Debug.Print "|" & cboCia.BoundText & "|"
'    Debug.Print "|" & cboRuta1.BoundText & "|"
    Set cboLocalAsig.RowSource = objZona.ListaLocal(s, cboRuta1.BoundText, "", s)
    cboLocalAsig.BoundColumn = "COD_LOCAL"
    cboLocalAsig.ListField = "local_dex"
End Sub

Private Sub cboRuta1_Click(Area As Integer)
    ' Dim s As String
    'If esLoad Then Exit Sub
     
    ''If Area = 0 Then Exit Sub
    
    'cboLocalAsig.BoundText = ""
    's = cboCia.BoundText ' codigo cia:10(BTL),99(Mi Farma)
    'MsgBox "S, cboRuta1.BoundText: " + s + "," + cboRuta1.BoundText
    
   ' Set cboLocalAsig.RowSource = objZona.ListaLocal(s, cboRuta1.BoundText)
   ' cboLocalAsig.BoundColumn = "COD_LOCAL"
   ' cboLocalAsig.ListField = "local_dex"
    Dim sCia As String
    'sCia = cboCia.BoundText ' codigo cia
    'Set cboLocalAsig2.RowSource = objZona.ListaLocal(objUsuario.CodigoEmpresa, cboRuta2.BoundText)
    'Set cboLocalAsig2.RowSource = objZona.ListaLocal(sCia, cboRuta2.BoundText)
    'cboLocalAsig2.BoundColumn = "COD_LOCAL"
    'cboLocalAsig2.ListField = "local_dex"
End Sub

Private Sub cboRuta2_Click(Area As Integer)
    Dim sCia As String
    sCia = cboCiaMot.BoundText ' codigo cia
    'Set cboLocalAsig2.RowSource = objZona.ListaLocal(objUsuario.CodigoEmpresa, cboRuta2.BoundText)
    Set cboLocalAsig2.RowSource = objZona.ListaLocal(sCia, cboRuta2.BoundText, "", sCia)
    cboLocalAsig2.BoundColumn = "COD_LOCAL"
    cboLocalAsig2.ListField = "local_dex"
End Sub



Private Sub Check2_Click()
    If Check2.Value = "0" Then
        cboRuta2.Enabled = False
        cboLocalAsig2.Enabled = False
        cboMotorizado.Enabled = False
        cboCiaMot.Enabled = False
    Else
        cboRuta2.Enabled = True
        cboLocalAsig2.Enabled = True
        cboMotorizado.Enabled = True
        cboRuta2.BoundText = cboRuta1.BoundText  'strZonaLocalAsignado
        cboLocalAsig2.BoundText = cboLocalAsig.BoundText
        'cboCiaMot.Enabled = True
        cboCiaMot.BoundText = cboCia.BoundText
        cboLocalAsig2.BoundText = cboLocalAsig.BoundText
    End If
End Sub


Private Sub cmdAceptar_Click()

On Error GoTo Control
    'I.ECASTILLO 07.07.2021
    If gstrFlagValRut = "1" Then
        If strNumPedRef <> strNumProforma And Len(Trim(strNumPedRef)) > 0 Then
            MsgBox "El pedido fue ruteado automaticamente, no se puede rutear otra vez.", vbInformation, App.ProductName
            Exit Sub
        End If
    End If
    'F.ECASTILLO 07.07.2021
    
    If cboLocalAsig.BoundText = "" Then
        MsgBox "Tiene que seleccionar local", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    
    Dim strPing As String
    Dim strServer As String
    strServer = objPedido.DevuelveIPServer(cboLocalAsig.BoundText)
    If strServer <> "" Then
    strPing = fnPing(strServer)
        If strPing <> "" Then
            MsgBox "El local seleccionado no tiene linea actualmente, intentolo luego" & strPing, vbExclamation, App.ProductName
        End If
    End If
    If cboMotorizado.BoundText = "" Then
        If cboCia.BoundText = "99" Then
            MsgBox "No se puede rutear pedidos a MIFARMA sin asignar motorizado", vbOKOnly + vbExclamation, "Asignar Motorizado"
            Exit Sub
        End If
        If frm_DLV_Seguimiento.Option2.Value = True Then
            If MsgBox("Desea pasar el Pedido a Asignado sin escoger un Motorizado?", vbQuestion + vbYesNo, "Confirme") = vbNo Then Exit Sub
        End If
    End If
    
    
    If chkAsignaSinValidar.Value = 1 Then
        If MsgBox("Desea asignar el pedido sin validar Transferecia ?", vbQuestion + vbYesNo, "Confirme") = vbNo Then Exit Sub
    End If
    
    
    
    'objUsuario.CodigoMotorizado = cboMotorizado.BoundText
    
'    ReservaStock
'    If ErrReservarStock = False Then
    Grabar
    
    If blnEnviaMensajeDelivery = True Then
    On Error Resume Next
    sub_EnviarMensaje cboLocalAsig.BoundText
    On Error GoTo Control
    End If
    
    If bolEsLlamadoCab Then frm_DLV_Seguimiento.ListaPedido
    If intErr = 0 Then Unload Me
    
'    End If
    
    'Grabar
    
    'agregado por pherrera (no funca todavia)
    'para eviar el aviso a la pantalla del local
    


   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbExclamation, App.ProductName
    
    grdPedidoStock.MoveFirst
    grdTransferencia.MoveFirst
End Sub

Private Sub cmdActObserv_Click()
    ActObsRuta objUsuario.CodigoEmpresa, _
            strLocal, _
            strNumProforma, _
            ctlTextBox12.Text, _
            ctlTextBox1.Text, _
            txtObsMotorizado.Text
            
            
            
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdGrabarLocal_Click()
Dim strMensaje As String
strMensaje = objPedido.AsignaLocal(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, txtNumPedido.Text, cboLocal.BoundText)
If strMensaje = "" Then
    MsgBox "Se asisgnó el local", vbExclamation, App.ProductName
Else
    MsgBox strMensaje, vbExclamation, App.ProductName
End If
End Sub

Private Sub cmdPedido_Click()
On Error GoTo CtrlErr

    If Not cmdPedido.Enabled Then Exit Sub

    frm_VTA_DetallePedido.NumeroPedido = rsPedido("NUM_PROFORMA").Value
    frm_VTA_DetallePedido.CodigoLocal = rsPedido("COD_LOCAL_REF").Value
    frm_VTA_DetallePedido.ReCargaDetPedido
    frm_VTA_DetallePedido.Show vbModal
Exit Sub
CtrlErr:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub cmdStockEnLinea_Click()
    Dim strFic As String
    Dim strParam As String
    
    If cboLocalAsig.BoundText = vbNullString Then
        MsgBox "Seleccione un local de despacho", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    ' Esto es para el caso de que el programa a ejecutar no exista
    On Local Error Resume Next
        
    strFic = "explorer.exe"
    strParam = """http://10.85.8.200/delivery/p_stock_pedido.php?scia=" & objUsuario.CodigoEmpresa & _
               "&sloc=" & strLocal & _
               "&spro=" & strNumProforma & _
               "&sref=" & cboLocalAsig.BoundText & """ "
    Shell strFic & " " & strParam, vbMaximizedFocus

    ' Si se produce un error, lo comprobamos aquí
    If Err Then
        MsgBox "Se ha producido el siguiente error:" & vbCrLf & _
               Err.Number & ", " & Err.Description & vbCrLf & _
               "al intentar ejecutar:" & vbCrLf & _
               strFic & " " & strParam
    'Else
        'MsgBox "Se está ejecutando: " & strFic & " " & strParam
    End If
    
    ' Nos aseguramos que el valor del error sea cero
    Err = 0
End Sub

Private Sub cmdTranferencia_Click()
    On Error GoTo Control

    If cboLocalAsig.BoundText = "" Then
        MsgBox "Tiene que selccionar local", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    If grdPedidoStock.ApproxCount = 0 Then Exit Sub
    
    frm_DLV_Stock_Total.strTranferencias = True
    frm_DLV_Stock_Total.strDescripcionProducto = grdPedidoStock.Columns(1).Value
    'PHERRERA 28/09/07 cambio cboLocal por cboLocalAsig
    frm_DLV_Stock_Total.G_LOCAL_ASIGNADO = cboLocalAsig.BoundText
    frm_DLV_Stock_Total.G_LOCAL_SAP_ASIGNADO = Mid(cboLocalAsig.Text, 1, 3)
    frm_DLV_Stock_Total.strCodigoProducto = grdPedidoStock.Columns(0).Value
    'arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "Pedido", "STOCK")
    frm_DLV_Stock_Total.strStockLocal = grdPedidoStock.Columns("Ctd.Dispon").Value
    
    frm_DLV_Stock_Total.strCantPedida = grdPedidoStock.Columns("Ctd.Pedido").Value
    frm_DLV_Stock_Total.intCantFraccionamiento = grdPedidoStock.Columns("CTD_FRACCIONAMIENTO").Value
    'frm_DLV_Stock_Total.intCantProd = grdPedidoStock.Columns("CTD_PRODUCTO").Value
    'frm_DLV_Stock_Total.intCantProdFrac = grdPedidoStock.Columns("CTD_PRODUCTOFRAC").Value
    'frm_DLV_Stock_Total.intCantProd1 = grdPedidoStock.Columns("CTD_PRODUCTO1").Value
    'frm_DLV_Stock_Total.intCantProdFrac1 = grdPedidoStock.Columns("CTD_PRODUCTOFRAC1").Value
    'frm_DLV_Stock_Total.lblFalta.Caption = grdPedidoStock.Columns("DIF").Value
    frm_DLV_Stock_Total.DIF = grdPedidoStock.Columns("DIF").Value
    frm_DLV_Stock_Total.strCodZona = strZonaLocalAsignado
    ''''''''frm_DLV_Stock_Total.CantidadFaltante
    If val(grdPedidoStock.Columns("FLG_FRACCIONAMIENTO").Value) = 0 Then
        'frm_DLV_Stock_Total.FlgFraccionado = False
        frm_DLV_Stock_Total.chkFraccionamiento.Enabled = False
    Else
        'frm_DLV_Stock_Total.FlgFraccionado = True
        frm_DLV_Stock_Total.chkFraccionamiento.Enabled = True
    End If
    
    'B JCT, 04ABR12 Asignar CIA,
    frm_DLV_Stock_Total.cboCia.BoundText = cboCia.BoundText 'cboCiaAsig.BoundText
    frm_DLV_Stock_Total.cboCia.Enabled = False
    ' E JCT
    
    frm_DLV_Stock_Total.Show vbModal
    grdTransferencia.MoveFirst
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdTransferencias_Click()
On Error GoTo handle
    Dim objProforma As New clsProforma
    Dim x As String
    Dim strCadCodProducto As String
    Dim strCadCtdProducto  As String
    Dim strCadCtdProductoFrac  As String
    Dim strLocalOrigen  As String
    Dim strLocalDestino As String
    Dim i As Integer
    i = 0
    
    While i <= objVenta.Distribucion.UpperBound(1)
        strCadCodProducto = strCadCodProducto & objVenta.Distribucion(i, 1) & "|"
        If objVenta.Distribucion(i, 7) = 0 Then
            strCadCtdProducto = strCadCtdProducto & objVenta.Distribucion(i, 4) & "|"
            strCadCtdProductoFrac = strCadCtdProductoFrac & "0" & "|"
        Else
            strCadCtdProducto = strCadCtdProducto & "0" & "|"
            strCadCtdProductoFrac = strCadCtdProductoFrac & objVenta.Distribucion(i, 4) & "|"
        End If
        strLocalOrigen = strLocalOrigen & objVenta.Distribucion(i, 8) & "|"
        strLocalDestino = strLocalDestino & objVenta.Distribucion(i, 9) & "|"
        i = i + 1
    Wend
    If Not objVenta.Distribucion.UpperBound(1) = -1 Then
        x = objProforma.GrabaTransferencia(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strNumProforma, strCadCodProducto, strCadCtdProducto, strCadCtdProductoFrac, strLocalOrigen, strLocalDestino, objUsuario.Codigo)
    End If
    
    If Not x = "" Then GoTo handle
        Dim NumeroTransferencia As String
        x = objProforma.Asigna(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strNumProforma, objUsuario.Codigo, cboMotorizado.BoundText, objUsuario.NombrePC, ctlTextBox12.Text, strLocalDestino, "", "", "", "", "")
        If Not x = "" Then GoTo handle
    Set objProforma = Nothing
    MsgBox "Se grabo satisfactoriamente la Proforma :" & strNumProforma
    Exit Sub
handle:
    MsgBox Err.Description & x
    Set objProforma = Nothing
End Sub

Private Sub ctlTextBox10_KeyPress(KeyAscii As Integer)

End Sub

'Private Sub Command2_Click()
''On Error GoTo handle
'
''ReservaStock
''If ErrReservarStock = False Then
''    Grabar
''End If
'PreReservaStock
'
''handle:
''    MsgBox Err.Description, vbCritical, App.ProductName
'End Sub
'I.ECASTILLO 27.07.2020
Private Function PreReservaStock() As String
On Error GoTo Err
    Dim jsonMaestro, jsonMaestroDetalle As String
    Dim dicMaestro As New Dictionary
    Dim obj As New Dictionary
    Dim arrLocales As New XArrayDB
    Dim arrProductos As New XArrayDB
    Dim arrProductosLocales As New XArrayDB
    Dim codLocalBtl, prueba, commaM, commaMD As String
    Dim x, i, j, ultimo As Integer
    Dim existeLocal As Boolean
    Dim strA, strB, strC As Variant
    Dim rsCia As oraDynaset
    Dim sCia, sCodLoc, localSAP, estConfig As String
    Dim v_local, v_localSap, v_numProforma, v_cia As String
    Dim estConfigLocales As String
    Dim flgPosu, codProductoPosuInk As String
    Dim countTransf As Integer
'    productosValidos = ""
'    prueba = validaStock
'    If Not prueba = "" Then MsgBox prueba, vbOKOnly + vbCritical: grdPedidoStock.MoveFirst: grdTransferencia.MoveFirst
    If Len(productosValidos) = 0 Then GoTo salir
    'OBTENER SOLO LOCALES
    'OBTENER PRODUCTOS X LOCALES
    codLocalBtl = cboLocalAsig.BoundText
    arrLocales.ReDim 0, -1, 0, 1
    'LOCAL-PRINCIPAL
    If Len(codLocalBtl) > 0 Then arrLocales.AppendRows: arrLocales(0, 0) = codLocalBtl: arrLocales(0, 1) = 1
    'LOCALES-TRANSFERENCIAS
    grdTransferencia.MoveFirst
    While Not grdTransferencia.EOF
        If arrLocales.Count(1) > 0 Then
            x = 0: existeLocal = False
            While x <= arrLocales.Count(1) - 1 And existeLocal = False
                If arrLocales(x, 0) = grdTransferencia.Columns("Origen").Value Then existeLocal = True
                x = x + 1
            Wend
            If existeLocal = False Then
                arrLocales.AppendRows
                ultimo = arrLocales.Count(1) - 1
                arrLocales(ultimo, 0) = grdTransferencia.Columns("Origen").Value
                arrLocales(ultimo, 1) = 2
                countTransf = countTransf + 1
            End If
        End If
        grdTransferencia.MoveNext
    Wend
    If gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "RESVTRANSF") = 0 Then
        If countTransf > 0 Then PreReservaStock = "B": Exit Function
    End If
    'PRODUCTOS X LOCALES
    x = 0: estConfig = 0
    Debug.Print arrLocales.Count(1)
    For x = 0 To arrLocales.Count(1) - 1
        'OBTIENE CIA
        sCodLoc = "": localSAP = "": jsonMaestroDetalle = "": flgPosu = "": codProductoPosuInk = ""
        sCodLoc = arrLocales(x, 0)
        Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, sCodLoc)
        If (rsCia.RecordCount > 0) Then
          sCia = CStr(rsCia(1))
        End If
        Set rsCia = Nothing
        flgPosu = objLocal.EvaluaLocalInkMig(sCodLoc)
        If flgPosu = "N" Then
            localSAP = objLocal.GetCodInka(sCodLoc)
        Else
            localSAP = objLocal.GetCodPosu(sCodLoc)
        End If
'        If arrLocales(X, 1) = 1 Then estConfig = objLocal.GetEstConfig(sCia, sCodLoc, "RESERVA_STOCK")
        'productosValidos => 'producto°local°cantidad°flg_frac°ctd_frac°desc|,...
        strA = "": Debug.Print productosValidos
        strA = Split(productosValidos, "|") 'producto°local°cantidad°flg_frac°ctd_frac°desc
        i = 0
        For i = 0 To UBound(strA) - 1
            strB = "": Debug.Print strA(i)
            strB = Split(strA(i), "°")
            If sCodLoc = strB(1) Then
                If flgPosu = "N" Then
                    codProductoPosuInk = objProducto.GetCodInka(strB(0))
                Else
                    codProductoPosuInk = objProducto.GetCodPosu(strB(0))
                End If
                jsonMaestroDetalle = jsonMaestroDetalle & "{" & _
                                                            "'codProducto':'" & codProductoPosuInk & "'," & _
                                                            "'quantity':'" & strB(2) & "'," & _
                                                            "'isFractional':'" & strB(3) & "'," & _
                                                            "'unitQuantity':'" & strB(4) & "'," & _
                                                            "'desProducto':'" & Replace(strB(5), "'", "") & "'" & _
                                                           "},"
                
            End If
        Next i
        If Len(jsonMaestroDetalle) > 0 Then
            jsonMaestroDetalle = left(jsonMaestroDetalle, Len(jsonMaestroDetalle) - 1)
            jsonMaestro = jsonMaestro & "{'cia':'" & sCia & "','numProforma':'" & strNumProforma & "','codLocal':'" & localSAP & "','codLocalBtl':'" & sCodLoc & "','posu':'" & flgPosu & "','reserveList':[" & jsonMaestroDetalle & "]},"
        End If
        Debug.Print jsonMaestro
    Next x
    'productosValidos => cadena con producto y locales, si es nulo entonces existen productos con error
    jsonMaestro = "{'data':[" & left(jsonMaestro, Len(jsonMaestro) - 1) & "]}"
    Debug.Print jsonMaestro
    '{'data':[{'cia':'94','numProforma':'0000001549','codLocal':'AY2','codLocalBtl':'S40','posu':'N','detalle':[]},{'cia':'94','numProforma':'0000001549','codLocal':'FN0','codLocalBtl':'AB9','posu':'F','detalle':[{'codProducto':'515335','quantity':'3','isFractional':'1','unitQuantity':'30','desProducto':'515335 - GLICENEX SR COM 750 mg CJA  x 30 COM'}]]}
    Set dicMaestro = JsonConverter.ParseJson(jsonMaestro)
    '-------------------------------------------
    Dim msgReserva As String
    Dim msgError As String
    Dim msgComma As String
    Debug.Print dicMaestro("data").Count()
'    If estConfig = 0 Then Exit Function
    msgReserva = objPedido.ReservaStock(1, dicMaestro)
    If Len(msgReserva) = 0 Then PreReservaStock = "A": Exit Function Else GoTo salir
salir:
'    MsgBox "No se procesa reserva de stock, por que existen productos con 'errores'.", vbOKOnly + vbCritical, App.ProductName:: grdPedidoStock.SetFocus
    grdPedidoStock.MoveFirst
    grdTransferencia.MoveFirst
    PreReservaStock = "No se procesa reserva de stock, por que existen productos con 'errores'."
    Exit Function
Err:
    Err.Raise Err.Number, "frm_DLV_Ruteo.PreReservaStock", Err.Description
End Function
'F.ECASTILLO 27.07.2020

Private Sub Form_Load()
 Dim rsZonaLocalAsig As oraDynaset
 Dim objLocal As New clsLocal
 Dim columna As TrueDBGrid70.Column
 Dim sCodLoc, sCia As String
 Dim rsCia As oraDynaset

 On Error GoTo Control

    ' get cia desde bd, jct 02Abr12
    ' B JCT, el combo de cia, debe ser segun local de despacho
    esLoad = True
    productosValidos = ""
    
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    'arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO_2", "Pedido", "STOCK", "FLG_FRACCIONAMIENTO", "CTD_FRACCIONAMIENTO", "DIF", "CTD_PRODUCTO", "CTD_PRODUCTOFRAC", "CTD_PRODUCTO1", "CTD_PRODUCTOFRAC1")
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Codigo", "Descripión", "Ctd.Pedido", "Ctd.Dispon", "flg_fraccionamiento", "ctd_fraccionamiento", "DIF", "CtdUnd", "CtdFra", "CantProducto1", "CantProductoFrac1", "DES_PRODUCTO_2")
    
    arrAncho = Array(0, 6650, 1000, 1000, 800, 800, 700, 0, 0, 0, 0, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    
    grdPedidoStock.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdPedidoStock.Columns("FLG_FRACCIONAMIENTO").Visible = False
    grdPedidoStock.Columns("CTD_FRACCIONAMIENTO").Visible = False
    grdPedidoStock.Columns("DIF").Visible = False
    'grdPedidoStock.Columns("CTD_PRODUCTO").Visible = False
    grdPedidoStock.Columns("CtdUnd").Visible = False
    'grdPedidoStock.Columns("CTD_PRODUCTO1").Visible = False
    grdPedidoStock.Columns("CantProducto1").Visible = False
    'grdPedidoStock.Columns("CTD_PRODUCTOFRAC").Visible = False
    grdPedidoStock.Columns("CtdFra").Visible = False
    'grdPedidoStock.Columns("CTD_PRODUCTOFRAC1").Visible = False
    grdPedidoStock.Columns("CantProductoFrac1").Visible = False
    grdPedidoStock.Columns(0).FetchStyle = True
    grdPedidoStock.Columns(1).FetchStyle = True
    grdPedidoStock.Columns(2).FetchStyle = True
    grdPedidoStock.Columns(3).FetchStyle = True
    grdPedidoStock.Columns(4).FetchStyle = False
    grdPedidoStock.Columns(5).FetchStyle = False
    grdPedidoStock.Columns(6).FetchStyle = False
    
    
    For Each columna In grdPedidoStock.Columns
            columna.AllowSizing = False
    Next
    
    sCodLoc = strLocalPedido
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, sCodLoc)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    'Else
    '  MsgBox "No se puede obtener la CIA...."
    '  Exit Sub
    End If
    Set rsCia = Nothing
    
    ' E JCT
    
    
    '*************** B JCT 27MAR12, CARGA DE CIA  ***********************************************
    'MsgBox "strLocalPedido : " + strLocalPedido
    
    Set cboCia.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCia.ListField = "Des"
    cboCia.BoundColumn = "Cod"
    
    
    
    ' cia para motorizado
    Set cboCiaMot.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCiaMot.ListField = "Des"
    cboCiaMot.BoundColumn = "Cod"
    
    ' cia leida de pedido, proforma
    Set cboCiaAsig.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCiaAsig.ListField = "Des"
    cboCiaAsig.BoundColumn = "Cod"
    
    
    
    cboCia.Enabled = False
    cboCiaMot.Enabled = False
    '****************E JCT 27MAR12 **********************************************
    
    
    'Set cboLocal.RowSource = objlocal.Lista(objUsuario.CodigoEmpresa, "")
    Set cboLocal.RowSource = objLocal.Lista(sCia, "")
    cboLocal.ListField = "local_dex2"
    cboLocal.BoundColumn = "COD_LOCAL"
    
    
    
    
    'Set cboLocalAsig2.RowSource = objlocal.Lista(objUsuario.CodigoEmpresa, "")
    Set cboLocalAsig2.RowSource = objLocal.Lista(sCia, "")
    cboLocalAsig2.ListField = "local_dex"
    cboLocalAsig2.BoundColumn = "COD_LOCAL"
    Set objLocal = Nothing
    
    Dim objMotorizado As New clsMotorizado
    Set cboMotorizado.RowSource = objMotorizado.Lista
    cboMotorizado.BoundColumn = "COD_MOTORIZADO"
    cboMotorizado.ListField = "NOMBRE"
    Set objMotorizado = Nothing
    
    
    
    Set cboRuta2.RowSource = objZona.Lista
    cboRuta2.BoundColumn = "COD_ZONA"
    cboRuta2.ListField = "DES_ZONA"
    
    Set cboRuta1.RowSource = objZona.Lista
    cboRuta1.BoundColumn = "COD_ZONA"
    cboRuta1.ListField = "DES_ZONA"
    
    
    
    Set rsPedido = objPedido.ListaCabecera(objUsuario.CodigoEmpresa, strLocal, strNumProforma)
    'Set rsPedido = objPedido.ListaCabecera(sCia, strLocal, strNumProforma)
    txtNumPedido.Text = "" & rsPedido("NUM_PROFORMA").Value
    txtNombreCliente.Text = "" & rsPedido("DES_CLIENTE").Value
    cboLocal.BoundText = "" & rsPedido("COD_LOCAL_REF").Value
    
    
    '+++++ asignar las 3 cia segun el local leido
    
    ' B JCT, el combo de cia, debe ser segun local de despacho
    
    
    
    sCodLoc = "" & rsPedido("COD_LOCAL_REF").Value
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, sCodLoc)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
      cboCia.BoundText = sCia
      cboCiaAsig.BoundText = sCia
      cboCiaMot.BoundText = sCia
      If cboCia.BoundText = "99" Then
         cmdStockEnLinea.Visible = True
      Else
         cmdStockEnLinea.Visible = False
      End If
 
     'Else
     ' MsgBox "No se puede Asignar la CIA...."
    End If
    Set rsCia = Nothing
    If objLocal.EvaluaLocalInkMig(sCodLoc) = "N" Then
        cmdTranferencia.Visible = False
    Else
        cmdTranferencia.Visible = True
    End If
    
    ' E JCT
    
    '''' chequear por que edwin dijo que si era la transferencia debería salir otros datos
    ''''If rsPedido("FLG_TRANSFERENCIA").Value = "1" Then cboLocal.BoundText = "" & rsPedido("COD_LOCAL").Value
    
    
    
    txtTeleoperadora.Text = "" & rsPedido("COD_USUARIO").Value & "-" & rsPedido("CAJERA").Value
    txtDocIdent.Text = "" & rsPedido("DES_TIPODOC").Value
    txtDireccionEntrega.Text = "" & rsPedido("DES_AUX_CLI_DIRECC").Value
    txtUrbanizacion.Text = "" & rsPedido("URBANIZACION").Value
    'cboLocal.BoundText = "" & rsPedido("COD_LOCAL_DESPACHO").Value
    txtUbigeo.Text = "" & rsPedido("distrito").Value
    txtReferencia.Text = "" & rsPedido("DES_AUX_CLI_REF").Value
    ctlTextBox12.Text = "" & rsPedido("OBS_NOTA_LOCAL").Value
    ctlTextBox1.Text = "" & rsPedido("OBS_NOTA_RUTEO").Value
    txtObsMotorizado.Text = "" & rsPedido("OBS_NOTA_MOTORIZADO").Value
    
'    If rsPedido("obs_nota_RUTEO").Value <> "" Then
'        ctlTextBox1.Text = "" & rsPedido("obs_nota_RUTEO").Value & " INT " & rsPedido("NUM_PROFORMA_ORIG").Value
'    End If
    
    CodCliente = rsPedido("COD_CLIENTE_DLV").Value
    CodDireccionCli = rsPedido("COD_DIRECCION_CLI").Value
    
    'Set rsZonaLocalAsig = objZona.ListaZonaLocal(objUsuario.CodigoEmpresa, cboLocal.BoundText)
    'Set rsZonaLocalAsig = objZona.ListaZonaLocal(sCia, cboLocalAsig.BoundText)
    
    Set rsZonaLocalAsig = objZona.ListaZonaLocal(sCia, cboLocal.BoundText)
    
    If Not rsZonaLocalAsig.EOF Then
        strZonaLocalAsignado = rsZonaLocalAsig("COD_ZONA").Value
    End If
    cboRuta1.BoundText = strZonaLocalAsignado
    
    'setear local asignado segun zona
    
    ' deberia ser local por zona, jct
    Set cboLocalAsig.RowSource = objZona.ListaLocal(sCia, cboRuta1.BoundText, "", sCia)
    cboLocalAsig.BoundColumn = "COD_LOCAL"
    cboLocalAsig.ListField = "local_dex"
    cboLocalAsig.BoundText = "" & rsPedido("COD_LOCAL_REF").Value
    
    
    
    
    ''''' La misma nota cuando un pedido es una transferencia
    '''''If rsPedido("FLG_TRANSFERENCIA").Value = "1" Then cboLocalAsig.BoundText = "" & rsPedido("COD_LOCAL").Value
        
    strCodEstado = rsPedido("COD_ESTADO").Value
    
    
    'buscaPedidoStock 'comentado por pherrera 230609 se ejecutaba dos veces al cargar
    
    'Set objVenta.Distribucion =
    'cboMotorizado.BoundText = "" & rsPedido("COD_MOTORIZADO").Value
    
    If Not IsNull(rsPedido("COD_MOTORIZADO").Value) Then
        Check2.Value = 1
        cboMotorizado.BoundText = "" & rsPedido("COD_MOTORIZADO").Value
    End If
        
    ''grdPedidoStock.FetchRowStyle = True
    
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Local", "Codigo", "Descripción", "U/F", "Cant.", "Precio", "xTipoventa", "UndFra", "Origen", "Destino", "Origen", "Destino")
    arrAncho = Array(0, 0, 6650, 500, 700, 0, 0, 0, 0, 0, 900, 900)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    grdTransferencia.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdTransferencia.Columns(0).Visible = False
    grdTransferencia.Columns(5).Visible = False
    grdTransferencia.Columns(6).Visible = False
    grdTransferencia.Columns(7).Visible = False
    grdTransferencia.Columns(8).Visible = False
    grdTransferencia.Columns(9).Visible = False
    
    For Each columna In grdTransferencia.Columns
            columna.AllowSizing = False
    Next
    
    objVenta.LimpiaDistribucion
    llenaTransferencias strLocal, strNumProforma
    grdTransferencia.Array1 = objVenta.Distribucion
    grdTransferencia.Rebind
    Picture1.Visible = False
    
    cmdAceptar.Enabled = True
    If bolModoConsulta Then cmdAceptar.Enabled = False
    
    If rsPedido("FLG_ENTREGA_TERCERO").Value = "1" Then
        Picture1.Visible = True
        arrCampos = Array("", "", "", "")
        arrCaption = Array("Nombre", "Dirección", "Ref", "Telf")
        arrAncho = Array(2800, 3000, 3250, 1200)
        arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft)
        grdEntregaTerceros.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        For Each columna In grdEntregaTerceros.Columns
            columna.AllowSizing = False
    
        Next
        
        Call EntregaTerceros(IIf(IsNull(rsPedido("DES_AUX_RECOGE_NOMBRE").Value), "", rsPedido("DES_AUX_RECOGE_NOMBRE").Value), _
                    IIf(IsNull(rsPedido("DES_AUX_RECOGE_DIRECC").Value), "", rsPedido("DES_AUX_RECOGE_DIRECC").Value), _
                    IIf(IsNull(rsPedido("DES_AUX_RECOGE_REF").Value), "", rsPedido("DES_AUX_RECOGE_REF").Value), _
                    IIf(IsNull(rsPedido("DES_AUX_RECOGE_TLF").Value), "", rsPedido("DES_AUX_RECOGE_TLF").Value))
    End If
    
    txtNombreCliente.Bloqueado = True
    txtNumPedido.Bloqueado = True
    txtTeleoperadora.Bloqueado = True
    txtDocIdent.Bloqueado = True
    txtDireccionEntrega.Bloqueado = True
    txtUbigeo.Bloqueado = True
    txtUrbanizacion.Bloqueado = True
    txtReferencia.Bloqueado = False
    chkAsignaSinValidar.Value = 0
    
    esLoad = False
    
    txtNumPedidoIkf.Text = "" & rsPedido("NUM_PEDIDO_REF").Value 'ECASTILLO 08.05.2020
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Sub buscaPedidoStock()
 Dim sCia As String
 Dim CodLocal As String
 Dim response As oraDynaset
 sCia = cboCia.BoundText ' combo de cia, JCT
 
    If Not cboLocalAsig.BoundText = "" Then
        If rsPedido("FLG_TRANSFERENCIA").Value = "1" Then
            'Set grdPedidoStock.DataSource = objPedido.ListaStock(objUsuario.CodigoEmpresa, strNumProforma, rsPedido("COD_LOCAL").Value)
            Set response = objPedido.ListaStock(sCia, strNumProforma, rsPedido("COD_LOCAL").Value, rsPedido("COD_LOCAL").Value)
            CodLocal = rsPedido("COD_LOCAL").Value
        Else
            'Set grdPedidoStock.DataSource = objPedido.ListaStock(objUsuario.CodigoEmpresa, strNumProforma, cboLocalAsig.BoundText)
            Set response = objPedido.ListaStock(sCia, strNumProforma, cboLocalAsig.BoundText, rsPedido("COD_LOCAL").Value)
            CodLocal = cboLocalAsig.BoundText
        End If
    Else
        If Not cboLocal.BoundText = "" Then
            'Set grdPedidoStock.DataSource = objPedido.ListaStock(objUsuario.CodigoEmpresa, strNumProforma, cboLocal.BoundText)
            Set response = objPedido.ListaStock(sCia, strNumProforma, cboLocal.BoundText)
            CodLocal = cboLocal.BoundText
        Else
            'Set grdPedidoStock.DataSource = objPedido.ListaStock(objUsuario.CodigoEmpresa, strNumProforma, "")
            Set response = objPedido.ListaStock(sCia, strNumProforma, "")
            CodLocal = cboLocalAsig.BoundText
        End If
    End If
    Dim record As New XArrayDB
    Dim Producto() As String
    Dim responseProducto As Dictionary
    Dim i, x, a, b As Variant
    Dim aa, bb As Variant
    i = 0
    record.ReDim 0, -1, 0, 11
    'Debug.Print response.RecordCount
    ReDim Producto(response.RecordCount - 1)
    'Producto(0) = "a"
    With response
        .MoveFirst
        While Not .EOF
            'Debug.Print .Fields("COD_PRODUCTO")
            Producto(i) = .Fields("COD_PRODUCTO")
            i = i + 1
            .MoveNext
        Wend
    End With
    Dim codProductoSap As String
    Dim StockReal, flg_frac, ctd_frac As String
    Dim c, d As Variant
    'I. ECASTILLO 13.01.2021 agregar flg para activar
    Dim flgFun
    flgFun = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV") '1 => ACTIVO, 0 => INACTIVO
    If flgFun = "1" Then
        getStockReal sCia, CodLocal, Producto
    ElseIf flgFun = "0" Then
        StockBD CodLocal, Producto
    End If
    'F. ECASTILLO 13.01.2021
    i = 0
    'Debug.Print arrInfo.Count(1)
    'Debug.Print arrInfo
    With response
        .MoveFirst
        While Not .EOF
            a = 0: b = 0: flg_frac = 0: ctd_frac = 1: c = 0: d = 0
            flg_frac = .Fields("FLG_FRACCIONAMIENTO").Value: ctd_frac = .Fields("CTD_FRACCIONAMIENTO").Value 'DEFINIR EN CASO NO EXISTA EN WEBSERVICE
            'Debug.Print .Fields("COD_PRODUCTO").Value
            codProductoSap = objProducto.GetCodPosu(.Fields("COD_PRODUCTO").Value)
            StockReal = 0 '.Fields("STOCK").Value
            For x = 0 To arrInfo.Count(1) - 1
                'Debug.Print arrInfo(x, 2) & vbNewLine & arrInfo(x, 3) & vbNewLine & arrInfo(x, 4)
                If codProductoSap = arrInfo(x, 0) Then StockReal = IIf(arrInfo(x, 1) = 0, arrInfo(x, 2), arrInfo(x, 2) & IIf(arrInfo(x, 4) > 0, "F" & arrInfo(x, 4), "")): flg_frac = arrInfo(x, 1): ctd_frac = arrInfo(x, 3)
            Next x
            StockReal = IIf(Len(StockReal) = 0, 0, StockReal)
            record.AppendRows
            record(i, 0) = .Fields("COD_PRODUCTO").Value
            record(i, 1) = .Fields("DES_PRODUCTO_2").Value
            record(i, 2) = .Fields("PEDIDO").Value
            record(i, 3) = StockReal '.Fields("STOCK").Value
            record(i, 4) = flg_frac '.Fields("FLG_FRACCIONAMIENTO").Value 'ESTO DEBE SER REEMPLAZADO POR VALOR DE WS
            record(i, 5) = ctd_frac '.Fields("CTD_FRACCIONAMIENTO").Value 'ESTO DEBE SER REEMPLAZADO POR VALOR DE WS
            record(i, 7) = .Fields("CTD_PRODUCTO").Value
            record(i, 8) = .Fields("CTD_PRODUCTOFRAC").Value
            record(i, 9) = .Fields("CTD_PRODUCTO1").Value 'PEDIDO_ENTERO
            record(i, 10) = .Fields("CTD_PRODUCTO_FRAC1").Value 'PEDIDO_FRACCION
            record(i, 11) = .Fields("DES_PRODUCTO_2").Value
            a = CDec(val(record(i, 9)) * val(record(i, 5)) + val(record(i, 10)))
            'Debug.Print Replace(CStr(record(i, 3)), "F", ".")
            bb = Split(record(i, 3), "F")
            c = CDec(bb(0)) ' entero
            If UBound(bb) > 0 Then 'hay decimales
                d = CDec(bb(1))
            End If
            'd = 0.17
            'b = Replace(CStr(record(i, 3)), "F", ".") * Val(record(i, 5))
            b = CDec((c * record(i, 5)) + d)
            'b = 0.17
            'Debug.Print b
            record(i, 6) = b - a '.Fields("DIF").Value
            
'            record(i, 3) = Replace(CStr(Val(StockReal)), ".", "F")
            
            'Debug.Print record(i, 0) & "|" & record(i, 1) & "|" & record(i, 2) & "|" & record(i, 3) & "|" & record(i, 6)
            i = i + 1
            .MoveNext
        Wend
    End With
    'Debug.Print record(0, 1)
    grdPedidoStock.Array1 = record
    grdPedidoStock.Rebind
End Sub
'I.ECASTILLO 13.01.2021
Private Function StockBD(CodigoLocal As String, codigoProducto As Variant)
On Error GoTo Err
    Dim IsFractional As Integer
    Dim codProductoPosu, codLocalPosu As String
    Dim x, i As Integer
    
    codLocalPosu = objLocal.GetCodPosu(CodigoLocal)
    For x = 0 To UBound(codigoProducto)
        codProductoPosu = codProductoPosu & objProducto.GetCodPosu(codigoProducto(x)) & "#$"
    Next x
    codProductoPosu = left(codProductoPosu, Len(codProductoPosu) - 2)
    Dim arrResp As New XArrayDB
    Dim rsData As oraDynaset
    
    Set rsData = objProducto.ListaStockLocal(codLocalPosu, codProductoPosu)
    arrInfo.ReDim 0, -1, 0, 4
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        Dim ultimo As Integer
        While Not rsData.EOF
            ultimo = arrInfo.Count(1)
            arrInfo.AppendRows
            IsFractional = "" & rsData("FLG_FRACCIONAMIENTO")
            arrInfo(ultimo, 0) = "" & rsData("COD_PROD")
            arrInfo(ultimo, 1) = IsFractional
            arrInfo(ultimo, 2) = "" & rsData("STOCK_ENTERO") 'STOCK_ENTERO
            arrInfo(ultimo, 3) = "" & rsData("VAL_FRAC_LOCAL") 'CNT_FRACCION
            arrInfo(ultimo, 4) = "" & (rsData("STOCK_FRACCION") - (rsData("STOCK_ENTERO") * rsData("VAL_FRAC_LOCAL"))) 'STOCK_FRACCION
            rsData.MoveNext
        Wend
    End If
    Exit Function
Err:
    Err.Raise Err.Number, "frm_VTA_CantidadProducto", Err.Description
End Function
Private Function validaStock() As String
    Dim x, i, z, j As Integer
    Dim arrProductos As New XArrayDB
    Dim arrTransferencia As New XArrayDB
    Dim arrDuplicados As New XArrayDB
    Dim strProductosStock, codProducto, CodLocal, COMMA, ctd As String
    Dim msg As String
    Dim cab As String
    Dim strA, strB, strC As Variant
    Dim Falta As Integer
    Dim a, b, c, d, e, Existe As Integer
    Dim aa, bb
    Dim ii As Integer
    Dim msgCustom As String
    
    cab = "": msg = "": x = 0: i = 0: z = 0: ctd = 0: strProductosStock = ""
    grdPedidoStock.MoveFirst
    arrProductos.ReDim 0, -1, 0, 6
    'Array("Codigo", "Descripión", "Ctd.Pedido", "Ctd.Dispon", "flg_fraccionamiento", & _
           "ctd_fraccionamiento", "DIF", "CtdUnd", "CtdFra", "CantProducto1", "CantProductoFrac1", "DES_PRODUCTO_2")
    For x = 0 To grdPedidoStock.Array1.Count(1) - 1
        Debug.Print grdPedidoStock.Columns("CantProducto1").Value & vbNewLine & grdPedidoStock.Columns("ctd_fraccionamiento").Value & vbNewLine & grdPedidoStock.Columns("CantProductoFrac1").Value & _
                    vbNewLine & vbNewLine & grdPedidoStock.Columns("Ctd.Dispon").Value & vbNewLine & grdPedidoStock.Columns("ctd_fraccionamiento").Value
        a = 0: b = 0: codProducto = "": CodLocal = "": COMMA = "": aa = "": bb = 0
        a = val(grdPedidoStock.Columns("CantProducto1").Value) * val(grdPedidoStock.Columns("ctd_fraccionamiento").Value) + val(grdPedidoStock.Columns("CantProductoFrac1").Value) 'WMORI 24/09/2020 SE MODIFICA POR SOLICITUD DE BREISNER
'        a = Replace(CStr(grdPedidoStock.Columns("Ctd.Pedido").Value), "F", ".") * grdPedidoStock.Columns("ctd_fraccionamiento").Value
        aa = Split(grdPedidoStock.Columns("Ctd.Dispon").Value, "F")
        If UBound(aa) > 0 Then
            bb = aa(1)
        End If
        b = (aa(0) * val(grdPedidoStock.Columns("ctd_fraccionamiento").Value)) + bb
'        b = grdPedidoStock.Columns("Ctd.Dispon").Value * grdPedidoStock.Columns("ctd_fraccionamiento").Value
        codProducto = grdPedidoStock.Columns("Codigo").Value
        CodLocal = cboLocalAsig.BoundText
        ctd = grdPedidoStock.Columns("DIF")
        'a => cant_pedido, b => cant_stock
        If ctd < 0 Then 'no cumple con stock
            arrProductos.AppendRows
            arrProductos(i, 0) = codProducto
            arrProductos(i, 1) = a - b 'ctd 'Cantidad Falta?
            arrProductos(i, 2) = grdPedidoStock.Columns("Descripión").Value
            arrProductos(i, 3) = CodLocal
            arrProductos(i, 4) = a
            arrProductos(i, 5) = grdPedidoStock.Columns("flg_fraccionamiento").Value
            arrProductos(i, 6) = grdPedidoStock.Columns("ctd_fraccionamiento").Value
            If Abs(arrProductos(i, 1)) <> arrProductos(i, 4) Then 'cumple parcialmente
                strProductosStock = strProductosStock & _
                                codProducto & "°" & _
                                CodLocal & "°" & _
                                a & "°" & _
                                grdPedidoStock.Columns("flg_fraccionamiento").Value & "°" & _
                                grdPedidoStock.Columns("ctd_fraccionamiento").Value & "°" & _
                                grdPedidoStock.Columns("Descripión").Value & "°1°" & _
                                a - b & _
                                "|"
            End If
            i = i + 1
        Else 'si cumple con stock
            strProductosStock = strProductosStock & _
                                codProducto & "°" & _
                                CodLocal & "°" & _
                                a & "°" & _
                                grdPedidoStock.Columns("flg_fraccionamiento").Value & "°" & _
                                grdPedidoStock.Columns("ctd_fraccionamiento").Value & "°" & _
                                grdPedidoStock.Columns("Descripión").Value & "°1°" & _
                                a - b & _
                                "|"
        End If
        grdPedidoStock.MoveNext
    Next x
    x = 0: arrDuplicados.ReDim 0, -1, 0, 3
    Debug.Print arrProductos.Count(1)
    While x < arrProductos.Count(1)
        'Codigo, Cant.
        c = 0: d = 0: Existe = 0: i = 0
        arrTransferencia.ReDim 0, -1, 0, 5
        Debug.Print grdTransferencia.Array1.Count(1)
        Debug.Print arrProductos.Count(1)
        grdTransferencia.MoveFirst
        ii = 0
        For i = 0 To grdTransferencia.Array1.Count(1) - 1
            If arrProductos(x, 0) = grdTransferencia.Columns("Codigo").Value Then
                z = 0
                While z < arrTransferencia.Count(1) And Existe = 0
                    If (arrTransferencia(z, 0) = arrProductos(x, 0)) _
                        And (arrTransferencia(z, 2) = grdTransferencia.Columns("Origen").Value) Then Existe = 1: z = z - 1
                    z = z + 1
                Wend
                If Existe = 0 Then
                    arrTransferencia.AppendRows
                    arrTransferencia(ii, 0) = arrProductos(x, 0)
                    arrTransferencia(ii, 1) = grdTransferencia.Columns("Cant.").Value
                    arrTransferencia(ii, 2) = grdTransferencia.Columns("Origen").Value
                    arrTransferencia(ii, 3) = arrProductos(x, 6)
                    arrTransferencia(ii, 4) = grdTransferencia.Columns("U/F").Value
                    arrTransferencia(ii, 5) = 0
                    ii = ii + 1
'                Else
'                    If arrTransferencia(z, 1) <> grdTransferencia.Columns("Cant.").Value Then
'                        arrTransferencia(z, 1) = CInt(arrTransferencia(z, 1)) + CInt(grdTransferencia.Columns("Cant.").Value)
'                    End If
                End If
            End If
            grdTransferencia.MoveNext
        Next i
        z = 0
        Debug.Print arrTransferencia.Count(1)
        While z < arrTransferencia.Count(1)
            i = 0: c = 0: d = 0
            While i < arrTransferencia.Count(1) And d = 0
                If arrTransferencia(z, 0) = arrTransferencia(i, 0) Then
                    If (arrTransferencia(z, 5) <> 0 Or arrTransferencia(i, 5) <> 0) Then
                        d = 1: c = 0 'producto ya fue procesado
                    Else
                        e = IIf(arrTransferencia(z, 4) = "U", arrTransferencia(i, 1) * arrTransferencia(i, 3), arrTransferencia(i, 1))
                        d = 0: c = c + e
                    End If
                End If
                i = i + 1
            Wend
            If d = 0 Then
                If arrProductos(x, 0) = arrTransferencia(z, 0) Then
                    ' 17.12.2020 | se agrego Or para que permita transferir más de lo faltante
                    If (arrProductos(x, 1) = c Or arrProductos(x, 1) < c) And "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "VLDSTKDLV+") = "0" Then
                        arrTransferencia(z, 5) = 1
'                        strProductosStock = strProductosStock & arrTransferencia(z, 0) & "-" & arrTransferencia(z, 2) & ","
                        a = IIf(arrTransferencia(z, 4) = "U", arrTransferencia(z, 1) * arrTransferencia(z, 3), arrTransferencia(z, 1))
                        strProductosStock = strProductosStock & _
                                    arrTransferencia(z, 0) & "°" & _
                                    arrTransferencia(z, 2) & "°" & _
                                    a & "°" & _
                                    arrProductos(x, 5) & "°" & _
                                    arrProductos(x, 6) & "°" & _
                                    arrProductos(x, 2) & "°2°0" & _
                                    "|"
                        arrProductos.DeleteRows (x)
                        x = x - 1
                    ElseIf arrProductos(x, 1) < c Then
                        x = x
                        If "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "VLDSTKDLV+") = "1" Then
                            msgCustom = "Se esta intentado realizar transferencia que supera lo solicitado."
                            GoTo errorCustom
                        End If
                    Else
                        arrDuplicados.AppendRows
                        arrDuplicados(x, 0) = arrProductos(x, 0) 'codigo
                        arrDuplicados(x, 1) = arrProductos(x, 2) 'descripcion
                        arrDuplicados(x, 2) = arrProductos(x, 3) 'local
                    End If
                End If
            End If
            z = z + 1
        Wend
        x = x + 1
    Wend
    x = 0
    While x < arrProductos.Count(1)
        If cab = "" Then cab = "Falta transferencia por insuficiencia de Stock para los siguientes Productos: " & vbNewLine & "===========================================": msg = cab & vbNewLine
        msg = msg & arrProductos(x, 2)
        x = x + 1
    Wend
    'Detectar productos con stock pero con transferencia
    Debug.Print strProductosStock
    If "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "VLDSTKDLV+") = "1" Then
    strA = Split(strProductosStock, "|")
    x = 0
    While x <= UBound(strA) - 1
        Debug.Print (strA(x))
        strB = Split(strA(x), "°")
        i = 0
        Debug.Print strB(6)
        If strB(6) = 1 Then
            i = 0: c = 0: e = 0: d = 0: CodLocal = ""
            For i = 0 To grdTransferencia.Array1.Count(1) - 1
                If strB(0) = grdTransferencia.Columns("Codigo").Value Then
'                    CodLocal = grdTransferencia.Columns("Origen").Value
                    e = IIf(grdTransferencia.Columns("U/F").Value = "U", grdTransferencia.Columns("Cant.").Value * strB(4), grdTransferencia.Columns("Cant.").Value)
                    c = c + e
                End If
            Next i
            i = 0
            
            For i = 0 To grdTransferencia.Array1.Count(1) - 1
                Debug.Print strB(2) 'Cantidad Pedida
                Debug.Print strB(7) 'Cantidad Faltante
                If (strB(0) = grdTransferencia.Columns("Codigo").Value) Then
                    If strB(7) = 0 Then 'Si no le falta stock
                        arrDuplicados.AppendRows
                        arrDuplicados(x, 0) = strB(0)
                        arrDuplicados(x, 1) = strB(5)
                        arrDuplicados(x, 2) = strB(1)
                    End If
'                    CodLocal = grdTransferencia.Columns("Origen").Value
'                    e = IIf(grdTransferencia.Columns("U/F").Value = "U", grdTransferencia.Columns("Cant.").Value * strB(i, 4), grdTransferencia.Columns("Cant.").Value)
'                    c = c + e
                    If CInt(strB(7)) = c Then
'                        strProductosStock = strProductosStock & _
'                                        strB(0) & "°" & _
'                                        grdTransferencia.Columns("Origen").Value & "°" & _
'                                        IIf(grdTransferencia.Columns("U/F").Value = "U", grdTransferencia.Columns("Cant.").Value * strB(4), grdTransferencia.Columns("Cant.").Value) & "°" & _
'                                        strB(3) & "°" & _
'                                        strB(4) & "°" & _
'                                        strB(5) & "°2°0" & _
'                                        "|"
                    ElseIf CInt(strB(7)) < c Then
                        msgCustom = "Se esta intentado realizar transferencia que supera lo solicitado."
                        GoTo errorCustom
                    End If
                End If
            Next i
        End If
        x = x + 1
    Wend
    End If
    strA = "": strB = ""
    Debug.Print strProductosStock
    strA = Split(strProductosStock, "|")
    strProductosStock = ""
    x = 0
    While x <= UBound(strA) - 1
        Debug.Print (strA(x))
        strB = Split(strA(x), "°")
        i = 0: a = 0
        If strB(6) = 1 Then
            While i <= UBound(strA) - 1
                strC = Split(strA(i), "°")
                j = 0
                If strC(6) = 2 And strB(0) = strC(0) Then 'transf
                    a = a + strC(2)
                End If
                i = i + 1
            Wend
            strProductosStock = strProductosStock & _
                                strB(0) & "°" & _
                                strB(1) & "°" & _
                                strB(2) - a & "°" & _
                                strB(3) & "°" & _
                                strB(4) & "°" & _
                                strB(5) & "°" & _
                                strB(6) & "°" & _
                                strB(7) & "|"
        Else
            strProductosStock = strProductosStock & _
                                strB(0) & "°" & _
                                strB(1) & "°" & _
                                strB(2) & "°" & _
                                strB(3) & "°" & _
                                strB(4) & "°" & _
                                strB(5) & "°" & _
                                strB(6) & "°" & _
                                strB(7) & "|"
        End If
        x = x + 1
    Wend
    Debug.Print strProductosStock
    Dim cab2, msg2 As String
    x = 0: cab2 = "": msg2 = ""
    While x < arrDuplicados.Count(1)
        If cab2 = "" Then cab2 = "Lo siguientes productos no necesitan transferencia: " & vbNewLine & "===========================================": msg2 = cab2 & vbNewLine
        msg2 = msg2 & arrDuplicados(x, 2) & "|" & arrDuplicados(x, 1) & vbNewLine & vbNewLine & "Desea continuar?"
        x = x + 1
    Wend
    If Not cab2 = "" Then
        If MsgBox(msg2, vbYesNo + vbInformation) = vbNo Then validaStock = "Corregir transferencias": Exit Function
    End If
    If Not cab = "" Then validaStock = msg Else productosValidos = strProductosStock: Debug.Print strProductosStock: Exit Function
    Exit Function
errorCustom:
    MsgBox msgCustom, vbCritical, App.CompanyName
    validaStock = "Corregir data"
    Exit Function
End Function

Public Function getStockReal(Cia As String, CodigoLocal As String, codigoProducto As Variant)
    Dim obj As New Dictionary
    Dim ArrCode As Variant
    Dim x, i As Integer
    Dim codProductoPosu, codLocalPosu As String
    Dim vSplit As Variant
    Dim vStockF As Double
    Dim a, b, c, d As Double
    Debug.Print UBound(codigoProducto)
    ReDim ArrCode(UBound(codigoProducto))
    For x = 0 To UBound(codigoProducto)
        Debug.Print codigoProducto(x)
        codProductoPosu = objProducto.GetCodPosu(codigoProducto(x))
        ArrCode(x) = codProductoPosu
    Next x
    codLocalPosu = objLocal.GetCodPosu(CodigoLocal)
    'ArrCode = Array(codProductoPosu)
    'If CIA = "94" Then CIA = "1DLV" Else CIA = "0DLV"
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    Dim sCia As String
    sCia = "" & objLocal.GetMarcaLocalPosu(codLocalPosu, 1)
    sCia = Trim(sCia)
    'F.ECASTILLO 17.09.2021
    Set obj = objWS.GetStockRealWS(Cia, codLocalPosu, ArrCode, sCia)
    arrInfo.ReDim 0, -1, 0, 4
    If obj.Count > 0 Then
        Debug.Print obj("data").Count()
        For x = 1 To obj("data").Count()
            'For i = 1 To obj("data")(X)("fractionType").Count()
                arrInfo.AppendRows
                arrInfo(x - 1, 0) = obj("data")(x)("productId")
                arrInfo(x - 1, 1) = obj("data")(x)("isFractional")
                'If arrInfo(X - 1, 1) = 0 Then
                    i = 1
                    For i = 1 To obj("data")(x)("fractionType").Count()
                        If obj("data")(x)("fractionType")(i)("fractionatedText") = "PACK_MODE" Then
                            vStockF = obj("data")(x)("fractionType")(i)("stock")
                            a = obj("data")(x)("fractionType")(i)("stock") 'stock_entero
                            c = obj("data")(x)("fractionType")(i)("unitQuantity") 'cnt_fraccionamiento
                        End If
                        
                        If obj("data")(x)("fractionType")(i)("fractionatedText") = "PART_MODE" Then
                            b = obj("data")(x)("fractionType")(i)("stock") 'stock_fraccion
                            'c = obj("data")(X)("fractionType")(i)("unitQuantity") 'cnt_fraccionamiento
                        End If
                    Next i
                'd = b - (a * c)
                arrInfo(x - 1, 2) = a 'STOCK_ENTERO
                arrInfo(x - 1, 3) = c 'CNT_FRACCION
                arrInfo(x - 1, 4) = b - (a * c) 'STOCK_FRACCION
                'Else
                    'Debug.Print (obj("data")(X)("fractionType").Count())
                    'If obj("data")(X)("fractionType").Count() = 2 Then
                    '    i = 1
                    '    For i = 1 To obj("data")(X)("fractionType").Count()
                    '        If obj("data")(X)("fractionType")(i)("fractionatedText") = "PART_MODE" Then
                    '            a = obj("data")(X)("fractionType")(i)("stock") 'stock_fraccion
                    '            c = obj("data")(X)("fractionType")(i)("unitQuantity") 'cnt_fraccionamiento
                    '        End If
                    '    Next i
                    '    i = 1
                    '    For i = 1 To obj("data")(X)("fractionType").Count()
                    '        If obj("data")(X)("fractionType")(i)("fractionatedText") = "PACK_MODE" Then
                    '            b = obj("data")(X)("fractionType")(i)("unitQuantity") 'cnt_fraccionamiento
                    '        End If
                    '    Next i
                    '    vStockF = a / b ' 49 / 80 | 2139 / 200
                    '    arrInfo(X - 1, 2) = Int(vStockF) ' 49 |
                    '    arrInfo(X - 1, 3) = b
                    'End If
                'End If
                'a = 0: b = 0: c = 0: d = 0
                'vSplit = Split(CStr(vStockF), ".")
                'If UBound(vSplit) > 0 Then 'si tiene fraccion, muestra entero
                '    arrInfo(X - 1, 4) = right(vSplit(1), 2) '+ (obj("data")(X)("fractionType")(1)("stock") Mod arrInfo(X - 1, 3))
                'Else
'                    If obj("data")(X)("fractionType").Count() = 2 Then
                '        i = 1
                '        For i = 1 To obj("data")(X)("fractionType").Count()
                '            If obj("data")(X)("fractionType")(i)("fractionatedText") = "PACK_MODE" Then
                '                a = obj("data")(X)("fractionType")(1)("stock")
                '            End If
                '        Next i
                '        i = 1
                '        For i = 1 To obj("data")(X)("fractionType").Count()
                '            If obj("data")(X)("fractionType")(i)("fractionatedText") = "PART_MODE" Then
                '                a = obj("data")(X)("fractionType")(i)("stock")
                '            End If
                '        Next i
'                    End If
                '    i = 1
'                    For i = 1 To obj("data")(X)("fractionType").Count()
'                        If obj("data")(X)("fractionType")(i)("fractionatedText") = "PART_MODE" Then
'                            a = obj("data")(X)("fractionType")(i)("stock")
'                        End If
'                    Next i

                '    Dim iDivisor As Double
                '    iDivisor = Val(arrInfo(X - 1, 3))

                '    If iDivisor = 0 Then iDivisor = 1
                '        arrInfo(X - 1, 4) = (a Mod iDivisor)
                '    End If
            'Next i
        Next x
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set objPedido = Nothing
End Sub

Function llenaTransferencias(CodigoLocal As String, NumeroProforma As String) As XArrayDB
    Dim rsTransferencias As oraDynaset
    Dim i As Integer

    Set rsTransferencias = objPedido.ListaTransferencia(objUsuario.CodigoEmpresa, CodigoLocal, NumeroProforma)

    While Not rsTransferencias.EOF
        objVenta.AgregaDistribucion rsTransferencias("COD_LOCAL").Value, _
                                    rsTransferencias("COD_PRODUCTO").Value, _
                                    rsTransferencias("DES_PRODUCTO").Value, _
                                    IIf(val(rsTransferencias("CTD_PRODUCTO").Value) > 0, val(rsTransferencias("CTD_PRODUCTO").Value), val(rsTransferencias("CTD_PRODUCTO_FRAC").Value)), _
                                    IIf(val(rsTransferencias("CTD_PRODUCTO").Value) > 0, "0", "1"), _
                                    0, _
                                    Pedido_DLV, _
                                    rsTransferencias("COD_LOCAL_ORIGEN").Value, _
                                    rsTransferencias("COD_LOCAL_DESTINO").Value, _
                                    rsTransferencias("COD_LOCAL_SAP_ORIGEN").Value, _
                                    rsTransferencias("COD_LOCAL_SAP_DESTINO").Value
        rsTransferencias.MoveNext
    Wend
    
    
'''CodigoLocal As String, _
'''                            Codigo As String, _
'''                            Descripcion As String, _
'''                            Cantidad As String, _
'''                            FlagFraccion As String, _
'''                            Precio As Double, _
'''                            TipoVenta As TipoVenta, _
'''                            CodigoLocalOrigen As String, _
'''                            CodigoLocalDestino As String
    
    
    
      
End Function

Private Sub Grabar()
'On Error GoTo handle
    Dim objProforma As New clsProforma
    Dim x As String
    Dim strCadCodProducto As String
    Dim strCadCtdProducto  As String
    Dim strCadCtdProductoFrac  As String
    Dim strLocalOrigen  As String
    Dim strLocalDestino As String
    Dim i As Integer
    Dim strAsigna As String
    Dim msgValidaStock, msgPreReservaStock As String
    Dim estConfig As String
    Dim rsCia As oraDynaset
    Dim sCia As String
    Dim posu As String
    Dim flgReserva As String
    i = 0: estConfig = 0
    
    
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, strLocalPedido)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    estConfig = objLocal.GetEstConfig(sCia, strLocalPedido, "RESERVA_STOCK")
    posu = objLocal.EvaluaLocalInkMig(strLocalPedido)
    flgReserva = objProforma.getFlgReserva("10", strLocal, strNumProforma)
    strAsigna = "SI"
    'I.ECASTILLO 27.07.2020 | 13.01.2021 agregar flg para activar
    Dim flgFun
    flgFun = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV") '1 => ACTIVO, 0 => INACTIVO
    If flgFun = "1" Then
        If strFlgTransf = 0 And flgReserva = 0 And frm_DLV_Seguimiento.Option1.Value <> True Then
            msgValidaStock = validaStock
            If Not msgValidaStock = "" Then Err.Raise vbObjectError + 1, "Ruta", msgValidaStock: grdPedidoStock.MoveFirst: grdTransferencia.MoveFirst
            'If estConfig <> 0 Or posu = "N" Then
            If estConfig <> 0 Then
                'A => RESERVO, B => NO RESERVA, CADENA => ERROR
                msgPreReservaStock = PreReservaStock
                If msgPreReservaStock = "A" Then
                ElseIf msgPreReservaStock = "B" Then
                    estConfig = 0
                Else
                    Err.Raise vbObjectError + 1, "Ruta", msgPreReservaStock
                End If
            End If
        End If
    End If
    'F.ECASTILLO 27.07.2020
    While i <= objVenta.Distribucion.UpperBound(1)
        strCadCodProducto = strCadCodProducto & objVenta.Distribucion(i, 1) & "|"
        If objVenta.Distribucion(i, 7) = 0 Then
            strCadCtdProducto = strCadCtdProducto & objVenta.Distribucion(i, 4) & "|"
            strCadCtdProductoFrac = strCadCtdProductoFrac & "0" & "|"
        Else
            strCadCtdProducto = strCadCtdProducto & "0" & "|"
            strCadCtdProductoFrac = strCadCtdProductoFrac & objVenta.Distribucion(i, 4) & "|"
        End If
        
        strLocalOrigen = strLocalOrigen & objVenta.Distribucion(i, 8) & "|"
        strLocalDestino = strLocalDestino & objVenta.Distribucion(i, 9) & "|"
        i = i + 1
    Wend
    If Not objVenta.Distribucion.UpperBound(1) = -1 Then
        x = objProforma.GrabaTransferencia(objUsuario.CodigoEmpresa, strLocal, strNumProforma, strCadCodProducto, strCadCtdProducto, strCadCtdProductoFrac, strLocalOrigen, strLocalDestino, objUsuario.Codigo)
    End If
    
    strLocalPedido = cboLocalAsig.BoundText
    
   ' If Check2.Value = 1 Then
        If Not x = "" Then Err.Raise vbObjectError + 1, "Ruta", x 'GoTo handle
        Dim NumeroTransferencia As String
        x = objProforma.Asigna(objUsuario.CodigoEmpresa, strLocal, strNumProforma, objUsuario.Codigo, cboMotorizado.BoundText, objUsuario.NombrePC, ctlTextBox1.Text, strLocalPedido, CodCliente, CodDireccionCli, txtReferencia.Text, strAsigna, ctlTextBox12.Text, chkAsignaSinValidar.Value, IIf(estConfig = 0, "", estConfig))
   ' End If
    Set objProforma = Nothing
    If Not x = "" Then Err.Raise vbObjectError + 1, "Ruta", x 'GoTo handle
    intErr = 0
    'MsgBox "Se grabó satisfactoriamente la Proforma :" & strNumProforma, vbInformation
'    Exit Sub
'handle:
    'Err.Raise x
    'intErr = 1
    'Set objProforma = Nothing
End Sub

Private Sub grdPedidoStock_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
Dim n As Double
Dim s As Double
Dim f As Integer
Dim e As Integer
    
    If Condition = 0 Then
        Select Case Col
               Case 0, 1, 2, 3
                    n = val(IIf(IsNull(grdPedidoStock.Columns("DIF").CellText(Bookmark)), 0, grdPedidoStock.Columns("DIF").CellText(Bookmark)))
                    If n < 0 Then
                        CellStyle.ForeColor = vbRed
                        CellStyle.Font.Bold = True
                    End If
        End Select
    End If
    If Condition = 2 Or Condition = 3 Then
        Select Case Col
               Case 0, 1, 2, 3
                    n = val(IIf(IsNull(grdPedidoStock.Columns("DIF").CellText(Bookmark)), 0, grdPedidoStock.Columns("DIF").CellText(Bookmark)))
                    If n < 0 Then
                        CellStyle.ForeColor = vbYellow
                        CellStyle.BackColor = &H8000000D
                        CellStyle.Font.Bold = True
                    End If
        End Select
    End If
End Sub

Private Sub grdPedidoStock_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
On Error GoTo handle
    If val(grdPedidoStock.Columns(6).CellText(Bookmark)) < 0 Then
        RowStyle.BackColor = RGB(240, 128, 128)
    End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdPedidoStock_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo CtrlErr

'Comentado por Jahzeel por indicación de Pablo

'/*    If rsPedido("FLG_ENTREGA_TERCERO").Value = "0" Then
'        If Val(IIf(IsNull(grdPedidoStock.Columns("DIF").Value), 0, grdPedidoStock.Columns("DIF").Value)) < 0 Then
'            If frm_DLV_Seguimiento.Option1.Value = False Then
'                cmdTranferencia.Enabled = True
'            End If
'            If frm_DLV_Seguimiento.Option2.Value = True Then
'                If frm_DLV_Seguimiento.grdPedidos.Columns("FLG_TRANSFERENCIA") = 1 Then
'                    cmdTranferencia.Enabled = False
'                Else
'                    cmdTranferencia.Enabled = True
'                End If
'            End If
'        Else
'            cmdTranferencia.Enabled = False
'        End If
'    Else
'        cmdTranferencia.Enabled = True
'    End If */
    
    cmdTranferencia.Enabled = True
    
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub grdTransferencia_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo handle

    Select Case KeyCode
           Case vbKeyDelete
                If grdTransferencia.ApproxCount = 0 Then
                    Exit Sub
                End If
                grdTransferencia.Delete
                
                objPedido.EliminaTra objUsuario.CodigoEmpresa, _
                                    objUsuario.CodigoLocal, _
                                    Replace(txtNumPedido.Text, "-", "")
                                    
                EvaluaLocalesTransf
                
    End Select
Exit Sub
handle:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.ProductName
End Sub


Private Sub EntregaTerceros(ByVal Nombre As String, ByVal Direccion As String, ByVal Referencia As String, ByVal Telefono As String)
    On Error GoTo CtrlErr
    grdEntregaTerceros.Array1 = objVenta.AgregaEntregaTerceros(Nombre, Direccion, Referencia, Telefono)
    grdEntregaTerceros.Rebind
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName

End Sub

Private Sub ActObsRuta(ByVal strCia As String, _
                        ByVal strLocal As String, _
                        ByVal strNumProforma As String, _
                        ByVal strObsNotaLocal As String, _
                        ByVal strObsNotaRuta As String, _
                        ByVal strObsMotorizado As String)
                        
                        
    On Error GoTo handle
        
            objPedido.ActObsRuta strCia, strLocal, strNumProforma, strObsNotaLocal, strObsNotaRuta, strObsMotorizado
            
            MsgBox "Se actualizo las Observaciones", vbInformation + vbOKOnly, App.ProductName
                        
    Exit Sub
    
handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub sub_EnviarMensaje(ByVal vstrCod_Local As String)
'1. Cargar las maquinas por local
'2. A todas las que tienen IP pasar el mensaje
'Dim i As Integer
Dim host As String
Dim rsMaquinas As oraDynaset
Dim data As String
Dim Pre As String
Dim NumIp As String
Dim objMaquina As clsMaquina
Set objMaquina = New clsMaquina
Set rsMaquinas = objMaquina.Lista(objUsuario.CodigoEmpresa, vstrCod_Local)

Pre = "MENSDELI"
data = fn_Encriptar(Pre & "Tiene un pedido delivery pendiente " & Me.strNumProforma & " asignado a las " & objUsuario.sysdate)

With frm_DLV_Seguimiento.wskPrincipal
    Do Until rsMaquinas.EOF
        NumIp = IIf(IsNull(rsMaquinas!NUM_IP), "", rsMaquinas!NUM_IP)
        If fn_EsIPCorrecta(NumIp) = True Then
                host = rsMaquinas!NUM_IP
                .RemoteHost = host
                .RemotePort = 1025
                .SendData data
        End If
        rsMaquinas.MoveNext
    Loop
End With
Set rsMaquinas = Nothing
End Sub

Public Sub EvaluaLocalesTransf()
    On Error GoTo Err
    Dim InkaNeto As Boolean
    InkaNeto = False
    grdTransferencia.MoveFirst
    While Not grdTransferencia.EOF
        'Debug.Print grdTransferencia.Columns(8).Value
        If InkaNeto = False And objLocal.EvaluaLocalInkMig(grdTransferencia.Columns(8).Value) = "N" Then
            InkaNeto = True
        End If
        grdTransferencia.MoveNext
    Wend
    If InkaNeto = True Then
        ctlTextBox12.MaxLength = 100
        If ctlTextBox12.Text <> "" Then
            ctlTextBox12.Text = Mid(ctlTextBox12.Text, 1, 100)
        End If
    Else
        ctlTextBox12.MaxLength = 0
    End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub
