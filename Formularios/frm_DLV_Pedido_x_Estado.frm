VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_DLV_Pedido_x_Estado 
   Caption         =   "Pedidos por Estado"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13155
   Icon            =   "frm_DLV_Pedido_x_Estado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   22
      Top             =   120
      Width           =   13095
      Begin VB.CommandButton cmdRevertir 
         Caption         =   "&Revertir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   11520
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox ChkConexSalud 
         Caption         =   "Conexión Salud"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5640
         TabIndex        =   41
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CheckBox chkEstPendiente 
         Caption         =   "Pendientes"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkCallao 
         Caption         =   "Callao"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   3360
         Width           =   1095
      End
      Begin vbp_Ventas.ctlTextBox txtProforma 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Tipo            =   3
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
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   9960
         Picture         =   "frm_DLV_Pedido_x_Estado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58589185
         CurrentDate     =   39170
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58589185
         CurrentDate     =   39170
      End
      Begin vbp_Ventas.ctlDataCombo cboEstado 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox txtNombre 
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         Tipo            =   2
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
      Begin vbp_Ventas.ctlTextBox txtTelefono 
         Height          =   315
         Left            =   5880
         TabIndex        =   4
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Tipo            =   5
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
      Begin vbp_Ventas.ctlTextBox txtDireccion 
         Height          =   315
         Left            =   4920
         TabIndex        =   8
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         Tipo            =   2
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
      Begin vbp_Ventas.ctlDataCombo cboOrigen 
         Height          =   315
         Left            =   7200
         TabIndex        =   5
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboMotorizado 
         Height          =   315
         Left            =   5640
         TabIndex        =   16
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox TxtDocumento 
         Height          =   315
         Left            =   8760
         TabIndex        =   19
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         Tipo            =   3
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
      Begin vbp_Ventas.ctlDataCombo ctlCboTipoDoc 
         Height          =   315
         Left            =   8760
         TabIndex        =   18
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboRuteador 
         Height          =   315
         Left            =   5640
         TabIndex        =   17
         Top             =   3000
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboDistrito 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   2595
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboProvincia 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   2205
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboDepartamento 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboPais 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboUrbanizacion 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   3000
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlEmisorPed 
         Height          =   315
         Left            =   5640
         TabIndex        =   15
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo CboModalidad 
         Height          =   315
         Left            =   8760
         TabIndex        =   42
         Top             =   3000
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboEmpresa 
         Height          =   315
         Left            =   8280
         TabIndex        =   44
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   8280
         TabIndex        =   45
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modalidad"
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
         Left            =   8760
         TabIndex        =   43
         Top             =   2760
         Width           =   885
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   4920
         X2              =   4920
         Y1              =   1620
         Y2              =   3600
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Urbanización"
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
         Left            =   120
         TabIndex        =   40
         Top             =   3060
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito"
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
         Left            =   120
         TabIndex        =   39
         Top             =   2660
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Left            =   120
         TabIndex        =   38
         Top             =   2260
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
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
         Left            =   120
         TabIndex        =   37
         Top             =   1860
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pais"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1500
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruteador"
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
         Left            =   5640
         TabIndex        =   35
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emisor de Pedido"
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
         Left            =   5640
         TabIndex        =   34
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento"
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
         Left            =   8760
         TabIndex        =   33
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
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
         Left            =   8760
         TabIndex        =   32
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   12720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorizado"
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
         Left            =   5640
         TabIndex        =   31
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
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
         Left            =   7200
         TabIndex        =   30
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
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
         Left            =   4920
         TabIndex        =   29
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
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
         Left            =   5880
         TabIndex        =   28
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Left            =   2160
         TabIndex        =   27
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proforma"
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
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Fin"
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
         Left            =   4560
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Inicio"
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
         Left            =   3240
         TabIndex        =   24
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   600
      End
   End
   Begin vbp_Ventas.ctlGrilla grdPedido 
      Height          =   4215
      Left            =   0
      TabIndex        =   21
      Top             =   3840
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7435
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_DLV_Pedido_x_Estado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim objProforma As New clsProforma
Dim objDocumento As New clsDocumento
Dim strDepartamento As String
Dim strProvincia As String
Dim strDistrito As String
Dim strUbigeo As String
'I.ECASTILLO 21.10.2020
Dim strNumProforma As String
'F.ECASTILLO 21.10.2020

Private Sub cmdBuscar_Click()
On Error GoTo handle
    Dim strEstado As String
    Dim strOrigen As String
    Dim strMotorizado As String
    If cboEstado.BoundText = "000" Then strEstado = "" Else strEstado = cboEstado.BoundText
    If cboOrigen.BoundText = "AAA" Then strOrigen = "" Else strOrigen = cboOrigen.BoundText
    If cboMotorizado.BoundText = "" Then strMotorizado = "" Else strMotorizado = cboMotorizado.BoundText
    'if (txtProforma.Text = "") or (TxtDocumento.Text= "" and ctlCboTipoDoc.BoundText = "") then MsgBox ""
    ''cboEmpresa.BoundText,
    Set grdPedido.DataSource = objProforma.Lista_x_Estado(objUsuario.CodigoEmpresa, _
                                                          strEstado, _
                                                          objUsuario.CodigoLocal, _
                                                          dtpInicio.Value, _
                                                          dtpFin.Value + 1, _
                                                          Trim(txtTelefono.Text), _
                                                          Trim(txtProforma.Text), _
                                                          Trim(txtDireccion.Text), _
                                                          Trim(txtNombre.Text), _
                                                          strOrigen, _
                                                          strMotorizado, _
                                                          Trim(TxtDocumento.Text), _
                                                          ctlCboTipoDoc.BoundText, _
                                                          ctlEmisorPed.BoundText, cboPais.BoundText, cboDepartamento.BoundText, cboProvincia.BoundText, cboDistrito.BoundText, cboUrbanizacion.BoundText, cboRuteador.BoundText, chkEstPendiente.Value, ChkConexSalud.Value, CboModalidad.BoundText, cboEmpresa.BoundText)
    
    grdPedido.SetFocus
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
'I.ECASTILLO 21.10.2020
Private Sub cmdRevertir_Click()
On Error GoTo handle
    Dim resp As String
    If Len(Trim(strNumProforma)) = 0 Then
        MsgBox "Seleccione pedido a revertir", vbExclamation, App.ProductName
    ElseIf Len(objUsuario.CodigoEmpresa) <= 0 _
        Or Len(objUsuario.CodigoLocal) <= 0 _
        Or Len(objUsuario.Codigo) <= 0 _
    Then
        If gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGMSGRVPD") = 1 Then
            MsgBox "Algunas variables de usuario se encuentran vacias", vbExclamation, App.ProductName
        Else
            GoTo consulta
        End If
    Else
consulta:
        resp = objProforma.ReviertePedido(objUsuario.CodigoEmpresa, _
                                          objUsuario.CodigoLocal, _
                                          strNumProforma, _
                                          objUsuario.Codigo, _
                                          "1" _
                                          )
        If Len(Trim(resp)) > 0 Then
            MsgBox resp, vbCritical, App.ProductName
        Else
            'grdPedido.DataSource.Refresh
            cmdBuscar_Click
        End If
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
'F.ECASTILLO 21.10.2020

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo handle
    Set cboEstado.RowSource = objProforma.ListaEstado("")
    cboEstado.ListField = "DES_ESTADO_PEDIDO"
    cboEstado.BoundColumn = "COD_ESTADO_PEDIDO"
    cboEstado.BoundText = "000"
    
    Set ctlCboTipoDoc.RowSource = objDocumento.ListaTipo(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
    ctlCboTipoDoc.ListField = "DESCRIPCION"
    ctlCboTipoDoc.BoundColumn = "CODIGO"
    ctlCboTipoDoc.BoundText = "*"
    
    Set ctlEmisorPed.RowSource = objUsuario.ListaUsuarioDLV
    ctlEmisorPed.ListField = "DES"
    ctlEmisorPed.BoundColumn = "COD"
    ctlEmisorPed.BoundText = "*"
    
    Set cboRuteador.RowSource = objUsuario.ListaUsuarioDLVRuta
    cboRuteador.ListField = "DES"
    cboRuteador.BoundColumn = "COD"
    cboRuteador.BoundText = "*"
    
    
    Set cboEmpresa.RowSource = objUsuario.ListaEmpresa("")
    cboEmpresa.ListField = "DESCRIP"
    cboEmpresa.BoundColumn = "CIA"
    cboEmpresa.BoundText = "*"
    
    dtpInicio.Value = objUsuario.sysdate
    dtpFin.Value = objUsuario.sysdate
    Dim objLocal As New clsLocal
    
    Set cboOrigen.RowSource = objLocal.Lista("", "", "002")
    'Set cboOrigen.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
    'Set cboOrigen.RowSource = objLocal.Lista_Inc_Todos(objUsuario.CodigoEmpresa)
    cboOrigen.BoundColumn = "COD_LOCAL"
    cboOrigen.ListField = "local_dex2"
    cboOrigen.BoundText = "AAA"
    
    Set objLocal = Nothing
        Set cboPais.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PAIS", 0)
    cboPais.ListField = "Descripcion"
    cboPais.BoundColumn = "Codigo"
    cboPais.BoundText = "00"


    Set cboDepartamento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0, "[SELECCIONAR]")
    cboDepartamento.ListField = "Descripcion"
    cboDepartamento.BoundColumn = "Codigo"
    cboDepartamento.BoundText = "*"
    
    Dim objMotorizado As New clsMotorizado
    'Set cboMotorizado.RowSource = objMotorizado.Lista("", "")
    Set cboMotorizado.RowSource = objMotorizado.Lista_Inc_Todos("")
    cboMotorizado.BoundColumn = "COD_MOTORIZADO"
    cboMotorizado.ListField = "NOMBRE"
    cboMotorizado.BoundText = ""
    cboOrigen.Text = "[TODOS]": cboMotorizado.Text = "TODOS"
    Set objMotorizado = Nothing
    SetGrd
    
    'Agrega filtro de modalidad
    
    Set CboModalidad.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_MODALIDAD.FN_LISTA_TODOS", 0)
    CboModalidad.ListField = "DES_MODALIDAD_VENTA"
    CboModalidad.BoundColumn = "COD_MODALIDAD_VENTA"
    CboModalidad.BoundText = "*"
    
    'I.ECASTILLO 21.10.2020
    'Identificar si el usuario tiene rol/perfil coordinador, si lo es habilita btnRevertir
    Dim strEncontro As String
    'Tomando en cuenta que las coordinadoras tienen perfil operador,
    'al parecer solo coordinadoras asignan meta
    strEncontro = Trim(objUsuario.AsignaMetaDLV(objUsuario.CodigoAplicacion, _
                                                objUsuario.CodigoMenuAsigna, _
                                                objUsuario.Codigo))
    If strEncontro = "1" Then
        cmdRevertir.Visible = True: cmdRevertir.Enabled = True
    Else
        cmdRevertir.Enabled = False: cmdRevertir.Visible = False
    End If
    'F.ECASTILLO 21.10.2020
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo handle
    Set objProforma = Nothing
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdPedido_DblClick()
On Error GoTo handle
    Dim frm As New frm_VTA_DetallePedido
    If grdPedido.ApproxCount = 0 Then Exit Sub
    'frm_VTA_DetallePedido.Numeropedido = grdPedido.DataSource("NUM_PROFORMA")
    'frm_VTA_DetallePedido.CodigoLocal = grdPedido.DataSource("COD_LOCAL_REF")
    'frm_VTA_DetallePedido.Show
    frm.NumeroPedido = grdPedido.DataSource("NUM_PROFORMA")
    frm.CodigoLocal = grdPedido.DataSource("COD_LOCAL_REF")
    frm.ReCargaDetPedido
    frm.Show
    Set frm = Nothing
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub SetGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
            arrCampos = Array("NUM_PROFORMA", "NUM_PEDIDO_REF", "DES_ESTADO_PEDIDO", _
                              "FCH_REGISTRA", "DES_AUX_CLI_TLF", _
                              "DES_AUX_CLI_NOMBRE", "DES_AUX_CLI_DIRECC", _
                               "CIA_REF", "COD_LOCAL_SAP", "NOM_USUARIO")
                              
            arrCaption = Array("Pedido", "Pedido Ref.", "Estado", _
                               "Fecha", "Telefono", _
                               "Nombre Cliente", "Dirección", _
                                "Empresa", "Local", "Emisor Pedido")
                               
            arrAncho = Array(1200, 1200, 1500, _
                             1000, 1200, _
                             2500, 3500, _
                             700, 600, 3200)
                             
            arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, _
                                  dbgCenter, dbgCenter, _
                                  dbgLeft, dbgLeft, _
                                  dbgCenter, dbgLeft, dbgLeft)
                                  
            grdPedido.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub cboProvincia_Change()
    Set cboDistrito.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, cboDepartamento.BoundText, cboProvincia.BoundText, "[SELECCIONAR]")
    cboDistrito.ListField = "Descripcion"
    cboDistrito.BoundColumn = "Codigo"
    cboDistrito.BoundText = "*"
End Sub


Private Sub cboDepartamento_Change()
    Set cboProvincia.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, cboDepartamento.BoundText, "[SELECCIONAR]")
    cboProvincia.ListField = "Descripcion"
    cboProvincia.BoundColumn = "Codigo"
    cboProvincia.BoundText = "*"
End Sub


Private Sub cboDistrito_Change()
    strUbigeo = cboDepartamento.BoundText & cboProvincia.BoundText & cboDistrito.BoundText
    Set cboUrbanizacion.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_URBANIZACION", 0, "2", strUbigeo)
    cboUrbanizacion.ListField = "DES_URBANIZACION"
    cboUrbanizacion.BoundColumn = "COD_URBANIZACION"
End Sub

Private Sub chkCallao_Click()
If chkCallao.Value = "1" Then
    cboDepartamento.BoundText = "07"
    cboProvincia.BoundText = "01"
Else
    On Error GoTo Y
       cboDepartamento.BoundText = Mid(objUsuario.UbigeoLocal, 1, 2)
       cboProvincia.BoundText = Mid(objUsuario.UbigeoLocal, 3, 2)
       cboDistrito.BoundText = Mid(objUsuario.UbigeoLocal, 5, 2)
End If
cboDistrito.SetFocus
Exit Sub
Y:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub grdPedido_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo Control
    If grdPedido.ApproxCount > 0 Then
        strNumProforma = grdPedido.Columns("NUM_PROFORMA").Value
    Else
        strNumProforma = ""
    End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub
