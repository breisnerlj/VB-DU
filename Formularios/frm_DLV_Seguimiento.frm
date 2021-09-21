VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_DLV_Seguimiento 
   BackColor       =   &H80000013&
   Caption         =   "Seguimiento de Pedidos"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   18285
   Icon            =   "frm_DLV_Seguimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   18285
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option3 
      Caption         =   "Anulados"
      Height          =   435
      Left            =   4560
      TabIndex        =   45
      Top             =   0
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   375
      Left            =   13920
      TabIndex        =   38
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57999361
      CurrentDate     =   40170
   End
   Begin VB.CommandButton cmdReclamo 
      Caption         =   "Reclamo"
      Height          =   650
      Left            =   10080
      Picture         =   "frm_DLV_Seguimiento.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdCambioLocalDesp 
      Caption         =   "Local Entrega"
      Height          =   650
      Left            =   11365
      Picture         =   "frm_DLV_Seguimiento.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdPedido 
      Caption         =   "Detalle Pedido"
      Height          =   650
      Left            =   12650
      Picture         =   "frm_DLV_Seguimiento.frx":0E1E
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   650
      Left            =   13935
      Picture         =   "frm_DLV_Seguimiento.frx":13A8
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Asignar"
      Height          =   650
      Left            =   1337
      Picture         =   "frm_DLV_Seguimiento.frx":1932
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdLiberar 
      Caption         =   "Liberar"
      Height          =   650
      Left            =   7440
      Picture         =   "frm_DLV_Seguimiento.frx":1EBC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   650
      Left            =   8640
      Picture         =   "frm_DLV_Seguimiento.frx":2446
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdAvisado 
      Caption         =   "&Avisado"
      Height          =   650
      Left            =   2520
      Picture         =   "frm_DLV_Seguimiento.frx":29D0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdVerificado 
      Caption         =   "Ret.Verif."
      Height          =   615
      Left            =   120
      Picture         =   "frm_DLV_Seguimiento.frx":2F5A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9555
      Width           =   1200
   End
   Begin VB.CommandButton cmdLlevando 
      Caption         =   "Llevando"
      Height          =   650
      Left            =   2554
      Picture         =   "frm_DLV_Seguimiento.frx":34E4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdLlegadaDestino 
      Caption         =   "Llega Destino"
      Height          =   650
      Left            =   3771
      Picture         =   "frm_DLV_Seguimiento.frx":3A6E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdEntregado 
      Caption         =   "Entregado"
      Height          =   650
      Left            =   4988
      Picture         =   "frm_DLV_Seguimiento.frx":3FF8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9540
      Width           =   1200
   End
   Begin VB.CommandButton cmdLlegadaLocal 
      Caption         =   "Llegada a local"
      Height          =   650
      Left            =   6205
      Picture         =   "frm_DLV_Seguimiento.frx":4582
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9540
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiempos"
      Height          =   915
      Left            =   0
      TabIndex        =   6
      Top             =   7980
      Width           =   18105
      Begin vbp_Ventas.ctlGrilla grdTiempo 
         Height          =   615
         Left            =   30
         TabIndex        =   7
         Top             =   240
         Width           =   18015
         _ExtentX        =   31776
         _ExtentY        =   1085
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Asignados"
      Height          =   435
      Left            =   5760
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000004&
      Caption         =   "&No Asignados"
      Height          =   435
      Left            =   6960
      TabIndex        =   1
      Top             =   0
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4530
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":4B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":4F37
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":536D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":5755
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":5B2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2610
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":5F11
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":65D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":6CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":71F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Seguimiento.frx":76F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin vbp_Ventas.ctlDataCombo cboZona 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   90
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlGrilla grdPedidos 
      Height          =   7485
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   13203
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   495
      Left            =   45
      TabIndex        =   8
      Top             =   8940
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   873
      ColorDefault    =   -2147483639
      ColorDefault    =   -2147483639
      Enabled         =   0   'False
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
   Begin MSComCtl2.DTPicker dtpFin 
      Height          =   375
      Left            =   15600
      TabIndex        =   39
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57999361
      CurrentDate     =   40170
   End
   Begin VB.Label Label19 
      Caption         =   "=INKAFARMA"
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
      Left            =   11895
      TabIndex        =   44
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   11640
      Top             =   30
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000008&
      Height          =   135
      Left            =   10440
      Top             =   30
      Width           =   255
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "=MIFARMA"
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
      Left            =   10695
      TabIndex        =   43
      Top             =   0
      Width           =   960
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   9240
      Top             =   30
      Width           =   255
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "=BTL"
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
      Left            =   9495
      TabIndex        =   42
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   255
      Left            =   13320
      TabIndex        =   41
      Top             =   60
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      Height          =   255
      Left            =   15360
      TabIndex        =   40
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Alt + R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   10260
      TabIndex        =   37
      Top             =   10200
      Width           =   840
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   4
      Left            =   11800
      TabIndex        =   35
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "=Urg.+Transf."
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
      Left            =   11895
      TabIndex        =   33
      Top             =   255
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "=Transf."
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
      Left            =   10695
      TabIndex        =   32
      Top             =   255
      Width           =   720
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   11640
      Top             =   285
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   10440
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "=Urgente"
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
      Left            =   9495
      TabIndex        =   31
      Top             =   255
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   9240
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   3
      Left            =   13010
      TabIndex        =   30
      Top             =   10200
      Width           =   480
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   2
      Left            =   14295
      TabIndex        =   28
      Top             =   10200
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   555
      TabIndex        =   26
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   1772
      TabIndex        =   25
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3780
      TabIndex        =   24
      Top             =   10560
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2989
      TabIndex        =   23
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4206
      TabIndex        =   22
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   5423
      TabIndex        =   21
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   6640
      TabIndex        =   20
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   0
      Left            =   7857
      TabIndex        =   19
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   1
      Left            =   9075
      TabIndex        =   18
      Top             =   10200
      Width           =   330
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Zona:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mostar los pedidos"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1320
   End
   Begin VB.Menu mnuMantenimientos 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnuMotorizados 
         Caption         =   "&Motorizados"
      End
      Begin VB.Menu mnuZonas 
         Caption         =   "&Zonas"
      End
      Begin VB.Menu mnuasiglocal 
         Caption         =   "&Asinación Locales"
      End
   End
End
Attribute VB_Name = "frm_DLV_Seguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPedido As New clsProforma
Dim objLocal As New clsLocal
Dim objWS As New clsWebService
Dim objProducto As New clsProducto
Public WithEvents wskPrincipal As CSocketMaster
Attribute wskPrincipal.VB_VarHelpID = -1

Private Sub cboZona_Change()
    ListaPedido
End Sub

Private Sub CmdActualizar_Click()
    On Error GoTo Control
    ListaPedido

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error :" & Err.Number
End Sub

Private Sub CmdAnular_Click()
Dim strMensaje As String
'I.ECASTILLO 27.10.2020
Dim respCancelaDigital As String
Dim respRevierteReserva As String
'F.ECASTILLO 27.10.2020
Dim Bookmark As Variant
On Error GoTo Control

    If Not cmdAnular.Enabled Then Exit Sub

    If grdPedidos.ApproxCount = 0 Then Exit Sub
    If MsgBox("Se va anular la proforma " & grdPedidos.Columns("NUM_PROFORMA") & " , desea continuar ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        'I.ECASTILLO 24.07.2020
        If grdPedidos.Columns("FLG_TRANSFERENCIA") = 0 Then
            If objPedido.AnulaReservaStock(grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objUsuario.Codigo) = False Then
                GoTo ErrorReserva
            End If
            'I.ECASTILLO 27.10.2020
            ' 17.12.2020 | se comenta pero volver a usar
            respRevierteReserva = objPedido.revierteReservaStock("10", grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), "2", "1")

            respCancelaDigital = objPedido.cancelaDigital("10", grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"))
            If respCancelaDigital <> "1" Then 'no cancelo
                If respCancelaDigital <> "0" Then 'hubo algun error
                    MsgBox "Error al cancelar pedido en Digital" & vbNewLine & respCancelaDigital, vbOKOnly + vbCritical, App.ProductName: Exit Sub
                End If
                GoTo anulaBD 'este pedido debe ser anulado regularmente
            Else 'si cancelo
                GoTo continuar
            End If
            'F.ECASTILLO 27.10.2020
        End If
        'F.ECASTILLO 24.07.2020
anulaBD:
        strMensaje = objPedido.Anula(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objUsuario.Codigo)
        If strMensaje <> "" Then
            MsgBox strMensaje, vbCritical, App.ProductName
            Bookmark = grdPedidos.Bookmark
            grdPedidos.SetFocus
        End If
'I.ECASTILLO 27.10.2020
continuar:
'F.ECASTILLO 27.10.2020
        Bookmark = grdPedidos.Bookmark
        CmdActualizar_Click
        grdPedidos.Bookmark = Bookmark
    End If
    Exit Sub
ErrorReserva:
    'MsgBox "Ocurrio un error al anular reserva", vbOKOnly + vbCritical, "Error: anular reserva"
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub








'Private Sub cmdAusencia_Click()
'Dim msgbo As Variant
'    msgbo = MsgBox("¿El Motorizado se encuentra ausente?", vbYesNoCancel, App.ProductName)
'    If msgbo = vbCancel Then Exit Sub
'Dim objMotorizado As New clsMotorizado
'Dim strMensaje As String
'    strMensaje = objMotorizado.GrabaAusencia(grdAsistenciaMotorizado.Columns(0).Value, IIf(msgbo = vbYes, True, False))
'    If strMensaje = "" Then
'        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
'    Else
'        MsgBox strMensaje, vbCritical, App.ProductName
'    End If
'    cmdBuscar_Click
'    Set objMotorizado = Nothing
'End Sub

Private Sub cmdAvisado_Click()
    Dim strMensaje As String
    Dim Bookmark As Variant
    Dim k, i As Integer
    
On Error GoTo Control

    If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoAsignado Then Exit Sub
   ' If grdPedidos.SelBookmarks.Count = 0 Then Exit Sub
    
    i = grdPedidos.SelBookmarks.Count - 1
    For k = i To 0 Step -1
        grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoAvisado, objUsuario.Codigo)
    Next
    If grdPedidos.SelBookmarks.Count = 0 Then
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoAvisado, objUsuario.Codigo)
    End If
    If strMensaje <> "" Then
'        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
'    Else
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    Bookmark = grdPedidos.Bookmark
    CmdActualizar_Click
    grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdCambioLocalDesp_Click()
On Error GoTo Control

    If Not cmdCambioLocalDesp.Enabled Then Exit Sub

    If grdPedidos.ApproxCount <= 0 Then Exit Sub
    'If cboZona.Text = "" Then Exit Sub
    If grdPedidos.Columns("DES_ESTADO").Value <> "Ver" Then MsgBox "El estado del pedido tiene que ser Verificado", vbCritical, Caption: Exit Sub
    frm_DLV_CambioLocalDespacho.pZonax = Trim(cboZona.BoundText) & "  " & Trim(cboZona.Text)
    frm_DLV_CambioLocalDespacho.pLocalx = grdPedidos.Columns("COD_LOCAL").Value
    frm_DLV_CambioLocalDespacho.pPedidox = grdPedidos.Columns("NUM_PROFORMA").Value
    frm_DLV_CambioLocalDespacho.pCiaRefx = grdPedidos.Columns("COD_CIA_REF").Value
    frm_DLV_CambioLocalDespacho.pLocalRefx = grdPedidos.Columns("COD_LOCAL_REF").Value
    frm_DLV_CambioLocalDespacho.pLocalSapfx = grdPedidos.Columns("COD_LOCAL_SAP_REF").Value
    frm_DLV_CambioLocalDespacho.Show vbModal
    
    CmdActualizar_Click
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

'Private Sub cmdBuscar_Click()
'    Dim objMotorizado As New clsMotorizado
'    Set grdAsistenciaMotorizado.DataSource = objMotorizado.ListaDisponible(cboLocalAsig2.BoundText, DTPicker1.Value)
'    Set objMotorizado = Nothing
'End Sub

Private Sub cmdDetalle_Click()

    If Not cmdDetalle.Enabled Then Exit Sub
    If grdPedidos.ApproxCount = 0 Then Exit Sub
    'If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") = objPedido.PedidoAsignado Or Option1.Value = True Then Exit Sub
    
    On Error GoTo handle
    
    frm_DLV_Ruteo.bolModoConsulta = False
    If grdPedidos.Columns("COD_ESTADO") >= objPedido.PedidoLlevando And grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoLiberado Then
        'MsgBox "No se puede asignar el pedido, primero debe Liberarlo", vbOKOnly + vbExclamation, App.ProductName
        'Exit Sub
        frm_DLV_Ruteo.bolModoConsulta = True
    End If
    
    frm_DLV_Ruteo.strLocal = grdPedidos.Columns("COD_LOCAL")
    frm_DLV_Ruteo.strLocalPedido = grdPedidos.Columns("COD_LOCAL_REF") ' local pedido
    frm_DLV_Ruteo.strNumProforma = grdPedidos.Columns("NUM_PROFORMA")
    frm_DLV_Ruteo.strNumPedRef = grdPedidos.Columns("NUM_PEDIDO_REF")
    frm_DLV_Ruteo.strFlgTransf = grdPedidos.Columns("FLG_TRANSFERENCIA")
    frm_DLV_Ruteo.bolEsLlamadoCab = True
    If grdPedidos.Columns("COD_ESTADO") = objPedido.PedidoAsignado Or Option1.Value = True Then
        Call CambioTipoTransferencia
    End If
    If Option2.Value = True Then
        If grdPedidos.Columns("FLG_TRANSFERENCIA") = 1 Then
            Call CambioTipoTransferencia
        End If
    End If
    frm_DLV_Ruteo.Show vbModal
    Exit Sub
    
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    
End Sub

Sub CambioTipoTransferencia()
    frm_DLV_Ruteo.cboRuta1.Enabled = False
    frm_DLV_Ruteo.cboLocalAsig.Enabled = False
    frm_DLV_Ruteo.cmdTranferencia.Enabled = False
    frm_DLV_Ruteo.Check1.Enabled = False
    'frm_DLV_Ruteo.chkLocConStock.Enabled = False
    frm_DLV_Ruteo.grdPedidoStock.Enabled = False
    
    frm_DLV_Ruteo.cboRuta1.Visible = False
    frm_DLV_Ruteo.cboLocalAsig.Visible = False
    frm_DLV_Ruteo.cmdTranferencia.Visible = False
    frm_DLV_Ruteo.Check1.Visible = False
    'frm_DLV_Ruteo.chkLocConStock.Visible = False
    
    '''Label12.Visible = False
    '''Label13.Visible = False
End Sub

Private Sub cmdEntregado_Click()
    Dim strMensaje As String
    Dim Bookmark  As Variant
    Dim k, i As Integer
    Dim objDelivery As clsDelivery
On Error GoTo Control

    If Not cmdEntregado.Enabled Then Exit Sub
    
    
    '********************************************************************'
    '** Validando para cambiar al estado al motorizado según la acción **'
    '**             08/02/2008 Por Cristhian Rueda                     **'
    '********************************************************************'
    
'''''     objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"))
'''''
'''''     If objUsuario.CodigoMotorizado <> "" Then
'''''        Dim vstrMensaje As String
'''''        vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
'''''                                                       objPedido.PedidoEntregado, _
'''''                                                       objUsuario.Codigo)
'''''        If vstrMensaje = "" Then
'''''         Else
'''''            MsgBox vstrMensaje, vbCritical, Caption
'''''            Exit Sub
'''''        End If
'''''     End If
'''''    '********************************************************************'
'''''
'''''
'''''    If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoLlegada Then Exit Sub
'''''    'If grdPedidos.SelBookmarks.Count = 0 Then Exit Sub
    
    Set objDelivery = New clsDelivery
    
    Bookmark = grdPedidos.Bookmark
    
    
    i = grdPedidos.SelBookmarks.Count - 1
    
    If i > 0 Then
        For k = i To 0 Step -1
            grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
            If objDelivery.ValidaEntregado(grdPedidos.Columns("COD_LOCAL"), _
                        grdPedidos.Columns("NUM_PROFORMA"), _
                        grdPedidos.Columns("COD_ESTADO"), _
                        grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoEntregado, objUsuario.Codigo)
            If strMensaje <> "" Then Exit For
        Next
    ElseIf grdPedidos.ApproxCount > 0 Then
            If objDelivery.ValidaEntregado(grdPedidos.Columns("COD_LOCAL"), _
                        grdPedidos.Columns("NUM_PROFORMA"), _
                        grdPedidos.Columns("COD_ESTADO"), _
                        grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoEntregado, objUsuario.Codigo)
    End If
       
    Set objDelivery = Nothing
        
    If strMensaje <> "" Then MsgBox strMensaje, vbCritical, App.ProductName
    
    CmdActualizar_Click
    grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

'Private Sub cmdEstado_Click(Index As Integer)
'Dim strCodigoMotorizado As String
'strCodigoMotorizado = "" & grdAsistenciaMotorizado.Columns(0).Value
'If Index = 1 Then strCodigoMotorizado = cboMotorizado.BoundText
'Dim objMorotizado As New clsMotorizado
'Dim Mensaje As String
'Mensaje = objMorotizado.GrabaEstado(strCodigoMotorizado, Index, objUsuario.Codigo)
'If Mensaje = "" Then
'    MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
'Else
'    MsgBox Mensaje, vbCritical, App.ProductName
'End If
'cmdBuscar_Click
'Set objMorotizado = Nothing
'End Sub

Private Sub Command2_Click()
 '   MsgBox "CAMBIAR DE ESTADO, MAY TRAMPO"
End Sub

Private Sub cmdLiberar_Click()
    Dim strMensaje As String
    Dim respRevierteReserva As String 'ECASTILLO 27.10.2020
    Dim Bookmark As Variant
    Dim k, i As Integer
    Dim objDelivery As clsDelivery
    
    If Not cmdLiberar.Enabled Then Exit Sub
    
    If grdPedidos.ApproxCount = 0 Then Exit Sub
    'If grdPedidos.SelBookmarks.Count = 0 Then Exit Sub
    
    
    '********************************************************************'
    '** Validando para cambiar al estado al motorizado según la acción **'
    '**             08/02/2008 Por Cristhian Rueda                     **'
    '********************************************************************'
    
'''''     objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"))
'''''
'''''     If objUsuario.CodigoMotorizado <> "" Then
'''''        Dim vstrMensaje As String
'''''
'''''        vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
'''''                                                       objPedido.PedidoLiberado, _
'''''                                                       objUsuario.Codigo)
'''''        If vstrMensaje = "" Then
'''''           ' MsgBox "Se Grabo la remesa sastifactoriamente", vbInformation, Caption
'''''        Else
'''''            MsgBox vstrMensaje, vbCritical, Caption
'''''        End If
'''''     End If
    '********************************************************************'
    
    Set objDelivery = New clsDelivery
    
    Bookmark = grdPedidos.Bookmark
    
    i = grdPedidos.SelBookmarks.Count - 1
    If i > 0 Then
        For k = i To 0 Step -1
            grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
            If objDelivery.ValidaLiberar(grdPedidos.Columns("COD_LOCAL"), _
                        grdPedidos.Columns("NUM_PROFORMA"), _
                        grdPedidos.Columns("COD_ESTADO"), _
                        grdPedidos.ApproxCount) Then
                        
                'I.ECASTILLO 24.07.2020
                If grdPedidos.Columns("FLG_TRANSFERENCIA") = 0 Then
                    If objPedido.AnulaReservaStock(grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objUsuario.Codigo) = False Then
                        GoTo ErrorReserva
                    End If
                    'I.ECASTILLO 27.10.2020
                    ' 17.12.2020 | se comenta pero volver a usar
                    respRevierteReserva = objPedido.revierteReservaStock("10", grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), "0", "1")
                    'F.ECASTILLO 27.10.2020
                End If
                'F.ECASTILLO 24.07.2020
                strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLiberado, objUsuario.Codigo)
                If strMensaje = "" Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoVerificado, objUsuario.Codigo)
                If strMensaje = "" Then strMensaje = objPedido.ActualizaMotorizado_Proforma(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), "")
                If strMensaje <> "" Then Exit For
            End If
        Next
    ElseIf grdPedidos.ApproxCount > 0 Then
        'I.ECASTILLO 24.07.2020
        If grdPedidos.Columns("FLG_TRANSFERENCIA") = 0 Then
            If objPedido.AnulaReservaStock(grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objUsuario.Codigo) = False Then
                GoTo ErrorReserva
            End If
            'I.ECASTILLO 27.10.2020
            ' 17.12.2020 | se comenta pero volver a usar
            respRevierteReserva = objPedido.revierteReservaStock("10", grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), "0", "1")
            'F.ECASTILLO 27.10.2020
        End If
        'F.ECASTILLO 24.07.2020
                
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLiberado, objUsuario.Codigo)
        If strMensaje = "" Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoVerificado, objUsuario.Codigo)
        If strMensaje = "" Then strMensaje = objPedido.ActualizaMotorizado_Proforma(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), "")
    End If
    
    Set objDelivery = Nothing
    
    If strMensaje <> "" Then MsgBox strMensaje, vbCritical, App.ProductName
    CmdActualizar_Click
    grdPedidos.Bookmark = Bookmark
    Exit Sub
ErrorReserva:
    'MsgBox "Ocurrio un error al anular reserva", vbOKOnly + vbCritical, "Error: anular reserva"
    Exit Sub
    
'''''    If strMensaje = "" Then
'''''        GoTo ACA
'''''    Else
'''''        MsgBox strMensaje, vbCritical, App.ProductName
'''''        GoTo aqui
'''''    End If
'''''
'''''    'PASAR A VERIFICADO
'''''ACA:
'''''    i = grdPedidos.SelBookmarks.Count - 1
'''''    For k = i To 0 Step -1
'''''        grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
'''''        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoVerificado, objUsuario.Codigo)
'''''    Next
'''''    If grdPedidos.SelBookmarks.Count = 0 Then
'''''        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoVerificado, objUsuario.Codigo)
'''''    End If
'''''    If strMensaje = "" Then
'''''        strMensaje = objPedido.ActualizaMotorizado_Proforma(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), "")
'''''        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
'''''    Else
'''''        MsgBox strMensaje, vbCritical, App.ProductName
'''''        GoTo aqui
'''''    End If
'''''aqui:
'''''    Bookmark = grdPedidos.Bookmark
'''''    CmdActualizar_Click
'''''    grdPedidos.Bookmark = Bookmark
End Sub

Private Sub cmdLlegadaDestino_Click()
    Dim strMensaje As String
    Dim Bookmark As Variant
    Dim k, i As Integer
    Dim objDelivery As clsDelivery
    
On Error GoTo Control

    If Not cmdLlegadaDestino.Enabled Then Exit Sub
'''''    If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoLlevando Then Exit Sub
    Set objDelivery = New clsDelivery
    Bookmark = grdPedidos.Bookmark
    i = grdPedidos.SelBookmarks.Count - 1
    If i > 0 Then
        For k = i To 0 Step -1
            grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
            If objDelivery.ValidaLlegadaDestino(grdPedidos.Columns("COD_ESTADO"), _
                            grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegada, objUsuario.Codigo)
            If strMensaje <> "" Then Exit For
        Next
    ElseIf grdPedidos.ApproxCount > 0 Then
        If objDelivery.ValidaLlegadaDestino(grdPedidos.Columns("COD_ESTADO"), _
                            grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegada, objUsuario.Codigo)
    End If
    If strMensaje <> "" Then MsgBox strMensaje, vbCritical, App.ProductName
    Set objDelivery = Nothing
    CmdActualizar_Click
    grdPedidos.Bookmark = Bookmark
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdLlegadaLocal_Click()
    Dim Bookmark As Variant
    Dim strMensaje As String
    Dim k, i As Integer
    Dim objDelivery As clsDelivery
    
On Error GoTo Control


    If Not cmdLlegadaLocal.Enabled Then Exit Sub
    
    '********************************************************************'
    '** Validando para cambiar al estado al motorizado según la acción **'
    '**             08/02/2008 Por Cristhian Rueda                     **'
    '********************************************************************'
    
'''''     objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"))
'''''
'''''     If objUsuario.CodigoMotorizado <> "" Then
'''''        Dim vstrMensaje As String
'''''        vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
'''''                                                       objPedido.PedidoLlegadaLocal, _
'''''                                                       objUsuario.Codigo)
'''''        If vstrMensaje = "" Then
'''''         Else
'''''            MsgBox vstrMensaje, vbCritical, Caption
'''''        End If
'''''     End If
'''''    '********************************************************************'
'''''
'''''    If strMensaje <> "" Then
''''''        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
''''''    Else
'''''        MsgBox strMensaje, vbCritical, App.ProductName
'''''    End If
'''''
'''''
'''''    If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoEntregado Then Exit Sub
'''''    'If grdPedidos.SelBookmarks.Count = 0 Then Exit Sub
    
    Set objDelivery = New clsDelivery
    
    Bookmark = grdPedidos.Bookmark
    
    i = grdPedidos.SelBookmarks.Count - 1
    
    If i > 0 Then
        For k = i To 0 Step -1
            grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
            If objDelivery.ValidaLlegadaLocal(grdPedidos.Columns("COD_LOCAL"), _
                        grdPedidos.Columns("NUM_PROFORMA"), _
                        grdPedidos.Columns("COD_ESTADO"), _
                        grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegadaLocal, objUsuario.Codigo)
            If strMensaje <> "" Then Exit For
        Next
    ElseIf grdPedidos.ApproxCount > 0 Then
        If objDelivery.ValidaLlegadaLocal(grdPedidos.Columns("COD_LOCAL"), _
                        grdPedidos.Columns("NUM_PROFORMA"), _
                        grdPedidos.Columns("COD_ESTADO"), _
                        grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegadaLocal, objUsuario.Codigo)
    End If
    Set objDelivery = Nothing
    If strMensaje <> "" Then MsgBox strMensaje, vbCritical, App.ProductName
    CmdActualizar_Click
    grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdLlevando_Click()
Dim strMensaje As String
Dim Bookmark As Variant
Dim k, i As Integer
Dim objDelivery As clsDelivery

On Error GoTo Control

    If Not cmdLlevando.Enabled Then Exit Sub
    
    '********************************************************************'
    '** Validando para cambiar al estado al motorizado según la acción **'
    '**             08/02/2008 Por Cristhian Rueda                     **'
    '********************************************************************'
    
''''''     objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"))
''''''
''''''
''''''     If grdPedidos.Columns("COD_ESTADO") = objPedido.PedidoAvisado And objUsuario.CodigoMotorizado = "" Then
''''''        MsgBox "Se tiene que asignar un Motorizado", vbOKOnly + vbExclamation, App.ProductName
''''''        Exit Sub
''''''     End If
''''''
''''''
''''''     If objUsuario.CodigoMotorizado <> "" Then
''''''        Dim vstrMensaje As String
''''''
''''''        vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
''''''                                                       objPedido.PedidoLlevando, _
''''''                                                       objUsuario.Codigo)
''''''        If vstrMensaje = "" Then
''''''           ' MsgBox "Se Grabo la remesa sastifactoriamente", vbInformation, Caption
''''''        Else
''''''            MsgBox vstrMensaje, vbCritical, Caption
''''''        End If
''''''     End If
''''''    '********************************************************************'
''''''
''''''
''''''    If grdPedidos.Columns("COD_ESTADO") = objPedido.PedidoAvisado Then
''''''        If MsgBox("Desea cambiar a Llevando sin necesidad que el pedido este Proformado ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirme") = vbNo Then
''''''            If (grdPedidos.ApproxCount = 0) Or (grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoProforma) Then Exit Sub
''''''        End If
''''''    ElseIf grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoProforma Then
''''''        MsgBox "No se puede cambiar el estado al pedido, debe estar en proforma (PRO) o avisado (AVI)", vbOKOnly + vbExclamation, "Error"
''''''        Exit Sub
''''''    End If
    
    
    Set objDelivery = New clsDelivery
    
    Bookmark = grdPedidos.Bookmark
    
    i = grdPedidos.SelBookmarks.Count - 1
    If i > 0 Then
        For k = i To 0 Step -1
            grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
            If objDelivery.ValidaLlevando(grdPedidos.Columns("COD_LOCAL"), _
                        grdPedidos.Columns("NUM_PROFORMA"), _
                        grdPedidos.Columns("COD_ESTADO"), _
                        grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlevando, objUsuario.Codigo)
            If strMensaje <> "" Then Exit For
        Next
    ElseIf grdPedidos.ApproxCount > 0 Then
        If objDelivery.ValidaLlevando(grdPedidos.Columns("COD_LOCAL"), _
                        grdPedidos.Columns("NUM_PROFORMA"), _
                        grdPedidos.Columns("COD_ESTADO"), _
                        grdPedidos.ApproxCount) Then strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlevando, objUsuario.Codigo)
    End If
    
    
    Set objDelivery = Nothing
    If strMensaje <> "" Then MsgBox strMensaje, vbCritical, App.ProductName
    
   CmdActualizar_Click
   grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdPedido_Click()
On Error GoTo CtrlErr

    If Not cmdPedido.Enabled Then Exit Sub
    If grdPedidos.ApproxCount = 0 Then Exit Sub
    frm_VTA_DetallePedido.NumeroPedido = grdPedidos.Columns("NUM_PROFORMA").Value
    frm_VTA_DetallePedido.CodigoLocal = grdPedidos.Columns("COD_LOCAL_REF")
    frm_VTA_DetallePedido.ReCargaDetPedido
    frm_VTA_DetallePedido.Show vbModal
Exit Sub
CtrlErr:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub cmdReclamo_Click()
    
    On Error GoTo handle
    
    If Not cmdReclamo.Enabled Then Exit Sub
    frm_DLV_Pedido_x_Estado.Show
    
    Exit Sub
    
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdVerificado_Click()

On Error GoTo handle
    If Not cmdVerificado.Enabled Then Exit Sub

    If grdPedidos.ApproxCount = 0 Then Exit Sub
    
    Dim strMensaje As String
    strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoVerificado, objUsuario.Codigo)
    If strMensaje = "" Then
        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
    Else
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    Exit Sub
  '  cmdBuscar_Click
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Dim tempCtrl  As Boolean, tempAlt As Boolean

On Error GoTo handle
    tempCtrl = (Shift And vbCtrlMask) > 0
    tempAlt = (Shift And vbAltMask) > 0

Select Case KeyCode
    Case vbKeyF1
        cmdVerificado_Click
    Case vbKeyF2
        cmdDetalle_Click
    Case vbKeyF3
        cmdLlevando_Click
    Case vbKeyF4
        cmdLlegadaDestino_Click
    Case vbKeyF5
        cmdEntregado_Click
    Case vbKeyF6
        cmdLlegadaLocal_Click
    Case vbKeyF7
        cmdLiberar_Click
    Case vbKeyF8
        CmdAnular_Click
    Case vbKeyF9
        cmdCambioLocalDesp_Click
    Case vbKeyF10
        cmdPedido_Click
    Case vbKeyF12
        CmdActualizar_Click
    Case tempCtrl And vbKeyTab
        If Option1.Value = True Then
            Option1.Value = False
            Option2.Value = True
        Else
            Option1.Value = True
            Option2.Value = False
        End If
    Case tempAlt And vbKeyR
        cmdReclamo_Click
        
End Select

Exit Sub

handle:

    MsgBox Err.Description, vbInformation + vbOKOnly, App.ProductName
End Sub

Sub Opciones_Formato(arrCampos As Variant, arrCaption As Variant, arrAncho As Variant, arrAlineacion As Variant, indice As Boolean, Optional arrFoco As Variant)
Dim columna As TrueDBGrid70.Column
    If indice = True Then
        arrCampos = Array("Sema", "FCH_REGISTRA", "NUM_PROFORMA", "NUM_PEDIDO_REF", "DES_ESTADO", "FLG_URGENTE_OP", "COD_LOCAL_REF", "COD_LOCAL_SAP_REF", "DES_LOCAL_REF", "DES_MOTORIZADO", "DES_AUX_CLI_DIRECC", "UBIGEO_LARGO", "FLG_NEC_TRANF", "NombreConvenio", "flg_convenio", "FCH_HORA_PACT_ENTR", "TE", "REFERENCIA", "FLG_NEC_TRANSF", "COD_ESTADO", "COD_LOCAL", "FLG_TRANSFERENCIA", "FLG_URGENTE", "OBS_NOTA_RUTEO", "COD_CIA_REF", "DES_AUX_CLI_NOMBRE")
        arrCaption = Array("Sema", "Fecha Generada", "Num. Pedido", "Pedido Ref.", "Estado", "URGENTE", "Local", "Local", "Des. Local", "Motorizado", "Dirección", "Dirección", "FLG_NEC_TRANF", "Convenio", "flg_convenio", "Fecha Entrega", "T.E.", "Referencia", "Tranf", "COD_ESTADO", "COD_LOCAL", "", "FLG_URGENTE", "Observacion Ruta", "", "Cliente")
        arrAncho = Array(300, 1550, 1000, 1000, 500, 200, 400, 800, 2000, 1000, 0, 3500, 700, 2000, 100, 1800, 600, 2200, 0, 0, 0, 0, 0, 0, 0, 2800)
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgLeft)
    Else
        arrCampos = Array("Sema", "FCH_REGISTRA", "NUM_PROFORMA", "NUM_PEDIDO_REF", "DES_ESTADO", "FLG_URGENTE_OP", "COD_LOCAL_REF", "COD_LOCAL_SAP_REF", "DES_LOCAL_REF", "DES_MOTORIZADO", "DES_AUX_CLI_DIRECC", "UBIGEO_LARGO", "FLG_NEC_TRANF", "NombreConvenio", "flg_convenio", "FCH_HORA_PACT_ENTR", "TE", "REFERENCIA", "FLG_NEC_TRANSF", "COD_ESTADO", "COD_LOCAL", "FLG_TRANSFERENCIA", "FLG_URGENTE", "OBS_NOTA_RUTEO", "COD_CIA_REF", "DES_AUX_CLI_NOMBRE")
        arrCaption = Array("Sema", "Fecha Generada", "Num. Pedido", "Pedido Ref.", "Estado", "URGENTE", "Local", "Local", "Des. Local", "Motorizado", "Dirección", "Dirección", "FLG_NEC_TRANF", "Convenio", "flg_convenio", "Fecha Entrega", "T.E.", "Referencia", "Tranf", "COD_ESTADO", "COD_LOCAL", "", "FLG_URGENTE", "Observacion Ruta", "", "Cliente")
        arrAncho = Array(300, 1550, 1000, 1000, 500, 200, 400, 800, 2000, 0, 0, 3500, 700, 2000, 100, 1800, 600, 0, 0, 0, 0, 0, 0, 3200, 0, 2800)
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgLeft)
    End If
    
    grdPedidos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdPedidos.RowHeight = 0
    grdPedidos.RowHeight = grdPedidos.RowHeight * 1.2
    grdPedidos.Columns(4).Font.Bold = True
    If Option1.Value = True Then
        grdPedidos.Columns(6).BackColor = pdblColorFondo
        grdPedidos.Columns(9).BackColor = pdblColorFondo
    End If
    grdPedidos.Columns(6).Visible = False
    grdPedidos.Columns(10).Visible = False
    'grdPedidos.Columns(8).AllowSizing = False
    grdPedidos.Columns(12).Visible = False
    'grdPedidos.Columns(10).AllowSizing = False
    grdPedidos.Columns(14).Visible = False
    grdPedidos.Columns(22).Visible = False
    grdPedidos.Columns(18).Visible = False
    grdPedidos.Columns(19).Visible = False
    grdPedidos.Columns(20).Visible = False
    grdPedidos.Columns(21).Visible = False
    grdPedidos.Columns(22).Visible = False
    grdPedidos.Columns(5).FetchStyle = True
    grdPedidos.Columns(18).FetchStyle = True
    grdPedidos.Columns(22).FetchStyle = True
    
    For Each columna In grdPedidos.Columns
        columna.AllowSizing = False
    Next
    
          
    If indice = True Then
        grdPedidos.Columns(9).Visible = True
        grdPedidos.Columns(17).Visible = True
        grdPedidos.Columns(23).Visible = False
        grdPedidos.MultiSelect = 2
    ElseIf indice = False Then
        grdPedidos.Columns(9).Visible = False
        grdPedidos.Columns(17).Visible = False
        grdPedidos.Columns(23).Visible = True
        grdPedidos.MultiSelect = 1
    End If
    
    grdPedidos.FetchRowStyle = True
    
    '''''''''''''''''''''''''''''''''
    grdPedidos.Columns(0).ValueItems.Translate = True
    Dim ValueItem3 As New TrueDBGrid70.ValueItem
    ValueItem3.DisplayValue = ImageList2.ListImages(3).Picture
    ValueItem3.Value = "1"
    grdPedidos.Columns(0).ValueItems.Add ValueItem3
    Set ValueItem3 = Nothing

    Dim ValueItem4 As New TrueDBGrid70.ValueItem
    ValueItem4.DisplayValue = ImageList2.ListImages(4).Picture
    ValueItem4.Value = "2"
    grdPedidos.Columns(0).ValueItems.Add ValueItem4
    Set ValueItem4 = Nothing
    
    Dim ValueItem5 As New TrueDBGrid70.ValueItem
    ValueItem5.DisplayValue = ImageList2.ListImages(5).Picture
    ValueItem5.Value = "3"
    grdPedidos.Columns(0).ValueItems.Add ValueItem5
    Set ValueItem5 = Nothing
    
    ''''''''''''''''''''''''''''''''''
    '    Dim ValueItem As New TrueDBGrid70.ValueItem
    '    Dim ValueItem1 As New TrueDBGrid70.ValueItem
    '
    '        grdPedidos.Columns(15).ValueItems.Translate = True
    '        ValueItem.DisplayValue = ImageList2.ListImages(1).Picture
    '        ValueItem.Value = "0"
    '        grdPedidos.Columns(15).ValueItems.Add ValueItem
    '        Set ValueItem = Nothing
    '
    '        ValueItem1.DisplayValue = ImageList2.ListImages(2).Picture
    '        ValueItem1.Value = "1"
    '        grdPedidos.Columns(15).ValueItems.Add ValueItem1
    '        Set ValueItem1 = Nothing
End Sub

Private Sub Form_Load()
   'agregado por pherrer para mensajes de DLV
    On Error Resume Next
        Set wskPrincipal = New CSocketMaster
        With wskPrincipal
            .Protocol = sckUDPProtocol
            .LocalPort = pstrPuerto
            .Bind
        End With
        '/////////////////////////////////////
On Error GoTo Control

    setteaFormulario Me
  '  DTPicker1.Value = objUsuario.sysdate
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant
    dtpInicio.Value = objUsuario.sysdate - 4
    dtpFin.Value = objUsuario.sysdate
    If Option1.Value = True Then
        Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
    Else
        Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, False, arrFoco)
    End If
    
    arrCampos = Array("Generado", "Verificado", "Asignado", "Avisado", "Proforma", "Llevando", "Llegada", "Entregado", "Llegada a Local", "Liberado", "Anulado", "Sistema", "Time")
    arrCaption = Array("Generado", "Verificado", "Asignado", "Avisado", "Proforma", "Llevando", "Llegada", "Entregado", "Llegada a Local", "Liberado", "Anulado", "Sistema", "Time")
    arrAncho = Array(1250, 1250, 1250, 1250, 1250, 1250, 1250, 1250, 1250, 1250, 1250, 0, 1250)
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
    grdTiempo.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdTiempo.Columns(11).Visible = False
    
    ListaPedido
    Dim objZona As New clsZona
    Set cboZona.RowSource = objZona.Lista
    cboZona.ListField = "DES_ZONA"
    cboZona.BoundColumn = "COD_ZONA"
    
    frm_DLV_Seguimiento.Caption = frm_DLV_Seguimiento.Caption & " " & gstrAplicacion & " * Ver: " & gstrVersion & " - " & gvarTNSNAME
    
    
   ' cmdBuscar_Click
    '''''''''tab dos
'    Dim objMotorizado As New clsMotorizado
'    Set cboMotorizado.RowSource = objMotorizado.Lista
'    cboMotorizado.BoundColumn = "COD_MOTORIZADO"
'    cboMotorizado.ListField = "NOMBRE"
'    Set objMotorizado = Nothing
    
'    Set cboRuta2.RowSource = objZona.Lista
'    cboRuta2.BoundColumn = "COD_ZONA"
'    cboRuta2.ListField = "DES_ZONA"
    EstadoBotones

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Sub ListaPedido()
    Dim Estado As String
    Dim posActual As Variant
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant
    
    
    'If Option2.Value = True Then Estado = objPedido.PedidoVerificado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, False, arrFoco)
    'If Option1.Value = True Then Estado = objPedido.PedidoAsignado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
    'I.ECASTILLO 17.12.2020
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "0" Then
        If Option2.Value = True Then Estado = objPedido.PedidoVerificado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, False, arrFoco)
        If Option1.Value = True Then Estado = objPedido.PedidoAsignado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
        If Option3.Value = True Then Estado = objPedido.PedidoVerificado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
    Else
        If Option2.Value = True Then Estado = objPedido.PedidoVerificado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, False, arrFoco)
        If Option1.Value = True Then Estado = objPedido.PedidoAsignado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
        If Option3.Value = True Then Estado = objPedido.PedidoAnulado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
    End If
    'F.ECASTILLO 17.12.2020
    posActual = grdPedidos.Bookmark
    Set grdPedidos.DataSource = objPedido.ListaDelivery(objUsuario.CodigoEmpresa, Estado, objUsuario.CodigoLocal, cboZona.BoundText, dtpInicio.Value, dtpFin.Value)
    'grdPedidos.Rebind
    'grdPedidos.Bookmark = posActual
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wskPrincipal.CloseSck
    Set objVenta = Nothing
    Set objUsuario = Nothing
    gclsOracle.Cerrar
End Sub

Private Sub grdPedidos_DblClick()
    cmdDetalle_Click
End Sub

Private Sub grdPedidos_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
Dim n As Integer
Dim f As Integer
    If Condition = 0 Then
        Select Case Col
            Case 4, 17, 21
              n = val(grdPedidos.Columns(21).CellText(Bookmark))
              f = val(grdPedidos.Columns(20).CellText(Bookmark))
              If n = 0 And f = 0 Then
                 CellStyle.ForeColor = RGB(255, 250, 250)
                 CellStyle.BackColor = RGB(255, 250, 250)
                 CellStyle.Font.Bold = True
              End If
              
              If n > 0 Then
                 CellStyle.ForeColor = RGB(220, 20, 60)
                 CellStyle.BackColor = RGB(220, 20, 60)
                 CellStyle.Font.Bold = True
              End If
              
              If n = 0 And f > 0 Then
                 CellStyle.ForeColor = RGB(50, 205, 50)
                 CellStyle.BackColor = RGB(50, 205, 50)
                 CellStyle.Font.Bold = True
              End If
              
              If n > 0 And f > 0 Then
                 CellStyle.ForeColor = RGB(0, 0, 205)
                 CellStyle.BackColor = RGB(0, 0, 205)
                 CellStyle.Font.Bold = True
              End If
              
        End Select
        
    End If
    If Condition = 2 Or Condition = 3 Or Condition = 10 Then
        Select Case Col
            Case 4, 17, 21
                    n = val(grdPedidos.Columns(21).CellText(Bookmark))
                    f = val(grdPedidos.Columns(20).CellText(Bookmark))
              If n = 0 And f = 0 Then
                 CellStyle.ForeColor = RGB(255, 250, 250)
                 CellStyle.BackColor = RGB(255, 250, 250)
                 CellStyle.Font.Bold = True
              End If
              
              If n > 0 Then
                 CellStyle.ForeColor = RGB(220, 20, 60)
                 CellStyle.BackColor = RGB(220, 20, 60)
                 CellStyle.Font.Bold = True
              End If
              
              If n = 0 And f > 0 Then
                 CellStyle.ForeColor = RGB(50, 205, 50)
                 CellStyle.BackColor = RGB(50, 205, 50)
                 CellStyle.Font.Bold = True
              End If
              
              If n > 0 And f > 0 Then
                 CellStyle.ForeColor = RGB(0, 0, 205)
                 CellStyle.BackColor = RGB(0, 0, 205)
                 CellStyle.Font.Bold = True
              End If
                    
                    
        End Select
    End If

End Sub

Private Sub grdPedidos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
On Error GoTo handle
    'If Val(grdPedidos.Columns(9).CellText(Bookmark)) = 1 Then
    '    'RowStyle.BackColor = &HFFF0E1
    'End If
    'If Val(grdPedidos.Columns(11).CellText(Bookmark)) = 1 Then
    '    'RowStyle.BackColor = &HE1F0FF
    'End If
    If val(grdPedidos.Columns("COD_CIA_REF").CellText(Bookmark)) = 99 Then
        RowStyle.BackColor = &HC0E0FF       '&H80FF&
    ElseIf val(grdPedidos.Columns("COD_CIA_REF").CellText(Bookmark)) = 94 Then
        RowStyle.BackColor = &H80FFFF
    End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdPedidos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdDetalle_Click
End Sub

Private Sub grdPedidos_RegistroSeleccionado(ByVal DatoColumna0 As String)
    '''If (grdPedidos.Columns("COD_ESTADO").Value = objPedido.PedidoLlevando) Or _
    ''' (grdPedidos.Columns("COD_ESTADO").Value = objPedido.PedidoLlegada) Or
    If (grdPedidos.Columns("COD_ESTADO").Value = objPedido.PedidoEntregado) Then
        cmdAnular.Enabled = False
        cmdLiberar.Enabled = False
    Else
        cmdAnular.Enabled = True
        cmdLiberar.Enabled = True
    End If

   'Set grdTiempo.DataSource = objPedido.ListaEstados(objUsuario.CodigoEmpresa, grdPedidos.Columns("NUM_PROFORMA"))
   
   
'   If grdPedidos.DataSource("OBS_NOTA_RUTEO").Value <> "" Then
'    ctlTextBox1.Text = "" & grdPedidos.DataSource("OBS_NOTA_RUTEO").Value & " INT " & grdPedidos.DataSource("NUM_PROFORMA_ORIG").Value
'   Else
'    ctlTextBox1.Text = ""
'   End If
   
   
End Sub

Private Sub mnuasiglocal_Click()
    frmGrabaLocalesPorZona.Show vbModal
End Sub

Private Sub mnucamblocaldesp_Click()
'
End Sub

Private Sub mnuMotorizados_Click()
    frmMotorizado.Show
End Sub

Private Sub mnuZonas_Click()
    frmZona.Show
End Sub

Private Sub Option1_Click()
    
On Error GoTo Control

    ListaPedido
    EstadoBotones
    ''''Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Option2_Click()
On Error GoTo Control

    ListaPedido
    EstadoBotones
    
    ''''Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, False, arrFoco)

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Option3_Click()
On Error GoTo Control
    ListaPedido
    EstadoBotones
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub EstadoBotones()
    If Option1.Value = True Then
        cmdVerificado.Enabled = False
        cmdDetalle.Enabled = True
        cmdAvisado.Enabled = True
        cmdLlevando.Enabled = True
        cmdLlegadaDestino.Enabled = True
        cmdEntregado.Enabled = True
        cmdLlegadaLocal.Enabled = True
        cmdLiberar.Enabled = True
        cmdAnular.Enabled = True
        cmdReclamo.Enabled = True
    End If
    
    If Option2.Value = True Then
        cmdVerificado.Enabled = True
        cmdDetalle.Enabled = True
        cmdAvisado.Enabled = False
        cmdLlevando.Enabled = False
        cmdLlegadaDestino.Enabled = False
        cmdEntregado.Enabled = False
        cmdLlegadaLocal.Enabled = False
        cmdLiberar.Enabled = False
        cmdAnular.Enabled = True
        cmdReclamo.Enabled = True
    End If
End Sub
