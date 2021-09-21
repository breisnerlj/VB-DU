VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_DLV_Verificacion 
   Caption         =   "Verificación "
   ClientHeight    =   10755
   ClientLeft      =   165
   ClientTop       =   -1530
   ClientWidth     =   15120
   Icon            =   "frm_DLV_Verificacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPedidos 
      Caption         =   "&Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13980
      Picture         =   "frm_DLV_Verificacion.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9960
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10635
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   18759
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Verificación"
      TabPicture(0)   =   "frm_DLV_Verificacion.frx":0894
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdSalir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdActualizar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdTarjetas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Motorizados"
      TabPicture(1)   =   "frm_DLV_Verificacion.frx":08B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label7"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label8"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label9"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label10"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label11"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label12"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label13"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label17"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cboMotorizado"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cboLocalAsig2"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cboRuta2"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "dtpFchIni"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "grdAsistenciaMotorizado"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdBuscar"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmdEstado(1)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdEstado(3)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmdEstado(4)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmdEstado(2)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cmdAusencia"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmdMotorizadosLocal"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "dtpFchFin"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtBusquedaMotorizado"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmdObservaciones"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chkTodos"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cboCia"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).ControlCount=   32
      Begin vbp_Ventas.ctlDataCombo cboCia 
         Height          =   315
         Left            =   -74160
         TabIndex        =   46
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.CommandButton cmdTarjetas 
         Caption         =   "&Reporte Tarjetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   9870
         Width           =   1335
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Incluir Salidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -63480
         TabIndex        =   43
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdObservaciones 
         Caption         =   "&Observaciones"
         Height          =   735
         Left            =   -64440
         Picture         =   "frm_DLV_Verificacion.frx":08CC
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   9000
         Width           =   1455
      End
      Begin vbp_Ventas.ctlTextBox txtBusquedaMotorizado 
         Height          =   375
         Left            =   -72480
         TabIndex        =   40
         Top             =   1560
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   661
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
      Begin MSComCtl2.DTPicker dtpFchFin 
         Height          =   345
         Left            =   -63240
         TabIndex        =   29
         Top             =   720
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   609
         _Version        =   393216
         Format          =   72876033
         CurrentDate     =   39057
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pago con tarjeta de credito por verificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4710
         Left            =   75
         TabIndex        =   20
         Top             =   345
         Width           =   15105
         Begin VB.Timer Timer1 
            Interval        =   40000
            Left            =   7560
            Top             =   1320
         End
         Begin VB.CommandButton CmdAnularTarj 
            Caption         =   "&Anular"
            Height          =   615
            Left            =   10320
            Picture         =   "frm_DLV_Verificacion.frx":0E56
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   4035
            Visible         =   0   'False
            Width           =   1095
         End
         Begin vbp_Ventas.ctlGrilla grdVerifTarjeta 
            Height          =   3495
            Left            =   45
            TabIndex        =   22
            Top             =   255
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   6165
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
         Begin vbp_Ventas.ctlTextBox txtObsCabina 
            Height          =   585
            Left            =   60
            TabIndex        =   23
            Top             =   4050
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   1032
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observación de Cabina"
            Height          =   195
            Left            =   60
            TabIndex        =   24
            Top             =   3810
            Width           =   1665
         End
      End
      Begin VB.CommandButton cmdMotorizadosLocal 
         Caption         =   "&Asignar Local"
         Height          =   735
         Left            =   -66165
         Picture         =   "frm_DLV_Verificacion.frx":13E0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton cmdAusencia 
         Caption         =   "&Registrar Ausencia"
         Height          =   735
         Left            =   -67860
         Picture         =   "frm_DLV_Verificacion.frx":196A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "&Salida"
         Height          =   735
         Index           =   2
         Left            =   -69555
         Picture         =   "frm_DLV_Verificacion.frx":1EF4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "&Ent. Refrigerio"
         Height          =   735
         Index           =   4
         Left            =   -71250
         Picture         =   "frm_DLV_Verificacion.frx":247E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "&Sal Refrigerio"
         Height          =   735
         Index           =   3
         Left            =   -72945
         Picture         =   "frm_DLV_Verificacion.frx":2A08
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "&Entrada"
         Height          =   735
         Index           =   1
         Left            =   -74640
         Picture         =   "frm_DLV_Verificacion.frx":2F92
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   600
         Left            =   -61320
         Picture         =   "frm_DLV_Verificacion.frx":351C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   540
         Width           =   1230
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "&Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5633
         Picture         =   "frm_DLV_Verificacion.frx":3AA6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   9870
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7950
         Picture         =   "frm_DLV_Verificacion.frx":4030
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   9870
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clientes Nuevos por Verificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   75
         TabIndex        =   1
         Top             =   5115
         Width           =   15120
         Begin VB.CommandButton CmdAnularCli 
            Caption         =   "&Anular"
            Height          =   615
            Left            =   10320
            Picture         =   "frm_DLV_Verificacion.frx":45BA
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   3885
            Visible         =   0   'False
            Width           =   1095
         End
         Begin vbp_Ventas.ctlTextBox txtObs 
            Height          =   585
            Left            =   75
            TabIndex        =   3
            Top             =   3900
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   1032
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin vbp_Ventas.ctlGrilla grdVerifCliente 
            Height          =   3375
            Left            =   45
            TabIndex        =   4
            Top             =   240
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   5953
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Observación de Cabina"
            Height          =   195
            Left            =   75
            TabIndex        =   5
            Top             =   3660
            Width           =   1665
         End
      End
      Begin vbp_Ventas.ctlGrilla grdAsistenciaMotorizado 
         Height          =   6735
         Left            =   -74655
         TabIndex        =   15
         Top             =   2160
         Width           =   14550
         _ExtentX        =   25665
         _ExtentY        =   11880
         Resalte         =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpFchIni 
         Height          =   345
         Left            =   -63240
         TabIndex        =   16
         Top             =   360
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65798145
         CurrentDate     =   39057
      End
      Begin vbp_Ventas.ctlDataCombo cboRuta2 
         Height          =   315
         Left            =   -68775
         TabIndex        =   17
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboLocalAsig2 
         Height          =   315
         Left            =   -74640
         TabIndex        =   18
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboMotorizado 
         Height          =   315
         Left            =   -71760
         TabIndex        =   19
         Top             =   1080
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label17 
         Caption         =   "Cia :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   45
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "F7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -63840
         TabIndex        =   42
         Top             =   9840
         Width           =   225
      End
      Begin VB.Label Label12 
         Caption         =   "Busqueda de Motorizado"
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
         Left            =   -74640
         TabIndex        =   39
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "F6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -65535
         TabIndex        =   37
         Top             =   9855
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -67230
         TabIndex        =   36
         Top             =   9855
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "F4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -68925
         TabIndex        =   35
         Top             =   9855
         Width           =   225
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "F3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -70620
         TabIndex        =   34
         Top             =   9855
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -72315
         TabIndex        =   33
         Top             =   9855
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -74010
         TabIndex        =   32
         Top             =   9855
         Width           =   225
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   -63840
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Left            =   -63840
         TabIndex        =   30
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Zona:"
         Height          =   195
         Left            =   -69375
         TabIndex        =   28
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Local"
         Height          =   195
         Left            =   -74640
         TabIndex        =   27
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Motorizado"
         Height          =   195
         Left            =   -71760
         TabIndex        =   26
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   -64680
         TabIndex        =   25
         Top             =   480
         Width           =   450
      End
   End
End
Attribute VB_Name = "frm_DLV_Verificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCliente As New clsCliente
Private objFormaPago As New clsFormaPago
Private objPedido As New clsProforma

Dim odynCli As oraDynaset
Dim odynTarj As oraDynaset
Dim bolPaso As Boolean
Dim sCia As String   ' jct para cia

Private Sub cboCia_Change()
 ' listar locales segun zona y cia or  listar todos or dejar vacio
 sCia = cboCia.BoundText
 sub_CargaLocal
 
End Sub

'''Public pNumProforma As String
'''Public pCodigoCli As String
'''Public pNombres As String
'''Public pApellidos As String
'''Public pDirecc As String
'''Public pDpto As String
'''Public pProv As String
'''Public pDist As String
'''Public pLocalRef As String
'''Public pTipoDoc As String
'''Public pOperadora As String
'''Public pMotivoRech As String
'''Public pTelefono As String
'''Public pBlnCliente As Boolean

Private Sub cboLocalAsig2_Click(Area As Integer)
On Error GoTo Control

    If cboLocalAsig2.BoundText = "" Or Area = 0 Then Exit Sub
    
    sub_CargaMotorizado "", cboLocalAsig2.BoundText
    
    cboMotorizado.BoundText = ""
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cboRuta2_Click(Area As Integer)
On Error GoTo Control

sCia = cboCia.BoundText
'MsgBox "sCia : " + sCia
If sCia = "" Then
  MsgBox "Seleccione una Cia", vbCritical, App.ProductName
  Exit Sub
  cboCia.SetFocus
End If
If Area = 0 Then Exit Sub
    sub_CargaLocal
    If cboRuta2.BoundText = "" Then
        sub_CargaMotorizado
    End If
    
    cboLocalAsig2.BoundText = ""
    cboMotorizado.BoundText = ""
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdAusencia_Click()
Dim msgbo As Variant
Dim objMotorizado As New clsMotorizado
Dim strMensaje As String

    On Error GoTo Handle
    msgbo = MsgBox("¿El Motorizado se encuentra ausente?", vbYesNoCancel + vbInformation, App.ProductName)
    If msgbo = vbCancel Then Exit Sub
    strMensaje = objMotorizado.GrabaAusencia(grdAsistenciaMotorizado.Columns(0).Value, IIf(msgbo = vbYes, True, False))
    If strMensaje = "" Then
        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
    Else
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    cmdBuscar_Click
    Set objMotorizado = Nothing
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdBuscar_Click()
    Dim objMotorizado As New clsMotorizado
    Dim posActual As Variant
On Error GoTo Control
    posActual = grdAsistenciaMotorizado.Bookmark
    Set grdAsistenciaMotorizado.DataSource = objMotorizado.ListaDisponible(cboLocalAsig2.BoundText, _
                                                                           CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                                           CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")), _
                                                                           chkTodos.Value)
    Set objMotorizado = Nothing
    grdAsistenciaMotorizado.Bookmark = posActual
    ''If grdAsistenciaMotorizado.ApproxCount <> 0 And SSTab1.Tab = 1 Then grdAsistenciaMotorizado.SetFocus
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdEstado_Click(Index As Integer)
Dim strCodigoMotorizado As String
On Error GoTo Control

    strCodigoMotorizado = "" & grdAsistenciaMotorizado.Columns(0).Value
    If Index = 1 Then strCodigoMotorizado = cboMotorizado.BoundText
    Dim objMorotizado As New clsMotorizado
    Dim mensaje As String
    mensaje = objMorotizado.GrabaEstado(strCodigoMotorizado, Index, objUsuario.Codigo)
    If mensaje = "" Then
        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
    Else
        MsgBox mensaje, vbCritical, App.ProductName
    End If
    
    '********************************************************************'
    '** Validando para cambiar al estado al motorizado según la acción **'
    '**             18/02/2008 Por Cristhian Rueda                     **'
    '********************************************************************'
    
     objUsuario.CodigoMotorizado = Trim(cboMotorizado.BoundText) 'objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, Trim(cboLocalAsig2.BoundText), grdPedidos.Columns("NUM_PROFORMA"))
     
     If objUsuario.CodigoMotorizado <> "" Then
        Dim vstrMensaje As String
        vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
                                                       objPedido.PedidoLlegadaLocal, _
                                                       objUsuario.Codigo)
        If vstrMensaje = "" Then
         Else
            MsgBox vstrMensaje, vbCritical, Caption
        End If
     End If
    '********************************************************************'
    
    cmdBuscar_Click
Set objMorotizado = Nothing
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdMotorizadosLocal_Click()
    Call grdAsistenciaMotorizado_DblClick
End Sub

Private Sub cmdObservaciones_Click()
    On Error GoTo Control
    If grdAsistenciaMotorizado.ApproxCount = 0 Then Exit Sub
    Dim varBookMark As Variant
    Dim frm As New frm_DLV_ObservacionMotorizados
        varBookMark = grdAsistenciaMotorizado.Bookmark
            frm.Codigo = grdAsistenciaMotorizado.Columns("COD_MOTORIZADO").Value
            frm.Fecha = grdAsistenciaMotorizado.Columns("FCH_ASIST").Value
            frm.Show vbModal
            cmdBuscar_Click
        grdAsistenciaMotorizado.Bookmark = varBookMark
        Set frm = Nothing
    Exit Sub
Control:
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdPedidos_Click()
On Error GoTo Handle
    frm_DLV_Pedido_x_Estado.Show
    'frm_VTA_DetallePedido.Show vbModal
    
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdTarjetas_Click()
    frm_DLV_BuscaTarjetas.Show
End Sub

Private Sub Form_Activate()
Dim Bookmark As Variant
    
                'Bookmark = grdVerifTarjeta.Bookmark
On Error GoTo Control

                psubActualiza
                'grdVerifTarjeta.Bookmark = Bookmark

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub Form_Load()

On Error GoTo Control

    setteaFormulario Me
    SeteaGrilla_Cli
    SeteaGrilla_Tarj
    CmdActualizar_Click
    
    dtpFchIni.Value = objUsuario.sysdate
    dtpFchFin.Value = dtpFchIni.Value
        
    SSTab1.Tab = 0
        
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant
    Dim Columna As TrueDBGrid70.Column
    
    arrCampos = Array("COD_MOTORIZADO", "NOMBRE", "COD_LOCAL", "FCH_ASISTENCIA", "FCH_HORA_INGRESO", "FCH_HORA_SAL_REF", "FCH_HORA_ING_REF", "FCH_HORA_SALIDA", "AUSENCIA", "DES_AUX_CLI_DIRECC", "COD_ESTADO_MOTORIZADO", "DES_ALIAS", "FLG_ACTIVO", "DES_NOMBRES", "DES_APELLIDOS", "DES_NUMERO", "FCH_ASIST", "OBSERVACIONES")
    arrCaption = Array("Codigo", "Nombre", "Local", "Fecha", "Entrada", "Sal. Refrigerio", "Ent. Refrigerio", "Salida", "Ausencia", "Tiempo Total", "COD_ESTADO_MOTORIZADO", "DES_ALIAS", "FLG_ACTIVO", "DES_NOMBRES", "DES_APELLIDOS", "DES_NUMERO", "FCH_ASIST", "Observaciones")
    arrAncho = Array(500, 3500, 500, 1300, 1300, 1300, 1300, 1300, 1300, 0, 0, 0, 0, 0, 0, 0, 0, 3500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    grdAsistenciaMotorizado.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdAsistenciaMotorizado.Columns("AUSENCIA").Visible = False
    grdAsistenciaMotorizado.Columns("DES_AUX_CLI_DIRECC").Visible = False
    grdAsistenciaMotorizado.Columns("COD_ESTADO_MOTORIZADO").Visible = False
    grdAsistenciaMotorizado.Columns("DES_ALIAS").Visible = False
    grdAsistenciaMotorizado.Columns("FLG_ACTIVO").Visible = False
    grdAsistenciaMotorizado.Columns("DES_NOMBRES").Visible = False
    grdAsistenciaMotorizado.Columns("DES_APELLIDOS").Visible = False
    grdAsistenciaMotorizado.Columns("DES_NUMERO").Visible = False
    grdAsistenciaMotorizado.Columns("FCH_ASIST").Visible = False
    grdAsistenciaMotorizado.MarqueeStyle = dbgHighlightRowRaiseCell
    For Each Columna In grdAsistenciaMotorizado.Columns
        Columna.AllowSizing = False
    Next
    
    frm_DLV_Verificacion.Caption = frm_DLV_Verificacion.Caption & " " & gstrAplicacion & " * Ver: " & gstrVersion & " - " & gvarTNSNAME
    
    'obj : carga de cias en combo
    'fch : 13-ABR-12
    'Aut : JCT
     Set cboCia.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
     cboCia.ListField = "Des"
     cboCia.BoundColumn = "Cod"
    
    'end
    
    cmdBuscar_Click
    
    sub_CargaMotorizado
    
    sub_CargaZona

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub CmdActualizar_Click()
Dim varBookMarkCliente As Variant
Dim varBookMarkVerifTarjeta As Variant

    On Error GoTo Control

    'varBookMarkCliente = grdVerifCliente.Bookmark
    'varBookMarkVerifTarjeta = grdVerifTarjeta.Bookmark
    psubActualiza
    'grdVerifCliente.Bookmark = varBookMarkCliente
'    grdVerifTarjeta.Bookmark = varBookMarkVerifTarjeta
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub grdAsistenciaMotorizado_DblClick()
Dim varBookMark As Variant
    If grdAsistenciaMotorizado.ApproxCount = 0 Then Exit Sub
    'bjct, 13-ABR-12
    'frmGrabaMotorizado.s
    'ejct
    
    varBookMark = grdAsistenciaMotorizado.Bookmark
    Call frmGrabaMotorizado.datos(grdAsistenciaMotorizado.Columns("COD_MOTORIZADO").Value, _
                                  grdAsistenciaMotorizado.Columns("COD_ESTADO_MOTORIZADO").Value, _
                                  grdAsistenciaMotorizado.Columns("COD_LOCAL").Value, _
                                  grdAsistenciaMotorizado.Columns("DES_NOMBRES").Value, _
                                  grdAsistenciaMotorizado.Columns("DES_APELLIDOS").Value, _
                                  grdAsistenciaMotorizado.Columns("DES_NUMERO").Value, _
                                  grdAsistenciaMotorizado.Columns("DES_ALIAS").Value, _
                                  grdAsistenciaMotorizado.Columns("FLG_ACTIVO").Value, _
                                  "-(Editar)", _
                                  cboCia.BoundText)
                            
    cmdBuscar_Click
    grdAsistenciaMotorizado.Bookmark = varBookMark

End Sub

'Private Sub grdAsistenciaMotorizado_GotFocus()
'    grdAsistenciaMotorizado.MarqueeStyle = dbgHighlightRow
'End Sub

Private Sub grdAsistenciaMotorizado_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            Call cmdEstado_Click(1)
        Case vbKeyF2
            Call cmdEstado_Click(3)
        Case vbKeyF3
            Call cmdEstado_Click(4)
        Case vbKeyF4
            Call cmdEstado_Click(2)
        Case vbKeyF5
            Call cmdAusencia_Click
        Case vbKeyF6
            Call cmdMotorizadosLocal_Click
        Case vbKeyF7
            Call cmdObservaciones_Click
    End Select
    
End Sub

'Private Sub grdAsistenciaMotorizado_LostFocus()
'    grdAsistenciaMotorizado.MarqueeStyle = dbgNoMarquee
'End Sub

Private Sub grdVerifCliente_DblClick()
     If grdVerifCliente.ApproxCount <= 0 Then Exit Sub
     frm_DLV_Verificacion_Cliente.pBlnCliente = True
     frm_DLV_Verificacion_Cliente.strCodigoCliente = "" & grdVerifCliente.Columns("COD_CLIENTE_DLV").Value
     frm_DLV_Verificacion_Cliente.strNumeroPedido = "" & grdVerifCliente.Columns("NUM_PROFORMA").Value
     frm_DLV_Verificacion_Cliente.strTelefono = "" & grdVerifCliente.Columns("DES_AUX_CLI_TLF").Value
     frm_VTA_DetallePedido.CodigoLocal = "" & odynCli("COD_LOCAL_REF")
     frm_DLV_Verificacion_Cliente.Show vbModal
     CmdActualizar_Click
End Sub

Private Sub grdVerifCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        grdVerifCliente_DblClick
    Case vbKeyF12
        MuestraDetalle grdVerifCliente.Columns(0).Value
    End Select
End Sub

Private Sub grdVerifCliente_RegistroSeleccionado(ByVal DatoColumna0 As String)
    txtObs.Text = "" & grdVerifCliente.DataSource("OBS_NOTA_VERIFICACION").Value
End Sub

Private Sub grdVerifTarjeta_DblClick()
Dim Bookmark As Variant
    On Error GoTo ERROR
    If grdVerifTarjeta.ApproxCount <= 0 Then Exit Sub
    
        '---------------------------------------------------------'
        ' Valores a pasar
        '---------------------------------------------------------'
'        frm_DLV_Verificacion_Tarjeta.xPedido.Caption = odynTarj("NUM_PROFORMA").Value
'        frm_DLV_Verificacion_Tarjeta.xCodCliente.Caption = odynTarj("COD_CLIENTE").Value
'        frm_DLV_Verificacion_Tarjeta.xNombre.Caption = odynTarj("NOMBRE").Value
'        frm_DLV_Verificacion_Tarjeta.xApellido.Caption = odynTarj("APELLIDO").Value
'        frm_DLV_Verificacion_Tarjeta.xDireccion.Caption = odynTarj("DIRECCION").Value
'        frm_DLV_Verificacion_Tarjeta.xDistrito.Caption = odynTarj("DISTRITO").Value
'        frm_DLV_Verificacion_Tarjeta.xCodLocalRef.Caption = odynTarj("COD_LOCAL_REF").Value
'        frm_DLV_Verificacion_Tarjeta.xCodTipoDoc.Caption = odynTarj("COD_TIPO_DOCUMENTO").Value
'        frm_DLV_Verificacion_Tarjeta.xDesAuxCliTlf.Caption = grdVerifTarjeta.Columns("DES_AUX_CLI_TLF").Value
'        frm_DLV_Verificacion_Tarjeta.xCodRetrazoTarj.Caption = "" & odynTarj("COD_RETRAZO_TARJETA").Value
    
     Dim frm As New frm_DLV_Verificacion_Tarjeta
     frm.carga grdVerifTarjeta.Columns("NUM_PROFORMA").Value, _
               grdVerifTarjeta.Columns("COD_CLIENTE").Value, _
               grdVerifTarjeta.Columns("NOMBRE").Value & " " & grdVerifTarjeta.Columns("APELLIDO").Value, _
               grdVerifTarjeta.Columns("DIRECCION").Value, _
               "" & grdVerifTarjeta.Columns("DISTRITO").Value, _
               "" & grdVerifTarjeta.Columns("COD_LOCAL_REF").Value, _
               "" & grdVerifTarjeta.Columns("COD_TIPO_DOCUMENTO").Value, _
               "" & grdVerifTarjeta.Columns("DES_AUX_CLI_TLF").Value, _
               "" & grdVerifTarjeta.Columns("COD_RETRAZO_TARJETA").Value, _
               Me.txtObsCabina.Text, _
               "" & grdVerifTarjeta.Columns("COD_DIRECCION_CLI").Value, _
               Me, objFormaPago, _
               "" & grdVerifTarjeta.Columns("COD_LOCAL_SAP_REF").Value
               
    Exit Sub
ERROR:
   MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Sub SeteaGrilla_Cli()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("NUM_PROFORMA", "FCH_REGISTRA", _
                      "DES_AUX_CLI_TLF", "DES_CLIENTE", _
                      "DIRECCION", "URBANIZACION", _
                      "DISTRITO", "TIEMPO", _
                      "COD_DIRECCION_CLI", "COD_CLIENTE_DLV")
                      
    arrCaption = Array("Proforma", "Fecha", _
                       "Telefono", "Cliente", _
                       "Dirección", "Urbanización", _
                       "Distrito", "Tiempo", _
                       "CodDireccion", "CodClienteDLV")
                       
    arrAncho = Array(1300, 1800, _
                     1200, 2500, _
                     2800, 2000, _
                     2000, 1800, _
                     0, 0)
                     
    arrAlineacion = Array(dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft)
                          
    grdVerifCliente.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Sub SeteaGrilla_Tarj()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("NUM_PROFORMA", "FCH_REGISTRA", _
                      "DES_AUX_CLI_TLF", "DES_CLIENTE", _
                      "DIRECCION", "URBANIZACION", _
                      "DISTRITO", "TIEMPO", _
                      "COD_CLIENTE", "COD_DIRECCION_CLI", _
                      "NOMBRE", "APELLIDO", _
                      "COD_LOCAL_REF", "COD_TIPO_DOCUMENTO", _
                      "COD_RETRAZO_TARJETA", "COD_LOCAL_SAP_REF")
                      
    arrCaption = Array("Proforma", "Fecha", _
                       "Telefono", "Cliente", _
                       "Dirección", "Urbanización", _
                       "Distrito", "Tiempo", _
                       "Cod. Cliente", "CodDirecc", _
                       "Nombre", "Apellido", _
                       "CodLocalRef", "CodTipoDoc", _
                       "CodRetrazoTarjeta", "CodLocalSapRef")
                       
    arrAncho = Array(1300, 1800, _
                     1500, 2500, _
                     2800, 2000, _
                     2000, 1200, _
                     900, 0, _
                     0, 0, _
                     0, 0, _
                     0, 0)
                     
    arrAlineacion = Array(dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft)
                          
    grdVerifTarjeta.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub grdVerifTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case vbKeyReturn
                grdVerifTarjeta_DblClick
           Case vbKeyF12
    End Select
End Sub

Private Sub MuestraDetalle(NumeroPedido As String)
    frm_VTA_DetallePedido.NumeroPedido = NumeroPedido
    frm_VTA_DetallePedido.ReCargaDetPedido
    frm_VTA_DetallePedido.Show
End Sub

Private Sub grdVerifTarjeta_RegistroSeleccionado(ByVal DatoColumna0 As String)
    txtObsCabina.Text = "" & grdVerifTarjeta.DataSource("OBS_NOTA_VERIFICACION").Value
End Sub

Private Function AnulaPedido(ByVal NumeroProforma As String, ByVal CodigoLocal As String) As Boolean
On Error GoTo Handle
    Dim objProforma As New clsProforma
    If objProforma.Anula(objUsuario.CodigoEmpresa, CodigoLocal, NumeroProforma, objUsuario.Codigo) = "" Then
        AnulaPedido = True
    Else
        AnulaPedido = False
    End If
    Set objProforma = Nothing
Exit Function
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    AnulaPedido = False
    Set objProforma = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set objVenta = Nothing
    Set objUsuario = Nothing
    gclsOracle.LimpiaParametros
    End
End Sub

Private Sub cmdSalir_Click()
    gclsOracle.Cerrar
    End
End Sub

Public Sub psubActualiza()

    Set odynCli = objCliente.ListaClixVerif(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
    Set grdVerifCliente.DataSource = odynCli
    
    Set odynTarj = objFormaPago.ListaTarjxVerif(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
    Set grdVerifTarjeta.DataSource = odynTarj
End Sub

Private Sub Timer1_Timer()
   bolPaso = Not (bolPaso)
   If bolPaso = True Then
      CmdActualizar_Click
      Beep
   End If
End Sub

Private Sub sub_CargaLocal()
    Dim objZona As New clsZona
    Dim sZona As String
    'Set cboLocalAsig2.RowSource = objZona.ListaLocal(objUsuario.CodigoEmpresa, cboRuta2.BoundText)
    ' bjct Local segun cia
'''    sZona = cboRuta2.BoundText
'''    If sZona = "" Then
'''     MsgBox "Seleccione una Zona", vbCritical, App.ProductName
'''     cboRuta2.SetFocus
'''     Exit Sub
'''    End If
    
    Set cboLocalAsig2.RowSource = objZona.ListaLocal(sCia, cboRuta2.BoundText)
    'ejct
    
    cboLocalAsig2.BoundColumn = "COD_LOCAL"
    cboLocalAsig2.ListField = "local_dex"
    Set objZona = Nothing
End Sub

Private Sub sub_CargaZona()
    Dim objZona As New clsZona
    Set cboRuta2.RowSource = objZona.Lista
    cboRuta2.BoundColumn = "COD_ZONA"
    cboRuta2.ListField = "DES_ZONA"
    Set objZona = Nothing
End Sub

Private Sub sub_CargaMotorizado(Optional ByVal strCodigoMotorizado As String, Optional ByVal strCodLocalAsig As String)
    Dim objMotorizado As New clsMotorizado
    Set cboMotorizado.RowSource = objMotorizado.Lista(strCodigoMotorizado, strCodLocalAsig)
    'bjct
     'MsgBox "strCodigoMotorizado, strCodLocalAsig : " + strCodigoMotorizado + " , " + strCodLocalAsig
    'ejct
    cboMotorizado.BoundColumn = "COD_MOTORIZADO"
    cboMotorizado.ListField = "NOMBRE"
    Set objMotorizado = Nothing
End Sub

Private Sub txtBusquedaMotorizado_KeyPress(KeyAscii As Integer)
    grdAsistenciaMotorizado.DataSource.FindFirst "NOMBRE like '%" & Trim(txtBusquedaMotorizado.Text) & "%'"
End Sub
