VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_VTA_Documento 
   BorderStyle     =   0  'None
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   DrawMode        =   15  'Merge Pen Not
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox chkDiasFuturos 
      Caption         =   "&Días futuros"
      Height          =   255
      Left            =   480
      TabIndex        =   69
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCapacidades 
      Caption         =   "Capacidades"
      Height          =   315
      Left            =   1920
      TabIndex        =   68
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox txtTipoDocumento 
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Top             =   7290
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCodDireccion 
      Height          =   255
      Left            =   3540
      TabIndex        =   36
      Top             =   60
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4395
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7752
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos del Documento [F10]"
      TabPicture(0)   =   "frm_VTA_Documento.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblTitle(0)"
      Tab(0).Control(1)=   "lblTitle(2)"
      Tab(0).Control(2)=   "lblTitle(1)"
      Tab(0).Control(3)=   "lblTitle(3)"
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(6)=   "lblTitle(4)"
      Tab(0).Control(7)=   "lblRazones"
      Tab(0).Control(8)=   "txtDato(1)"
      Tab(0).Control(9)=   "txtDato(2)"
      Tab(0).Control(10)=   "txtDato(0)"
      Tab(0).Control(11)=   "ctlCboTipCliente"
      Tab(0).Control(12)=   "cmdBuscarRUC"
      Tab(0).Control(13)=   "txtCliente"
      Tab(0).Control(14)=   "cboRazonSocial"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Datos de Entrega [F11]"
      TabPicture(1)   =   "frm_VTA_Documento.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label16"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label18"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label19"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblTipoServicio"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblValTipoServicio"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "DTPicker4"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "DTPicker3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cboDistrito"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cboDepartamento"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cboProvincia"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cboPais"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "DTPicker2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "ctlTextBox4"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "ctlTextBox3"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "ctlTextBox2"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "ctlTextBox1"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "DTPicker1"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Check2"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Check1"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "grdDirecciones"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Check3"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chkFlgFchPactada"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtMetodoEntrega"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmdEditMetodo"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "Observaciones [F12]"
      TabPicture(2)   =   "frm_VTA_Documento.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label17"
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(2)=   "Label14"
      Tab(2).Control(3)=   "Label13"
      Tab(2).Control(4)=   "txtObservaciones(3)"
      Tab(2).Control(5)=   "txtObservaciones(2)"
      Tab(2).Control(6)=   "txtObservaciones(1)"
      Tab(2).Control(7)=   "txtObservaciones(0)"
      Tab(2).Control(8)=   "cmdCopiarObservaciones"
      Tab(2).ControlCount=   9
      Begin VB.CommandButton cmdEditMetodo 
         Caption         =   "Editar método"
         Height          =   375
         Left            =   5400
         TabIndex        =   73
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtMetodoEntrega 
         ForeColor       =   &H80000011&
         Height          =   315
         Left            =   1680
         TabIndex        =   72
         Text            =   "PROG miércoles 06/09 - 05:30pm a 08:30pm"
         Top             =   3730
         Width           =   3615
      End
      Begin VB.CheckBox chkFlgFchPactada 
         Caption         =   "Fecha &Pactada"
         Height          =   255
         Left            =   1920
         TabIndex        =   62
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdCopiarObservaciones 
         Caption         =   "Copiar &observaciones"
         Height          =   255
         Left            =   -70500
         TabIndex        =   61
         Top             =   1380
         Width           =   2055
      End
      Begin vbp_Ventas.ctlDataCombo cboRazonSocial 
         Height          =   315
         Left            =   -73380
         TabIndex        =   59
         Top             =   2400
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Pedido &Urgente"
         Height          =   255
         Left            =   1080
         TabIndex        =   45
         Top             =   720
         Width           =   1770
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   495
         Index           =   0
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         ToolTipText     =   "Observaciones para el local"
         Top             =   840
         Width           =   6315
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   495
         Index           =   1
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         ToolTipText     =   "Observaciones para Ruta"
         Top             =   1665
         Width           =   6315
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   495
         Index           =   2
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         ToolTipText     =   "Observaciones para el Motorizado"
         Top             =   2490
         Width           =   6315
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   495
         Index           =   3
         Left            =   -74760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         ToolTipText     =   "Observaciones para Verificación"
         Top             =   3360
         Width           =   6315
      End
      Begin vbp_Ventas.ctlGrilla grdDirecciones 
         Height          =   375
         Left            =   1080
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   661
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlTextBox txtCliente 
         Height          =   375
         Left            =   -74640
         TabIndex        =   1
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
         MaxLength       =   200
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
      Begin VB.CommandButton cmdBuscarRUC 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   -69660
         TabIndex        =   2
         Top             =   1110
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Entrega a &Tercero"
         Height          =   255
         Left            =   3600
         TabIndex        =   46
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Entrega en &Local"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboTipCliente 
         Height          =   315
         Left            =   -73380
         TabIndex        =   3
         Top             =   1980
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   3480
         TabIndex        =   56
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39044
      End
      Begin vbp_Ventas.ctlTextBox ctlTextBox1 
         Height          =   375
         Left            =   1080
         TabIndex        =   47
         Top             =   960
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
         MaxLength       =   200
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
      Begin vbp_Ventas.ctlTextBox ctlTextBox2 
         Height          =   375
         Left            =   1080
         TabIndex        =   48
         Top             =   1320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
         MaxLength       =   200
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
      Begin vbp_Ventas.ctlTextBox ctlTextBox3 
         Height          =   375
         Left            =   1440
         TabIndex        =   54
         Top             =   2925
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         Enabled         =   0   'False
         MaxLength       =   60
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
      Begin vbp_Ventas.ctlTextBox ctlTextBox4 
         Height          =   375
         Left            =   1440
         TabIndex        =   55
         Top             =   3330
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         Enabled         =   0   'False
         MaxLength       =   12
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3960
         TabIndex        =   57
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39044
      End
      Begin vbp_Ventas.ctlDataCombo cboPais 
         Height          =   315
         Left            =   240
         TabIndex        =   50
         Top             =   1980
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboProvincia 
         Height          =   315
         Left            =   240
         TabIndex        =   52
         Top             =   2580
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboDepartamento 
         Height          =   315
         Left            =   3000
         TabIndex        =   51
         Top             =   1980
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboDistrito 
         Height          =   315
         Left            =   3000
         TabIndex        =   53
         Top             =   2580
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlTextBox txtDato 
         Height          =   315
         Index           =   0
         Left            =   -73380
         TabIndex        =   4
         Top             =   2880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Tipo            =   3
         MaxLength       =   11
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
      Begin vbp_Ventas.ctlTextBox txtDato 
         Height          =   315
         Index           =   2
         Left            =   -73380
         TabIndex        =   6
         Top             =   3645
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         MaxLength       =   200
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
      Begin vbp_Ventas.ctlTextBox txtDato 
         Height          =   315
         Index           =   1
         Left            =   -73380
         TabIndex        =   5
         Top             =   3262
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         MaxLength       =   150
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
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   5040
         TabIndex        =   64
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm:ss tt"
         Format          =   65208323
         UpDown          =   -1  'True
         CurrentDate     =   39044
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   315
         Left            =   5400
         TabIndex        =   65
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65208322
         CurrentDate     =   39044
      End
      Begin VB.Label lblValTipoServicio 
         Caption         =   "* Capacidad ya no disponible elige otro horario"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1680
         TabIndex        =   71
         Top             =   4100
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label lblTipoServicio 
         Caption         =   "Metodo de entrega"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   3820
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Recojo"
         Height          =   195
         Left            =   5400
         TabIndex        =   67
         Top             =   4440
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Entrega"
         Height          =   195
         Left            =   5040
         TabIndex        =   66
         Top             =   4440
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fec Pact Recojo"
         Height          =   195
         Left            =   3960
         TabIndex        =   63
         Top             =   4440
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblRazones 
         Caption         =   "Raz. Sociales del cliente :"
         Height          =   375
         Left            =   -74700
         TabIndex        =   60
         Top             =   2370
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74760
         TabIndex        =   44
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
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
         Height          =   195
         Left            =   -74760
         TabIndex        =   43
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74760
         TabIndex        =   42
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Verificación"
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
         Height          =   195
         Left            =   -74760
         TabIndex        =   41
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F2"
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
         Index           =   4
         Left            =   -74820
         TabIndex        =   35
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Datos de Impresión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74340
         TabIndex        =   34
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   -74640
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Index           =   3
         Left            =   -74700
         TabIndex        =   32
         Top             =   3705
         Width           =   720
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "RUC:"
         Height          =   195
         Index           =   1
         Left            =   -74700
         TabIndex        =   31
         Top             =   2940
         Width           =   390
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social:"
         Height          =   195
         Index           =   2
         Left            =   -74700
         TabIndex        =   30
         Top             =   3322
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1050
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1410
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefóno"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   3420
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fec Pact Entrega"
         Height          =   195
         Left            =   3480
         TabIndex        =   25
         Top             =   4440
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         Height          =   195
         Left            =   3000
         TabIndex        =   24
         Top             =   2340
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   3000
         TabIndex        =   23
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Provincía"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2340
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Pais"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1740
         Width           =   300
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Persona:"
         Height          =   195
         Index           =   0
         Left            =   -74700
         TabIndex        =   20
         Top             =   2040
         Width           =   990
      End
   End
   Begin VB.TextBox txtCodigoCliente 
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4800
      Picture         =   "frm_VTA_Documento.frx":0054
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   6105
      Picture         =   "frm_VTA_Documento.frx":05DE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdMasDatos 
      Caption         =   "&Mas Datos"
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   315
      Left            =   5100
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla grdFormaPago 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   300
      Width           =   180
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
      Left            =   4740
      TabIndex        =   15
      Top             =   7260
      Width           =   1215
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
      Left            =   6450
      TabIndex        =   14
      Top             =   7260
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F9"
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
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F5"
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
      Index           =   5
      Left            =   4740
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_Documento.frx":0B68
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Documento"
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
      TabIndex        =   10
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCliente As New clsCliente
Dim objDocumento As New clsDocumento
Dim objMaquina As New clsMaquina
Public blnTipoDoc As Boolean
Public strDlvDocumento As String
Public pblnEditCli As Boolean
Dim objLocal As New clsLocal
'Dim objWS As New clsWebService

Private Sub cboRazonSocial_Change()
    txtDato(0).Text = cboRazonSocial.BoundText
    txtDato(1).Text = cboRazonSocial.Text
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        ctlTextBox1.Enabled = True
        ctlTextBox2.Enabled = True
        ctlTextBox3.Enabled = True
        ctlTextBox4.Enabled = True
        cboPais.Enabled = True
        cboDepartamento.Enabled = True
        cboProvincia.Enabled = True
        cboDistrito.Enabled = True
        ''ctlTextBox1.SetFocus
    Else
        ctlTextBox1.Enabled = False
        ctlTextBox2.Enabled = False
        ctlTextBox3.Enabled = False
        ctlTextBox4.Enabled = False
        cboPais.Enabled = False
        cboDepartamento.Enabled = False
        cboProvincia.Enabled = False
        cboDistrito.Enabled = False
    End If
End Sub

Private Sub ctlTextBox5_KeyPress(KeyAscii As Integer)

End Sub

Private Sub cmdRazonSocial_Click()

End Sub

Private Sub chkFlgFchPactada_Click()
    If chkFlgFchPactada.Value = 1 Then
        Label18.Visible = True: Label6.Visible = True
        DTPicker1.Visible = True: DTPicker3.Visible = True
'        If objVenta.isLocalDcCappa = "1" Then
'            cmdCapacidades.Visible = True
'            chkDiasFuturos.Visible = True
'        End If
      Else
        Label18.Visible = False: Label6.Visible = False
        DTPicker1.Visible = False: DTPicker3.Visible = False
        cmdCapacidades.Visible = False
        chkDiasFuturos.Visible = False
    End If
End Sub
'I.ECASTILLO 27.10.2020
Private Sub cmdCapacidades_Click()
    'consume servicio de capacidades, por default sin nueva ventana
    Dim obj As New Dictionary
    Dim codLocalPosu As String
    Dim segmento As String
    Dim horaActual As String
    Dim horaEstimada As String
    Dim Tipo As String
    Tipo = IIf(chkDiasFuturos.Value = 1, "AM_PM", "EXP")
    horaActual = DateTime.Now '(Format(Now, "hh:mm"))
    horaEstimada = Format(DTPicker1.Value, "dd/mm/yyyy ") & Format(DTPicker3.Value, "hh:mm:ss AMPM")
    segmento = DateTime.dateDiff("n", horaActual, horaEstimada) 'hora pactada - hora actual
    If segmento < 0 And Tipo = "EXP" Then
        MsgBox "Para realizar esta acción debe ingresar datos de fecha/hora pactada " & vbNewLine & "(tener en cuenta que no se debe ingresar una fecha/hora anterior).", vbOKOnly + vbInformation
        Exit Sub
    End If
    codLocalPosu = objLocal.GetCodPosu(mdiPrincipal.ctlCliente1.LocalDespacho)
    'al consultar para días posteriores el response retorna datos suficientes para mostrar nueva ventana
    'Set obj = objWS.listaCapacidades("B88", segmento, tipo)
    frm_VTA_ListaCapacidades.Datos codLocalPosu, segmento, Tipo
    frm_VTA_ListaCapacidades.Show
End Sub
'F.ECASTILLO 27.10.2020
Private Sub cmdCopiarObservaciones_Click()

On Error GoTo Control
txtObservaciones(1).Text = txtObservaciones(0).Text
txtObservaciones(2).Text = txtObservaciones(0).Text
txtObservaciones(3).Text = txtObservaciones(0).Text

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdEditMetodo_Click()
    'Abrir ventana metodos y segmentos con parametro local
    
    'If mdiPrincipal.ctlCliente1.seleccionManualLocal = True Then
        frm_VTA_MetodosSegmentos.Parametro = objLocal.GetCodPosu(mdiPrincipal.ctlCliente1.LocalDespacho)
        frm_VTA_MetodosSegmentos.Tipo = 3
        frm_VTA_MetodosSegmentos.permiteCerrar = "1"
        frm_VTA_MetodosSegmentos.Show vbModal
        txtMetodoEntrega.Text = objVenta.bk_ServiceType & " " & Format(objVenta.bk_FechaCapacidad, "dddd dd/mm") & Format(objVenta.bk_HoraCapacidad, " - hh:mm am/pm") & Format(objVenta.bk_HoraCapacidad2, " a hh:mm am/pm")
        If gstrIndRAv3 = "2" Then
            Dim flgCapacidadDisp As String
            flgCapacidadDisp = objVenta.validaCapacidad
            lblValTipoServicio.Visible = False
            If flgCapacidadDisp <> "1" Then
                lblValTipoServicio.Visible = True
            End If
        End If
    'End If
End Sub

Private Sub ctlTextBox2_KeyPress(KeyAscii As Integer)
''''Dim objCliente As New clsCliente
''''    Dim strNombre As String
''''    strNombre = Trim(objVenta.CodigoCliente)
''''    If Len(strNombre) > 3 Then
''''        Set grdDirecciones.DataSource = objCliente.ListaDireccion(strNombre) 'strflgActivo)
''''        If grdDirecciones.DataSource.RecordCount > 0 Then grdDirecciones.Visible = True
''''    Else
''''        grdDirecciones.Visible = False
''''    End If
''''Set objCliente = Nothing

End Sub
Private Sub ctlTextBox2_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If KeyCode = vbKeyDown And grdDirecciones.Visible = True Then grdDirecciones.SetFocus
End Sub


Private Sub Form_Activate()
'''    penumVentCli = Documento
'''    grdFormaPago.SetFocus
    'I.ECASTILLO 27.10.2020
    cmdCapacidades.Visible = False
    chkDiasFuturos.Visible = False
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "0" Then
'        If objVenta.isLocalDcCappa = "1" Then
'            SSTab1.Tab = 1
'            'cmdCapacidades.Visible = True
'            'chkDiasFuturos.Visible = True
'            If objVenta.flgDatosCapacidad = False Then cmdAceptar.Enabled = False
'        Else
            SSTab1.Tab = 0
            ''cmdCapacidades.Visible = False
            ''chkDiasFuturos.Visible = False
            cmdAceptar.Enabled = True
'        End If
    Else
        SSTab1.Tab = 0
        cmdAceptar.Enabled = True
    End If
    'F.ECASTILLO 27.10.2020
    
    'ECASTILLO 06.05.2020 - VALIDAR SI LOCAL ES INKA Y AUN NO FUE MIGRADO
    'DE SER EL CASO LIMITAR LENGTH DE OBS A 100
    If objLocal.EvaluaLocalInkMig(mdiPrincipal.ctlCliente1.LocalDespacho) = "N" Then
        txtObservaciones(0).MaxLength = 100
        If txtObservaciones(0).Text <> "" Then
            txtObservaciones(0).Text = Mid(txtObservaciones(0).Text, 1, 100)
        End If
    Else
        txtObservaciones(0).MaxLength = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "LENOBSDLV") '120 '200
        txtObservaciones(0).MaxLength = IIf(Len(Trim(txtObservaciones(0).MaxLength)) = 0, 120, txtObservaciones(0).MaxLength)
        txtObservaciones(1).MaxLength = txtObservaciones(0).MaxLength
        txtObservaciones(2).MaxLength = txtObservaciones(0).MaxLength
        txtObservaciones(3).MaxLength = txtObservaciones(0).MaxLength
    End If
    'I.ECASTILLO 17.12.2020
    'If Len(Trim(objVenta.bk_ServiceType)) = 0 Then
    '    lblTipoServicio.Visible = False
    '    lblValTipoServicio.Visible = False
    'Else
    '    lblTipoServicio.Visible = True
    '    lblValTipoServicio.Visible = True
        'lblValTipoServicio.Caption = objVenta.bk_ServiceType & "    " & objVenta.bk_FechaCapacidad & " " & objVenta.bk_HoraCapacidad & " - " & objVenta.bk_HoraCapacidad2
        txtMetodoEntrega.Text = objVenta.bk_ServiceType & " " & Format(objVenta.bk_FechaCapacidad, "dddd dd/mm") & Format(objVenta.bk_HoraCapacidad, " - hh:mm am/pm") & Format(objVenta.bk_HoraCapacidad2, " a hh:mm am/pm")
        If gstrIndRAv3 = "2" Then
        Dim flgCapacidadDisp As String
        flgCapacidadDisp = objVenta.validaCapacidad
        lblValTipoServicio.Visible = False
        If flgCapacidadDisp <> "1" Then
            lblValTipoServicio.Visible = True
        End If
        End If
    'End If
    'F.ECASTILLO 17.12.2020
End Sub
Private Sub cboDepartamento_Change()
Set cboProvincia.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, cboDepartamento.BoundText)
    cboProvincia.ListField = "Descripcion"
    cboProvincia.BoundColumn = "Codigo"
End Sub
Private Sub cboProvincia_Change()
    Set cboDistrito.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, cboDepartamento.BoundText, cboProvincia.BoundText)
    cboDistrito.ListField = "Descripcion"
    cboDistrito.BoundColumn = "Codigo"
End Sub


Private Sub Form_Load()
    setteaFormulario Me
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

chkFlgFchPactada.Value = 0

DTPicker1.Value = objUsuario.sysdate

    blnTipoDoc = True

           arrCampos = Array("DES_DIRECCION", "COD_DIRECCION_CLI", "COD_UBIGEO")
           arrCaption = Array("Direccción", "COD_DIRECCION_CLI", "COD_UBIGEO")
           arrAncho = Array(5000, 0, 0)
                             
           arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft)
           grdDirecciones.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
           grdDirecciones.HeadLines = 0
           grdDirecciones.MarqueeStyle = dbgNoMarquee
    
    Set cboPais.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PAIS", 0)
    cboPais.ListField = "Descripcion"
    cboPais.BoundColumn = "Codigo"
    cboPais.BoundText = "00"
    Set cboDepartamento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0)
    cboDepartamento.ListField = "Descripcion"
    cboDepartamento.BoundColumn = "Codigo"

    
    Call SetGrd
    Set grdFormaPago.DataSource = objDocumento.ListaTipoDocVta(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.NombrePC)
    peTransacc = GbDocumento
    
    'function que carga el combo de tipo de cliente
    Set ctlCboTipCliente.RowSource = objCliente.ListaTipo
    
    ctlCboTipCliente.ListField = "DES"
    ctlCboTipCliente.BoundColumn = "COD"
    
    
    grdFormaPago.DataSource.FindFirst "CODIGO='" & objVenta.CodigoDocumentoVenta & "'"
    Set cboRazonSocial.RowSource = objCliente.ListaHistoricoRazonSocial(objVenta.CodigoCliente)
    cboRazonSocial.ListField = "DES_RAZON_SOCIAL"
    cboRazonSocial.BoundColumn = "DES_RUC_EMPRESA"


With objVenta
    txtTipoDocumento = .CodigoDocumentoVenta
    'If Not .CodigoCliente = "" Then
        Dim rsCliente As oraDynaset
        Set rsCliente = objCliente.Lista(objVenta.CodigoCliente)

        .NombreCliente = "" & rsCliente("DES_NOM_CLIENTE").Value & "  " & rsCliente("DES_APE_CLIENTE").Value
        .Ruc = "" & rsCliente("NUM_DOCUMENTO_ID").Value
        .DireccionCliente = "" & rsCliente("DES_DIRECCION_COMERCIAL").Value
        .DireccionClienteSocial = "" & rsCliente("DES_DIRECCION_SOCIAL").Value
        .RazonSocial = "" & rsCliente("DES_RAZON_SOCIAL").Value
        .TipoCliente = "" & rsCliente("FLG_TIPO_JURIDICA").Value
        .NumeroDocumentoID = "" & rsCliente("NUM_DOCUMENTO_ID").Value
        If objUsuario.TipoMaquina <> objUsuario.TipoMaquinaCabina Then
            .UbigeoEntrega = "" & rsCliente("UBIGEO").Value
            .DesReferenciaCli = "" & rsCliente("DES_REFERENCIA").Value
        End If
        
        
        
        
        

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''    TxtCliente.Text = objVenta.NombreCliente
'''''    ctlCboTipCliente.BoundText = objVenta.TipoCliente
'''''    If objVenta.TipoCliente = "1" Then
'''''        TxtDato(0).Text = objVenta.Ruc
'''''        TxtDato(1).Text = objVenta.RazonSocial
'''''        TxtDato(2).Text = objVenta.DireccionClienteSocial
'''''    Else
'''''        TxtDato(0).Text = objVenta.NumeroDocumentoID
'''''        TxtDato(1).Text = objVenta.NombreCliente
'''''        TxtDato(2).Text = objVenta.DireccionCliente
'''''    End If
    
    
    txtCliente.Text = objVenta.Out_NombreCliente
    ctlCboTipCliente.BoundText = objVenta.Out_Tipo
    txtDato(0).Text = objVenta.Out_NumeroId
    txtDato(1).Text = objVenta.Out_NombreCliente
    
    'If objVenta.CodModalidadVenta = Venta_Convenio Then
    'If IIf(objVenta.CodModalidadVenta <> "", objVenta.CodModalidadVenta, 0) = Venta_Convenio Then
    Debug.Print .CodModalidadVenta
    Debug.Print objVenta.CodModalidadVenta
    .CodModalidadVenta = IIf(.CodModalidadVenta <> "", .CodModalidadVenta, codModalidadVentaBK)
    If .CodModalidadVenta = Venta_Convenio Then
        pblnEditCli = True
    End If
    
    
    If pblnEditCli = True Then txtDato(2).Text = objVenta.Out_Direccion
    
    'ECASTILLO 06.05.2020 - VALIDAR SI LOCAL ES INKA Y AUN NO FUE MIGRADO
    'DE SER EL CASO LIMITAR LENGTH DE OBS A 100
    If objLocal.EvaluaLocalInkMig(mdiPrincipal.ctlCliente1.LocalDespacho) = "N" Then
        txtObservaciones(0).MaxLength = 100
        If txtObservaciones(0).Text <> "" Then
            txtObservaciones(0).Text = Mid(txtObservaciones(0).Text, 1, 100)
        End If
    Else
        txtObservaciones(0).MaxLength = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "LENOBSDLV") '120 '200
        txtObservaciones(0).MaxLength = IIf(Len(Trim(txtObservaciones(0).MaxLength)) = 0, 120, txtObservaciones(0).MaxLength)
        txtObservaciones(1).MaxLength = txtObservaciones(0).MaxLength
        txtObservaciones(2).MaxLength = txtObservaciones(0).MaxLength
        txtObservaciones(3).MaxLength = txtObservaciones(0).MaxLength
    End If
    'I.ECASTILLO 17.12.2020
    
    txtObservaciones(0).Text = objVenta.ObsNotaLocal
    txtObservaciones(1).Text = objVenta.ObsNotaRuteo
    txtObservaciones(2).Text = objVenta.ObsNotaMotorizado
    txtObservaciones(3).Text = objVenta.ObsNotaVerificacion
    
    
    
    
'''''''    TxtRazonSocial.Text = objVenta.RazonSocial
'''''''    txtDNI.Text = objVenta.NumeroDocumentoID
'''''''    txtNombres.Text = objVenta.NombreCliente
'''''''    txtApellidos.Text = objVenta.ApellidoCliente
'''''''    txtDireccionPJ.Text = objVenta.DireccionClienteSocial
'''''''    txtDireccionPN.Text = objVenta.DireccionCliente
    txtCodigoCliente.Text = objVenta.CodigoCliente
'''End If



'End If

If Not .CodigoDocumentoVenta = "" Then

''''''    grdFormaPago.DataSource.FindFirst "CODIGO='" & .CodigoDocumentoVenta & "'"
''''''    ctlCboTipCliente.BoundText = .TipoCliente
''''''    '*************************************************************
''''''    '*************************************************************
''''''    '*************************************************************
''''''    '*************************************************************
''''''        .TipoCliente = ctlCboTipCliente.BoundText
''''''    If .TipoCliente = "1" Then
''''''        txtDato(0).Text = .Ruc
''''''        txtDato(1).Text = .RazonSocial
''''''        txtDato(2).Text = .DireccionClienteSocial
''''''
''''''    Else
''''''        txtDato(0).Text = .NumeroDocumentoID
''''''        txtDato(1).Text = .NombreCliente
''''''        txtDato(2).Text = .DireccionCliente
''''''    End If
''''''       txtCodigoCliente.Text = objVenta.CodigoCliente
        txtTipoDocumento.Text = .CodigoDocumentoVenta

        
        ctlTextBox1.Text = .NombreClienteDLV
        ctlTextBox2.Text = .DireccionClienteDLV
        ctlTextBox3.Text = .DesReferenciaCli
        ctlTextBox4.Text = .DesAuxCliTlf
        
        
        
        If Not .UbigeoEntrega = "" Then
            Dim sUbigeo As String
            sUbigeo = .UbigeoEntrega
            cboDepartamento.BoundText = Mid(sUbigeo, 1, 2)
            cboProvincia.BoundText = Mid(sUbigeo, 3, 2)
            cboDistrito.BoundText = Mid(sUbigeo, 5, 2)
        End If

    If chkFlgFchPactada.Value = 1 Then
        DTPicker1.Visible = True: DTPicker3.Visible = True
        Label18.Visible = True: Label6.Visible = True
        
        
      Else
      
        DTPicker1.Visible = False: DTPicker3.Visible = False
        Label18.Visible = False: Label6.Visible = False
   End If
   
    ''If .FchHoraPactEntr <> "" Then DTPicker1.Value = Format(.FchHoraPactEntr, "dd/mm/yyyy") & Format(.HoraPactEntr, "hh:mm:ss AMPM")
    ''If .FchHoraPactRecog <> "" Then DTPicker2.Value = Format(.FchHoraPactRecog, "dd/mm/yyyy") & Format(.HoraPactRecog, "hh:mm:ss AMPM")
    
    
    If .FchHoraPactEntr <> "" Then DTPicker1.Value = CDate(Format(.FchHoraPactEntr, "dd/mm/yyyy"))
    
    If .HoraPactEntr <> "" Then DTPicker3.Value = CDate(Format(.HoraPactEntr, "hh:mm:ss AMPM"))
    If .FlgPactado <> "" Then chkFlgFchPactada.Value = .FlgPactado
    If .FlgEntregaLocal <> "" Then Check2.Value = .FlgEntregaLocal
    If .FlgUrgente <> "" Then Check3.Value = .FlgUrgente
    'I.ECASTILLO 27.10.2020
    If .flgDiasFuturos <> "" Then chkDiasFuturos.Value = .flgDiasFuturos
    'F.ECASTILLO 27.10.2020
   
End If

    '*************************************************************
    '*************************************************************
    '*************************************************************
    '*************************************************************
    
End With
cargaEntrega
End Sub

Private Sub cmdAceptar_Click()
    If gstrIndRAv3 = "2" Then
        Dim flgCapacidadDisp As String
        flgCapacidadDisp = objVenta.validaCapacidad
        lblValTipoServicio.Visible = False
        If flgCapacidadDisp <> "1" Then
            lblValTipoServicio.Visible = True
            cmdEditMetodo_Click
            Exit Sub
        End If
    End If
    With objVenta
'''''''        .TipoCliente = ctlCboTipCliente.BoundText
'''''''    If .TipoCliente = "1" Then
'''''''        .Ruc = txtDato(0).Text
'''''''        .RazonSocial = txtDato(1).Text
'''''''        .DireccionClienteSocial = txtDato(2).Text
'''''''    Else
'''''''        .NumeroDocumentoID = txtDato(0).Text
'''''''        .NombreCliente = txtDato(1).Text
'''''''        .DireccionCliente = txtDato(2).Text
'''''''    End If
    
    
    If ctlCboTipCliente.BoundText = "1" Then
            objVenta.CodigoCliente = objVenta.Out_CodigoCliente
            If objCliente.fnValidaRUC(txtDato(0).Text) < 0 Then
                MsgBox "El RUC es invalido, verifique", vbInformation, App.ProductName
                Exit Sub
            End If
    End If
    
    
    
    If Not objUsuario.EsDelivery Then
            objVenta.TipoCliente = ctlCboTipCliente.BoundText
            If objVenta.TipoCliente = "1" Then
                objVenta.Ruc = txtDato(0).Text
                objVenta.RazonSocial = txtDato(1).Text
                objVenta.DireccionClienteSocial = txtDato(2).Text
            Else
                objVenta.NumeroDocumentoID = txtDato(0).Text
                objVenta.NombreCliente = txtDato(1).Text
                objVenta.DireccionCliente = txtDato(2).Text
            End If
            objVenta.NombreClienteDLV = txtDato(1).Text
            objVenta.DireccionClienteDLV = txtDato(2).Text
    End If
    
    

    ''objVenta.Out_NombreCliente = txtCliente.Text
    objVenta.Out_Tipo = ctlCboTipCliente.BoundText
    objVenta.Out_NumeroId = txtDato(0).Text
    objVenta.Out_NombreCliente = txtDato(1).Text
    objVenta.Out_Direccion = txtDato(2).Text
    
   ' Select Case SSTab1.Tab
    '    Case 0
    
            ''** NOTA  22/12/2009
            '    ----
            ''** Ojo aqui se tiene que condicionar para cuando la venta sea convenio y la emisión de los documentos
            ''** FAC - BOL se realizen en la caja 1 y se impriman en la maquina del químico cuando sucede eso en la
            ''** caja 1 no tiene asociada a una ticketera y le saldra un mensaje de que el documento a emitir no corresponde
            ''** y es ahi cuando se entra a la pantalla de documentos y seleeciona el doc BOL y es aqui donde limpia el valor
            ''** de la descripción de cliente y da como resultado que no salga en la impresión de la factura
    
            .DesAuxCliDirecc = txtDato(2).Text
            .DesAuxCliNombre = txtDato(1).Text
     '   Case 1
            .DesAuxRecogeNombre = ctlTextBox1.Text
            .DesAuxRecogeDirecc = ctlTextBox2.Text
            .DesAuxRecogeRef = ctlTextBox3.Text
            .DesAuxRecogeTlf = ctlTextBox4.Text
            .EntregaTercero = Check1.Value
             
            .UbigeoEntrega = cboDepartamento.BoundText & cboProvincia.BoundText & cboDistrito.BoundText
            .FchHoraPactEntr = Format(DTPicker1.Value, "dd/mm/yyyy ") & Format(DTPicker3.Value, "hh:mm:ss")
            .FchHoraPactRecog = Format(DTPicker2.Value, "dd/mm/yyyy ") & Format(DTPicker4.Value, "hh:mm:ss")
            .HoraPactEntr = Format(DTPicker3.Value, "hh:mm:ss") 'ECASTILLO 27.10.2020 - SE DESCOMENTA LINEA
            '.HoraPactRecog = Format(DTPicker4.Value, "hh:mm:ss")
            
            .FlgEntregaLocal = Check2.Value
            .FlgUrgente = Check3.Value
            .FlgPactado = chkFlgFchPactada.Value
            .flgDiasFuturos = chkDiasFuturos.Value
    'End Select
    
    .ObsNotaLocal = txtObservaciones(0).Text
    .ObsNotaRuteo = txtObservaciones(1).Text
    .ObsNotaMotorizado = txtObservaciones(2).Text
    .ObsNotaVerificacion = txtObservaciones(3).Text
    .CodigoDocumentoVenta = txtTipoDocumento.Text
    '.CodigoCliente = Trim(LblCodClientex.Caption)
    Dim objDocumento As New clsDocumento
    On Error GoTo handle
    frmPedido.lblSiguiente.Caption = grdFormaPago.Columns(0).Value & " - " & objDocumento.ListaNumeroDisponible(objUsuario.CodigoEmpresa, objUsuario.NombrePC, grdFormaPago.Columns(0).Value)
    
    '**** En Caso de Delivery ****'
    If objUsuario.EsDelivery Then
        strDlvDocumento = "Usted Efectuara el Pago con" & "  " & grdFormaPago.Columns(1).Value
    Else
        .CodigoCliente = txtCodigoCliente.Text
    End If
    '*****************************'
    'blnEditCli = False
    
    'If Not objUsuario.EsDelivery Then
        If (grdFormaPago.Columns(0).Value = objVenta.TipoDocTKB Or grdFormaPago.Columns(0).Value = objVenta.TipoDocTKF) Then
            .SerieTKT = objMaquina.Serie_Ticketera(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdFormaPago.Columns(0).Value, objUsuario.NombrePC)
        End If
    'End If
    
    Set objDocumento = Nothing
    Set objMaquina = Nothing
    
    End With

    Unload Me
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub ctlCboTipCliente_Change()
    If ctlCboTipCliente.BoundText = "1" Then
        lblTitle(1).Caption = "RUC"
        lblTitle(2).Caption = "Razón Social"
        lblTitle(3).Caption = "Dirección"

        '''''''''''''''''''''''''
'''''''        fraPJ.Visible = True
'''''''        fraPN.Visible = False
'''''''        txtRUC.TabIndex = 2
'''''''        cmdBuscarRUC.TabIndex = 3
'''''''        txtRazonSocial.TabIndex = 4
'''''''        txtDireccionPJ.TabIndex = 5
    Else
        lblTitle(1).Caption = "D.N.I."
        lblTitle(2).Caption = "Nombre y Apellido"
        lblTitle(3).Caption = "Dirección"
        Debug.Print ctlCboTipCliente.BoundText
'''''''        fraPJ.Visible = False
'''''''        fraPN.Visible = True
'''''''        txtDNI.TabIndex = 2
'''''''        cmdBuscarDNI.TabIndex = 3
'''''''        txtNombres.TabIndex = 4
'''''''        txtApellidos.TabIndex = 5
'''''''        txtDireccionPN.TabIndex = 6
    End If
End Sub

'-- Persona Juridica --'
Private Sub cmdBuscarRUC_Click()
With frm_VTA_Cliente_Bus
    .Show vbModal
    ctlCboTipCliente.BoundText = objVenta.Out_Tipo
    txtCliente.Text = objVenta.Out_NombreCliente
'    LblCodClientex.Caption = objVenta.CodigoCliente
    If Not .Out_NombreCliente = "" Then
        txtDato(0).Text = .Out_NumeroId
        txtDato(1).Text = .Out_NombreCliente
        txtDato(2).Text = .Out_Direccion
        'pherrera 021107 estaba asi:
        'txtCodigoCliente.Text = objVenta.Out_CodigoCliente
        txtCodigoCliente.Text = .Out_CodigoCliente
    End If
End With
'''''    If ctlCboTipCliente.BoundText = "*" Then MsgBox "Seleccione un tipo de Cliente", vbCritical, Caption: Exit Sub
'''''    Set frm_VTA_ClienteDatos.GrdBusCliente.DataSource = objCliente.ListaClientesGen(txtRUC.Text, ctlCboTipCliente.BoundText)
'''''    If frm_VTA_ClienteDatos.GrdBusCliente.DataSource.RecordCount = 0 Then
'''''        cmdMasDatos_Click
'''''    Else
'''''        frm_VTA_ClienteDatos.Pantalla = 0
'''''        frm_VTA_ClienteDatos.Show vbModal
'''''    End If
    
End Sub

'-- Persona Natural --'
Private Sub cmdBuscarDNI_Click()
'frm_VTA_Cliente_Bus.Show vbModal
 'txtApellidos.Text = frm_VTA_Cliente_Bus.Out_NombreCliente

''''''    If ctlCboTipCliente.BoundText = "*" Then MsgBox "Seleccione un tipo de Cliente", vbCritical, Caption: Exit Sub
''''''    Set frm_VTA_ClienteDatos.GrdBusCliente.DataSource = objCliente.ListaClientesGen(txtDNI.Text, ctlCboTipCliente.BoundText)
''''''    If frm_VTA_ClienteDatos.GrdBusCliente.DataSource.RecordCount = 0 Then
''''''        cmdMasDatos_Click
''''''    Else
''''''        frm_VTA_ClienteDatos.Pantalla = 0
''''''        frm_VTA_ClienteDatos.Show vbModal
''''''    End If
    
End Sub

Private Sub cmdMasDatos_Click()
    frm_VTA_Cliente.strCodigo = txtCodigoCliente
    frm_VTA_Cliente.CargarValores
    frm_VTA_Cliente.Show vbModal
End Sub

Private Sub grdDirecciones_Click()
    ctlTextBox2.Text = "" & grdDirecciones.Columns("DES_DIRECCION").Value
    txtCodDireccion.Text = "" & grdDirecciones.Columns("COD_DIRECCION_CLI").Value
    
    grdDirecciones.Visible = False
On Error GoTo Y
        Dim strUbigeo As String
    strUbigeo = "" & grdDirecciones.Columns("COD_UBIGEO").Value
    If Not strUbigeo = "" Then
    On Error GoTo Y
        cboDepartamento.BoundText = Mid(strUbigeo, 1, 2)
        cboProvincia.BoundText = Mid(strUbigeo, 3, 2)
        cboDistrito.BoundText = Mid(strUbigeo, 5, 2)
Y:
    End If
End Sub

Private Sub grdDirecciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then grdDirecciones_Click
End Sub

Private Sub grdFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Control

   Dim ShiftDown, AltDown, CtrlDown
   
   
   ShiftDown = (Shift And vbShiftMask) > 0
   AltDown = (Shift And vbAltMask) > 0
   CtrlDown = (Shift And vbCtrlMask) > 0
'''''''''    Select Case KeyCode
'''''''''        Case vbKeyReturn
'''''''''
''''''''''''''''             If ctlCboTipCliente.Enabled Then
''''''''''''''''                ctlCboTipCliente.SetFocus
'''''''''''''''''             Else
'''''''''''''''''                If txtDNI.Enabled Then txtDNI.SetFocus
''''''''''''''''             End If
'''''''''    End Select
    
    
    On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then
                cmdAceptar_Click
            Else
                cmdBuscarRUC.SetFocus
            End If
        Case vbKeyF1
            grdFormaPago.SetFocus
        Case vbKeyF2
            ctlCboTipCliente.SetFocus
        Case vbKeyF9
            Call cmdMasDatos_Click
        Case vbKeyF10
            SSTab1.Tab = 0
            cmdBuscarRUC.SetFocus
        Case vbKeyF11
            SSTab1.Tab = 1
            Check3.SetFocus
        Case vbKeyF12
           SSTab1.Tab = 2
        Case vbKeyEscape
            Unload Me
    End Select

    
   Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName


End Sub

Private Sub grdFormaPago_RegistroSeleccionado(ByVal DatoColumna0 As String)
    txtTipoDocumento.Text = "" & grdFormaPago.Columns(0).Value
    
    
'    If grdFormaPago.Row = 1 Then
'        blnTipoDoc = True
'    End If
    
    
    If (grdFormaPago.Columns(0).Value = "FAC") Or (grdFormaPago.Columns(0).Value = "TKF") Then
        ctlCboTipCliente.BoundText = objCliente.TipoClienteJuridico
        ctlCboTipCliente.Enabled = False
        If objUsuario.CodigoLocal = "DLV" Then cboRazonSocial.Visible = True
    Else
        ctlCboTipCliente.BoundText = objCliente.TipoClienteNatural
        ctlCboTipCliente.Enabled = True
         cboRazonSocial.Visible = False
    End If
    objVenta.Out_Tipo = ctlCboTipCliente.BoundText
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyF1
            grdFormaPago.SetFocus
        Case vbKeyF2
            If ctlCboTipCliente.Enabled = True Then
                ctlCboTipCliente.SetFocus
            Else
                txtDato(0).SetFocus
            End If
        Case vbKeyF9
            Call cmdMasDatos_Click
        Case vbKeyF10
            SSTab1.Tab = 0
            cmdBuscarRUC.SetFocus
        Case vbKeyF11
            SSTab1.Tab = 1
            Check3.SetFocus
        Case vbKeyF12
           SSTab1.Tab = 2
        Case vbKeyEscape
            Unload Me
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

'    Select Case KeyCode
'

'   Dim ShiftDown, AltDown, CtrlDown
   'ShiftDown = (Shift And vbShiftMask) > 0
'   AltDown = (Shift And vbAltMask) > 0
'   CtrlDown = (Shift And vbCtrlMask) > 0
    
   'psub_KeyDownAplicacion KeyCode, Shift
    
 '   Select Case KeyCode
 '       Case vbKeyReturn And Shift 'ShiftDown
 '               cmdAceptar_Click
 '       Case vbKeyF1
 '           grdFormaPago.SetFocus
 '   End Select
End Sub

Private Sub cmdCancelar_Click()
'    blnEditCli = False
    Unload Me
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        
''        If ctlCboTipCliente.BoundText = "1" Then
''            ctlTextBox1.Text = txtRazonSocial.Text
''            ctlTextBox2.Text = txtDireccionPJ.Text
''        Else
''            ctlTextBox1.Text = txtNombres.Text & " " & txtApellidos.Text
''            ctlTextBox2.Text = txtDireccionPN.Text
''        End If
    Else
        
    End If
End Sub


Private Sub SetGrd()

    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("CODIGO", "DESCRIPCION")
    arrCaption = Array("Código", "Descripción")
    arrAncho = Array(1000, 3000)
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft)
    grdFormaPago.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

End Sub

Private Sub txtDato_GotFocus(Index As Integer)
    'Cambio 09/10/2007 Por Cristhian Rueda'
    pblnEditCli = True
End Sub

Private Sub txtObservaciones_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Sub cargaEntrega()
On Error GoTo handle
    With objVenta
        
        'SE COMENTO POR QUE LOS DATOS DEL CLIENTE SE CARGAN EN EL LOAD
        'COMENTADO POR JLOPEZ EL 15/11/2007
'        If pblnEditCli = True Then txtDato(2).Text = .DesAuxCliDirecc
'        txtDato(1).Text = .DesAuxCliNombre
        ctlTextBox1.Text = .DesAuxRecogeNombre
        ctlTextBox2.Text = .DesAuxRecogeDirecc
        ctlTextBox3.Text = .DesAuxRecogeRef
        ctlTextBox4.Text = .DesAuxRecogeTlf
        cboDepartamento.BoundText = Mid(.UbigeoEntrega, 1, 2)
        cboProvincia.BoundText = Mid(.UbigeoEntrega, 3, 2)
        cboDistrito.BoundText = Mid(.UbigeoEntrega, 5, 2)
        'DTPicker1.Value = Format(.FchHoraPactEntr, "dd/mm/yyyy")
        'DTPicker2.Value = Format(.FchHoraPactEntr, "hh:mm:ss")
        
        If .FchHoraPactEntr <> "" Then DTPicker1.Value = Format(.FchHoraPactEntr, "dd/mm/yyyy")
        If .FchHoraPactRecog <> "" Then DTPicker2.Value = Format(.FchHoraPactRecog, "dd/mm/yyyy") ' Format(.FchHoraPactRecog, "hh:mm:ss AMPM")
        If .HoraPactEntr <> "" Then DTPicker3.Value = Format(.HoraPactEntr, "hh:mm:ss AMPM")
        If .HoraPactRecog <> "" Then DTPicker4.Value = Format(.HoraPactRecog, "hh:mm:ss AMPM")
        
        Check2.Value = IIf(.FlgEntregaLocal = "", 0, .FlgEntregaLocal)
        Check3.Value = IIf(.FlgUrgente = "", 0, .FlgUrgente)
        Check1.Value = IIf(.EntregaTercero = "", 0, .EntregaTercero)
        chkFlgFchPactada.Value = IIf(.FlgPactado = "", 0, .FlgPactado)
    End With
    Exit Sub
handle:

End Sub
