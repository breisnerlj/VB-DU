VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmPedido_Ingre_Cli 
   Caption         =   "Datos del Cliente"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin vbp_Ventas.ctlDataCombo cboDepartamento 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   4260
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin MSMask.MaskEdBox mskFechaNac 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   2220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5040
      TabIndex        =   19
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   7320
      Width           =   1335
   End
   Begin VB.ComboBox Cbo_Tipo_Doc 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1935
   End
   Begin vbp_Ventas.ctlTextBox ctlTxtDNI 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   540
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Tipo            =   3
      Enabled         =   0   'False
      MaxLength       =   15
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
   Begin vbp_Ventas.ctlTextBox ctlTxtAPaterno 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1380
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlTextBox ctlTxtAMaterno 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlTextBox ctlTxtNombres 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      Tipo            =   2
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1920
      TabIndex        =   27
      Top             =   2580
      Width           =   2775
      Begin VB.OptionButton optFemenino 
         Caption         =   "Femenino"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optMasculino 
         Caption         =   "Masculino"
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
   End
   Begin vbp_Ventas.ctlTextBox ctlTxtDireccion 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   3840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlTextBox ctlTxtReferencias 
      Height          =   315
      Left            =   1920
      TabIndex        =   14
      Top             =   5940
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlTextBox ctlTxtEmail 
      Height          =   315
      Left            =   1920
      TabIndex        =   15
      Top             =   6360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlTextBox ctlTxtTelFijo 
      Height          =   315
      Left            =   1920
      TabIndex        =   16
      Top             =   6780
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlTextBox ctlTxtTelMovil 
      Height          =   315
      Left            =   1920
      TabIndex        =   17
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlDataCombo cboProvincia 
      Height          =   315
      Left            =   1920
      TabIndex        =   11
      Top             =   4680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlDataCombo cboDistrito 
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   5100
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlDataCombo cboTipoDireccion 
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   5520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlDataCombo ctlCboSuFijoDirecc 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   3420
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label Label002 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   56
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label011 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   55
      Top             =   3060
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label006 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   54
      Top             =   6840
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label010 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   53
      Top             =   6420
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label012 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label007 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   51
      Top             =   3900
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label005 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   50
      Top             =   1860
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono Móvil:"
      Height          =   195
      Left            =   480
      TabIndex        =   49
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono Fijo:"
      Height          =   195
      Left            =   480
      TabIndex        =   48
      Top             =   6840
      Width           =   960
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencias:"
      Height          =   195
      Left            =   480
      TabIndex        =   47
      Top             =   6000
      Width           =   900
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   195
      Left            =   480
      TabIndex        =   46
      Top             =   6420
      Width           =   480
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Dirección:"
      Height          =   195
      Left            =   480
      TabIndex        =   45
      Top             =   5580
      Width           =   1080
   End
   Begin VB.Label Label015 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   44
      Top             =   5580
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label016 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   43
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento:"
      Height          =   195
      Left            =   480
      TabIndex        =   42
      Top             =   4320
      Width           =   1050
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distrito:"
      Height          =   195
      Left            =   480
      TabIndex        =   41
      Top             =   5160
      Width           =   525
   End
   Begin VB.Label Label014 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   40
      Top             =   5160
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Lugar:"
      Height          =   195
      Left            =   480
      TabIndex        =   39
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   480
      TabIndex        =   38
      Top             =   3900
      Width           =   720
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sexo:"
      Height          =   195
      Left            =   480
      TabIndex        =   37
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia:"
      Height          =   195
      Left            =   480
      TabIndex        =   36
      Top             =   4740
      Width           =   705
   End
   Begin VB.Label Label008 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   35
      Top             =   2700
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label017 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   3480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label013 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   33
      Top             =   4740
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(*) Campos Obligatorios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2400
      TabIndex        =   32
      Top             =   7440
      Width           =   1995
   End
   Begin VB.Label Label009 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   2280
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label004 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   30
      Top             =   1440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label003 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   1020
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label LblSaldo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   540
      Width           =   975
   End
   Begin VB.Label LblTSaldo 
      Caption         =   "Saldo (S/.):"
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Nacimiento:"
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres:"
      Height          =   195
      Left            =   480
      TabIndex        =   25
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido Materno:"
      Height          =   195
      Left            =   480
      TabIndex        =   24
      Top             =   1860
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido Paterno:"
      Height          =   195
      Left            =   480
      TabIndex        =   23
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Documento:"
      Height          =   195
      Left            =   480
      TabIndex        =   22
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Documento:"
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Top             =   180
      Width           =   1230
   End
End
Attribute VB_Name = "FrmPedido_Ingre_Cli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objClienteD As New clsClienteD
Dim objVenta As New clsVenta
Dim objCliente As New clsCliente
Public b_monedero As Boolean
Public b_afiliar As Boolean

Private Sub cboDepartamento_Change()
    Set cboProvincia.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, cboDepartamento.BoundText, "[ SELECCIONAR ]")
    cboProvincia.ListField = "Descripcion"
    cboProvincia.BoundColumn = "Codigo"
    cboProvincia.BoundText = "*"
End Sub

Private Sub cboProvincia_Change()
    Set cboDistrito.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, cboDepartamento.BoundText, cboProvincia.BoundText, "[ SELECCIONAR ]")
    cboDistrito.ListField = "Descripcion"
    cboDistrito.BoundColumn = "Codigo"
    cboDistrito.BoundText = "*"
End Sub

Private Sub Form_Activate()
    If Not b_monedero Then
        If ctlTxtNombres.Text <> "" Then
            If ctlTxtAPaterno.Text = "" Then
                ctlTxtAPaterno.SetFocus
                Exit Sub
            End If
            If ctlTxtAMaterno.Text = "" Then
                ctlTxtAMaterno.SetFocus
                Exit Sub
            End If
            If mskFechaNac.Text = "__/__/____" Then
                mskFechaNac.SetFocus
                Exit Sub
            End If
        End If
    Else
        If Label002.Visible And Trim(ctlTxtDNI.Text) = "" Then
            ctlTxtDNI.SetFocus
            Exit Sub
        End If
                
        If Label003.Visible And Trim(ctlTxtNombres.Text) = "" Then
            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        If Label004.Visible And Trim(ctlTxtAPaterno.Text) = "" Then
            ctlTxtAPaterno.SetFocus
            Exit Sub
        End If
        
        If Label005.Visible And Trim(ctlTxtAMaterno.Text) = "" Then
            ctlTxtAMaterno.SetFocus
            Exit Sub
        End If
        
        If Label009.Visible And Trim(mskFechaNac.Text) = "__/__/____" Then
            mskFechaNac.SetFocus
            Exit Sub
        End If
        
        If Label008.Visible And Not optFemenino.Value And Not optMasculino.Value Then
            optFemenino.SetFocus
            Exit Sub
        End If
        
        If Label017.Visible And Trim(ctlCboSuFijoDirecc.BoundText) = "" Then
            ctlCboSuFijoDirecc.SetFocus
            Exit Sub
        End If
        
        If Label007.Visible And Trim(ctlTxtDireccion.Text) = "" Then
            ctlTxtDireccion.SetFocus
            Exit Sub
        End If
        
        If Label012.Visible And Trim(cboDepartamento.BoundText) = "" Then
            cboDepartamento.SetFocus
            Exit Sub
        End If
        
        If Label013.Visible And Trim(cboProvincia.BoundText) = "" Then
            cboProvincia.SetFocus
            Exit Sub
        End If
        
        If Label014.Visible And Trim(cboDistrito.BoundText) = "" Then
            cboDistrito.SetFocus
            Exit Sub
        End If
        
        If Label015.Visible And Trim(cboTipoDireccion.BoundText) = "" Then
            cboTipoDireccion.SetFocus
            Exit Sub
        End If
        
        If Label016.Visible And Trim(ctlTxtReferencias.Text) = "" Then
            ctlTxtReferencias.SetFocus
            Exit Sub
        End If
        
        If Label010.Visible And Trim(ctlTxtEmail.Text) = "" Then
            ctlTxtEmail.SetFocus
            Exit Sub
        End If
        
        If Label006.Visible And Trim(ctlTxtTelFijo.Text) = "" Then
            ctlTxtTelFijo.SetFocus
            Exit Sub
        End If
        
        If Label011.Visible And Trim(ctlTxtTelMovil.Text) = "" Then
            ctlTxtTelMovil.SetFocus
            Exit Sub
        End If
    
    End If
End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

    b_monedero = False
    b_afiliar = True
    
    Cbo_Tipo_Doc.AddItem frmPedido_Busca_Cli.CboTipoDoc.Text
    'Cbo_Tipo_Doc.ListIndex = frmPedido_Busca_Cli.CboTipoDoc.BoundText
    Cbo_Tipo_Doc.ListIndex = 0
    ctlTxtDNI.Text = frmPedido_Busca_Cli.ctlTxtDNI.Text

    mskFechaNac.Format = "dd/mm/yyyy"
    mskFechaNac.Mask = "##/##/####"
    mskFechaNac.Text = Date

    Set cboDepartamento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0, "[ SELECCIONAR ]")
    cboDepartamento.ListField = "Descripcion"
    cboDepartamento.BoundColumn = "Codigo"
    cboDepartamento.BoundText = "*"

    Set cboTipoDireccion.RowSource = objCliente.ListaTipoDireccionCEN
    cboTipoDireccion.ListField = "DES_TIPO_DIRECCION"
    cboTipoDireccion.BoundColumn = "COD_TIPO_DIRECCION"

    Set ctlCboSuFijoDirecc.RowSource = objCliente.SuFijoDirecc
    ctlCboSuFijoDirecc.ListField = "DES_ABREVIATURA_DIRECCION"
    ctlCboSuFijoDirecc.BoundColumn = "COD_SUFIJO_DIRECCION"
    ctlCboSuFijoDirecc.ListField2 = "DES_SUFIJO_DIRECCION"
    
    'visualiza el saldo
    Dim rs As oraDynaset
    Dim v_Bool As Boolean
    v_Bool = objVenta.MuestraFidelizado("FLGSALDO")
    LblTSaldo.Visible = v_Bool
    lblSaldo.Visible = v_Bool

    CargarCliente

    'cargar saldo
    lblSaldo.Caption = Format(objVenta.Saldo_Cliente(objVenta.CodigoCliente), "##0.00")

End Sub

Private Sub ReDibujarForm()
    Label017.Visible = False 'b_monedero
    Label17.Visible = False 'b_monedero
    ctlCboSuFijoDirecc.Visible = False 'b_monedero

    Label007.Visible = False 'b_monedero
    Label16.Visible = False 'b_monedero
    ctlTxtDireccion.Visible = False 'b_monedero

    Label012.Visible = False 'b_monedero
    Label20.Visible = False 'b_monedero
    cboDepartamento.Visible = False 'b_monedero

    Label013.Visible = False 'b_monedero
    Label14.Visible = False 'b_monedero
    cboProvincia.Visible = False 'b_monedero

    Label014.Visible = False 'b_monedero
    Label19.Visible = False 'b_monedero
    cboDistrito.Visible = False 'b_monedero

    Label015.Visible = False 'b_monedero
    Label23.Visible = False 'b_monedero
    cboTipoDireccion.Visible = False 'b_monedero

    Label016.Visible = False 'b_monedero
    Label25.Visible = False 'b_monedero
    ctlTxtReferencias.Visible = False 'b_monedero

    Label010.Visible = False 'b_monedero
    Label24.Visible = False 'b_monedero
    ctlTxtEmail.Visible = False 'b_monedero

    Label006.Visible = False 'b_monedero
    Label26.Visible = False 'b_monedero
    ctlTxtTelFijo.Visible = False 'b_monedero
    
    Label011.Visible = b_monedero
    Label27.Visible = b_monedero
    ctlTxtTelMovil.Visible = b_monedero
    
    If Not b_monedero Then
        cmdAceptar.top = 3260
        CmdCancelar.top = 3260
        Label10.top = 3380
        Me.Height = 4280
    Else
        cmdAceptar.top = 3575 '7560
        CmdCancelar.top = 3575 '7560
        Label10.top = 3695 '7680
        Me.Height = 4595 '8580
    End If
    
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub CargarClavesCampos()
    Dim rs As oraDynaset
    Dim v_Cod As String, v_Obl As String
    
    Set rs = gclsOracle.FN_Cursor("BTLPROD.PKG_MONEDERO_MF.FN_LISTA_CAMPOS", 0)
    
    Do Until rs.EOF
        v_Cod = "" & rs("COD_CAMPO").Value
        v_Obl = "" & rs("IND_OBLIGATORIO").Value
        Select Case v_Cod
            Case "002" 'DNI_CLIENTE
                Label002.Visible = IIf(v_Obl = "S", True, False)
            Case "003" 'NOMBRE_CLIENTE
                Label003.Visible = IIf(v_Obl = "S", True, False)
            Case "004" 'APEPAT_CLIENTE
                Label004.Visible = IIf(v_Obl = "S", True, False)
            Case "005" 'APEMAT_CLIENTE
                Label005.Visible = IIf(v_Obl = "S", True, False)
            Case "009" 'FECHA_NAC_CLIENTE
                Label009.Visible = IIf(v_Obl = "S", True, False)
            Case "008" 'SEXO_CLIENTE
                Label008.Visible = IIf(v_Obl = "S", True, False)
            Case "017" 'TIPO LUGAR
                Label017.Visible = IIf(v_Obl = "S", True, False)
            Case "007" 'DIREC_CLIENTE
                Label007.Visible = IIf(v_Obl = "S", True, False)
            Case "012" 'DEPARTAMENTO
                Label012.Visible = IIf(v_Obl = "S", True, False)
            Case "013" 'PROVINCIA
                Label013.Visible = IIf(v_Obl = "S", True, False)
            Case "014" 'DISTRITO
                Label014.Visible = IIf(v_Obl = "S", True, False)
            Case "015" 'TIPO DIRECCION
                Label015.Visible = IIf(v_Obl = "S", True, False)
            Case "016" 'REFERENCIAS
                Label016.Visible = IIf(v_Obl = "S", True, False)
            Case "010" 'EMAIL_CLIENTE
                Label010.Visible = IIf(v_Obl = "S", True, False)
            Case "006" 'TELEFONO_CLIENTE
                Label006.Visible = IIf(v_Obl = "S", True, False)
            Case "011" 'CELULAR_CLIENTE
                Label011.Visible = IIf(v_Obl = "S", True, False)
        End Select
        rs.MoveNext
    Loop
    
End Sub

Public Sub carga(TipoDoc As String, NumDoc As String)
    
    Cbo_Tipo_Doc.Text = TipoDoc
    ctlTxtDNI.Text = NumDoc
    
    ReDibujarForm
    If b_monedero Then CargarClavesCampos

    CargarCliente
    
    Me.Show vbModal

End Sub

Private Sub cmdCancelar_Click()
    frmPedido_Busca_Cli.v_delfrm = ""
    Unload Me
End Sub

Public Sub cmdAceptar_Click()
On Error GoTo CtrlErr
    'validar txts
    If Not b_monedero Then
        If Trim(ctlTxtNombres.Text) = "" Then
            MsgBox " Ingresar el Nombre completo, no puede ser vacio.", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal Else ctlTxtNombres.SetFocus
            Exit Sub
        End If
        If Trim(ctlTxtAPaterno.Text) = "" Then
            MsgBox " Ingresar el Apellido Paterno, no puede ser vacio. ", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal Else ctlTxtAPaterno.SetFocus
            Exit Sub
        End If
        If Trim(ctlTxtAMaterno.Text) = "" Then
            MsgBox " Ingresar el Apellido Materno, no puede ser vacio. ", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal Else ctlTxtAMaterno.SetFocus
            Exit Sub
        End If
        If Not IsDate(mskFechaNac.Text) Then
            MsgBox "Ingresar una fecha correcta por favor", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal Else mskFechaNac.SetFocus
            'Cancel = True
            Exit Sub
        End If
        If Trim(mskFechaNac.Text) = "__/__/____" Then   '"01/01/1900" Then
            MsgBox "Ingresar la Fecha de Nacimiento correcta, y verificar el Sexo del Cliente", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal Else mskFechaNac.SetFocus
            Exit Sub
        End If
    Else
        'DNI_CLIENTE
        If Label002.Visible And Trim(ctlTxtDNI.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
                
        'NOMBRE_CLIENTE
        If Label003.Visible And Trim(ctlTxtNombres.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'APEPAT_CLIENTE
        If Label004.Visible And Trim(ctlTxtAPaterno.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'APEMAT_CLIENTE
        If Label005.Visible And Trim(ctlTxtAMaterno.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'FECHA_NAC_CLIENTE
        If Label009.Visible Then
            If Trim(mskFechaNac.Text) = "__/__/____" Then
                MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
                If Not Me.Visible Then Me.Show vbModal
'                ctlTxtDNI.SetFocus
                Exit Sub
            End If
            If Not IsDate(mskFechaNac.Text) Then
                MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
                If Not Me.Visible Then Me.Show vbModal
'                ctlTxtDNI.SetFocus
                Exit Sub
            End If
        End If
        
        'SEXO_CLIENTE
        If Label008.Visible And Trim(ctlTxtDNI.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'TIPO LUGAR
        If Label017.Visible And Trim(ctlCboSuFijoDirecc.BoundText) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'DIREC_CLIENTE
        If Label007.Visible And Trim(ctlTxtDireccion.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'DEPARTAMENTO
        If Label012.Visible And Trim(cboDepartamento.BoundText) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'PROVINCIA
        If Label013.Visible And Trim(cboProvincia.BoundText) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'DISTRITO
        If Label014.Visible And Trim(cboDistrito.BoundText) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'TIPO DIRECCION
        If Label015.Visible And Trim(cboTipoDireccion.BoundText) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'REFERENCIAS
        If Label016.Visible And Trim(ctlTxtReferencias.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'EMAIL_CLIENTE
        If Label010.Visible And Trim(ctlTxtEmail.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'TELEFONO_CLIENTE
        If Label006.Visible And Trim(ctlTxtTelFijo.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
        
        'CELULAR_CLIENTE
        If Label011.Visible And Trim(ctlTxtTelMovil.Text) = "" Then
            MsgBox "Falta completar un dato obligatorio", vbCritical + vbOKOnly, App.ProductName
            If Not Me.Visible Then Me.Show vbModal
'            ctlTxtDNI.SetFocus
            Exit Sub
        End If
    
    End If
    
    Grabar_Cliente

'''    MsgBox "Bienvenido, " & ctlTxtNombres.Text & ", " & ctlTxtAPaterno.Text & " " & ctlTxtAMaterno.Text, vbInformation, "Actualización"
'''    frmPedido.lbl_Cliente.Caption = ctlTxtNombres.Text & ", " & ctlTxtAPaterno.Text & " " & ctlTxtAMaterno.Text
'''    frmPedido.pstrDniCli = Trim(Me.ctlTxtDNI.Text)
'''    frmPedido.pstrNomcli = frmPedido.lbl_Cliente.Caption
'''    frmPedido.loadOptions
    Unload Me
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
   
End Sub

Private Sub ctlTxtNombres_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ctlTxtAPaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ctlTxtAMaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub optFemenino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub optMasculino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub mskFechaNac_KeyPress(KeyAscii As Integer)
Dim dfecha As Date
On Error GoTo CtrlErr
    If KeyAscii = 13 Then
        If mskFechaNac.Text = "__/__/____" Then
            'MsgBox "La fecha es incorrecta", vbCritical + vbOKOnly, App.ProductName
            'mskFechaNac.Text = "01/01/1900"
        Else
            If Not IsDate(mskFechaNac.Text) Then
                MsgBox "Ingresar una fecha correcta por favor", vbCritical + vbOKOnly, App.ProductName
                KeyAscii = 0
            End If
        End If
    End If
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub mskFechaNac_Validate(Cancel As Boolean)
On Error GoTo CtrlErr
'    If Not IsDate(mskFechaNac.Text) Then
'        mskFechaNac.Text = "__/__/____"
'        MsgBox "Ingresar una fecha correcta por favor", vbCritical + vbOKOnly, App.ProductName
'        Cancel = True
'    Else
'        CmdAceptar.SetFocus
'    End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
   
End Sub

Public Sub CargarCliente()
    Dim vSex As String
    Dim rs As oraDynaset
    
    On Error GoTo CtrlErr
    
    'If Not b_monedero Then
    ReDibujarForm
    If b_monedero Then
        CargarClavesCampos
        Set rs = objVenta.fnListaAfiliadoMonedero(Mid(Cbo_Tipo_Doc.Text, 1, 3), ctlTxtDNI.Text)
        ctlTxtNombres.Text = "" & rs("DES_NOM_CLIENTE").Value
        ctlTxtAPaterno.Text = "" & rs("DES_APE_CLIENTE").Value
        ctlTxtAMaterno.Text = "" & rs("DES_APE2_CLIENTE").Value
        mskFechaNac.Text = IIf(IsNull(rs("FCH_NACIMIENTO").Value), "__/__/____", rs("FCH_NACIMIENTO").Value)
        vSex = IIf(IsNull(rs("FLG_SEXO").Value), "0", rs("FLG_SEXO").Value)
        optFemenino.Value = IIf(vSex = "0", True, False)
        optMasculino.Value = IIf(vSex = "1", True, False)
        ctlTxtTelMovil.Text = "" & rs("NUM_TEL_MOVIL").Value
        objVenta.CodigoCliente = "" & rs("COD_CLIENTE").Value
        Set rs = Nothing
    Else
        Set rs = objClienteD.ListaCliente(Mid(Cbo_Tipo_Doc.Text, 1, 3), ctlTxtDNI.Text, gstrCodTarjetaFid)
        ctlTxtNombres.Text = IIf(IsNull(rs("DES_NOM_CLIENTE").Value), "", rs("DES_NOM_CLIENTE").Value)
        ctlTxtAPaterno.Text = IIf(IsNull(rs("DES_APE_CLIENTE").Value), "", rs("DES_APE_CLIENTE").Value)
        ctlTxtAMaterno.Text = IIf(IsNull(rs("DES_APE2_CLIENTE").Value), "", rs("DES_APE2_CLIENTE").Value)
        mskFechaNac.Text = IIf(IsNull(rs("FCH_NACIMIENTO").Value), "__/__/____", rs("FCH_NACIMIENTO").Value)
        vSex = IIf(IsNull(rs("FLG_SEXO").Value), "0", rs("FLG_SEXO").Value)
        If vSex = "0" Then
            optFemenino.Value = True
        Else
            optMasculino.Value = True
        End If
        objVenta.CodigoCliente = "" & rs("COD_CLIENTE").Value
        Set rs = Nothing
    End If
    'si lo encuentra
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Public Sub Grabar_Cliente()
    Dim strMensaje As String
    Dim v_Sex As String, v_Ubi As String, v_monedero As Boolean
    Dim rs As oraDynaset
    Dim oDP As New clsDocumentoPago
On Error GoTo CtrlErr
    
    v_Sex = IIf(optFemenino.Value, "0", "1")
    v_Ubi = cboDepartamento.BoundText & cboProvincia.BoundText & cboDistrito.BoundText
    v_Ubi = IIf(Len(v_Ubi) < 6, "", v_Ubi)
    
    If b_monedero Then
        strMensaje = objVenta.fnGrabaAfiliadoMonedero(objVenta.CodigoCliente, _
                                                      Mid(Cbo_Tipo_Doc.Text, 1, 3), _
                                                      ctlTxtDNI.Text, _
                                                      ctlTxtNombres.Text, _
                                                      ctlTxtAPaterno.Text, _
                                                      ctlTxtAMaterno.Text, _
                                                      ctlTxtEmail.Text, _
                                                      v_Sex, _
                                                      mskFechaNac.Text, _
                                                      v_Ubi, _
                                                      gstrCodTarjetaMon, _
                                                      ctlTxtTelFijo.Text, _
                                                      ctlTxtTelMovil.Text, _
                                                      ctlCboSuFijoDirecc.BoundText, _
                                                      ctlTxtDireccion.Text, _
                                                      cboTipoDireccion.BoundText, _
                                                      ctlTxtReferencias.Text, _
                                                      "N", _
                                                      IIf(b_monedero, "S", "N"))
    Else
        strMensaje = objVenta.fnGrabaFidelizado(objVenta.CodigoCliente, _
                                                Mid(Cbo_Tipo_Doc.Text, 1, 3), _
                                                ctlTxtDNI.Text, _
                                                ctlTxtNombres.Text, _
                                                ctlTxtAPaterno.Text, _
                                                ctlTxtAMaterno.Text, _
                                                "", _
                                                ctlTxtEmail.Text, _
                                                "", _
                                                v_Sex, _
                                                mskFechaNac.Text, _
                                                0, _
                                                v_Ubi, _
                                                objUsuario.CodigoLocal, _
                                                objUsuario.Codigo, _
                                                gstrCodTarjetaFid)
    End If
   
    If Mid(strMensaje, 1, 3) = "ORA" Then
        'MsgBox strMensaje, vbCritical, App.ProductName
        Err.Raise vbObjectError, "Grabar_Cliente", strMensaje
    Else
        objVenta.CodigoCliente = strMensaje
        frmPedido_Busca_Cli.v_delfrm = strMensaje
        frmPedido.lbl_Cliente.Caption = ctlTxtNombres.Text & ", " & _
                                        ctlTxtAPaterno.Text & " " & _
                                        ctlTxtAMaterno.Text
        If b_monedero And b_afiliar Then _
            ImprimeCuponAfiliacion Mid(Cbo_Tipo_Doc.Text, 1, 3), _
                                   ctlTxtDNI.Text

    End If
    
    Set rs = Nothing
Exit Sub
CtrlErr:
    Err.Raise Err.Number, "Grabar_Cliente", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub mskFechaNac_GotFocus()
    mskFechaNac.SelStart = 0
    mskFechaNac.SelLength = 10
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If optFemenino.Value Then
                optMasculino.Value = True
            Else
                optFemenino.Value = True
            End If
    End Select
End Sub

Public Sub ImprimeCuponAfiliacion(ByVal TipoDocumento As String, ByVal numDocumento As String)
    Dim rst As oraDynaset
    Dim Linea(1 To 11) As String
    Dim intConstante As Integer, intX As Integer
    
    On Error GoTo Control
    
    Screen.MousePointer = vbHourglass
    
    Set rst = objVenta.fnListaAfiliadoMonedero(TipoDocumento, numDocumento)
    Linea(1) = "" & rst("NUM_DOCUMENTO_ID").Value
    Linea(2) = "" & rst("DES_NOM_CLIENTE").Value & " " & _
                    rst("DES_APE_CLIENTE").Value & " " & _
                    rst("DES_APE2_CLIENTE").Value
    Linea(3) = "" & rst("FCH_NACIMIENTO").Value
    Linea(4) = "" & rst("FLG_SEXO").Value
    Linea(5) = "" & rst("DES_EMAIL").Value
    Linea(6) = "" & rst("NUM_TEL_MOVIL").Value
    Linea(7) = "" & rst("NUM_TEL_FIJO").Value
    Linea(8) = "" & rst("DES_DIRECCION_SOCIAL").Value
    Linea(9) = "" & rst("DES_REFERENCIA").Value
    Linea(10) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV687", objUsuario.CodigoEmpresa)
    Linea(11) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV506", objUsuario.CodigoEmpresa)
            
    intConstante = 184
    
    If SetearImpresoraCupon = False Then
        MsgBox "No tiene la impresora de cupones Instalada o tiene diferente nombre", vbCritical, App.ProductName
        Exit Sub
    End If
    For intX = 0 To Printer.FontCount - 1
        Debug.Print Printer.Fonts(intX)
    Next
    
    'Logo
'''    mdiPrincipal.imgLogoBtl.Picture = mdiPrincipal.ImageList1.ListImages(5).Picture
'''    Printer.PaintPicture mdiPrincipal.imgLogoBtl, 1100, 20, 2150, 750
'''    Printer.CurrentY = 1154
    
    'Nombre cliente
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.Print ""
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, Trim(Replace(UCase(objUsuario.Empresa), "Í", "I")) & " - RUC: " & objUsuario.Ruc
    Printer.Print ""
    Printer.Print ""
    
    Printer.FontName = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = False
    Printer.Print "DNI/CE: "; Tab(12); Linea(1)
    Printer.Print ""
    Printer.Print "Nombre: "; Tab(12); Linea(2)
    Printer.Print ""
    Printer.Print "Fec.Nac.: "; Tab(12); IIf(Linea(3) = "", String(15, "_ "), Linea(3)); _
                                Tab(40); "Sexo: " & IIf(Linea(4) = "", String(5, "_ "), IIf(Linea(4) = "1", "M", "F"))
    Printer.Print ""
    Printer.Print "EMail: "; Tab(12); IIf(Linea(5) = "", String(100, "_ "), Linea(5))
    Printer.Print ""
    Printer.Print "Celular: "; Tab(12); IIf(Linea(6) = "", String(100, "_ "), Linea(6))
    Printer.Print ""
    Printer.Print "Telf.Fijo: "; Tab(12); IIf(Linea(7) = "", String(100, "_ "), Linea(7))
    Printer.Print ""
    Printer.Print "Dirección: "; Tab(12); IIf(Linea(8) = "", String(100, "_ "), Linea(8))
    Printer.Print ""
    Printer.Print Tab(12); IIf(Linea(9) = "", String(100, "_ "), Linea(9))
    Printer.Print ""
    Printer.Print ""
    Printer.Print "Firma: "; Tab(12); String(100, "_ ")
    Printer.Print ""

    Printer.FontName = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = False
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, Linea(10)
    Printer.Print ""
    Printer.Print ""
    
    Printer.FontName = "IDAutomation.com Code39"  'LETRA DE CODIGO DE BARRA'
    Printer.FontSize = 18
    Printer.FontBold = False
    centra_printer "*" & Linea(1) & "*"
    
    Printer.FontName = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = False
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, Linea(11)
    Printer.Print ""
    Printer.Print ""
    
    Printer.FontBold = True
    Printer.Print "CodV: " & objUsuario.Codigo & " - CodL: " & objUsuario.CodigoLocal
    Printer.Print "Fecha Actual: " & Format$(Now, "dd/mm/yyyy ttttt")
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.EndDoc
   
    Set rst = Nothing
    
    Screen.MousePointer = vbDefault
    MsgBox "Recoger el voucher de la cuponera", vbInformation, App.ProductName
    Exit Sub
Control:
    Set rst = Nothing
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

