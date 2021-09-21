VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlCliente 
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9420
   Begin vbp_Ventas.ctlDataCombo cboDespacho 
      Height          =   315
      Left            =   6705
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlGrilla GrdBusCliente 
      Height          =   1575
      Left            =   2760
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame fraLabel 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   58
      Top             =   2530
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "F3 - Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   60
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblBusqueda 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "<Enter> - Seleccionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   1860
      End
   End
   Begin VB.Frame fraLabel 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   63
      Top             =   3210
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "F3 - Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   65
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblBusqueda 
         BackColor       =   &H80000018&
         Caption         =   "<Enter> - Seleccionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Width           =   1860
      End
   End
   Begin vbp_Ventas.ctlGrilla grdDirecciones 
      Height          =   1335
      Left            =   2280
      TabIndex        =   62
      Top             =   1920
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2355
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2880
      Left            =   120
      TabIndex        =   46
      Top             =   960
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5080
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dirección"
      TabPicture(0)   =   "ctlCliente.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Adicionales"
      TabPicture(1)   =   "ctlCliente.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
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
         Height          =   2445
         Left            =   240
         TabIndex        =   66
         Top             =   360
         Width           =   8895
         Begin VB.CheckBox chkCallao 
            Caption         =   "Callao"
            Height          =   255
            Left            =   7560
            TabIndex        =   77
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CheckBox chkAñadir 
            Caption         =   "Añadir dirección"
            Height          =   195
            Left            =   5280
            TabIndex        =   22
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CommandButton cmdAñadir 
            Caption         =   "Añadir "
            Height          =   615
            Left            =   7080
            Picture         =   "ctlCliente.ctx":0038
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   1680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chkPrincipal 
            Caption         =   "Principal"
            Height          =   255
            Left            =   5280
            TabIndex        =   21
            Top             =   1710
            Width           =   975
         End
         Begin VB.CommandButton cmdVerDireccion 
            Height          =   375
            Left            =   7650
            Picture         =   "ctlCliente.ctx":05C2
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   210
            Width           =   495
         End
         Begin vbp_Ventas.ctlDataCombo cboUrbanizacion 
            Height          =   315
            Left            =   5160
            TabIndex        =   15
            Top             =   600
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboPais 
            Height          =   315
            Left            =   1140
            TabIndex        =   17
            Top             =   1320
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlTextBox txtDireccion 
            Height          =   315
            Left            =   1890
            TabIndex        =   12
            Top             =   240
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   556
            Tipo            =   2
            MaxLength       =   200
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
            Left            =   1140
            TabIndex        =   19
            Top             =   1680
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboDepartamento 
            Height          =   315
            Left            =   5250
            TabIndex        =   18
            Top             =   1320
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboDistrito 
            Height          =   315
            Left            =   1140
            TabIndex        =   14
            Top             =   600
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlTextBox txtUrbanizacion 
            Height          =   255
            Left            =   6360
            TabIndex        =   67
            Top             =   1680
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
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
         Begin vbp_Ventas.ctlTextBox txtReferencia 
            Height          =   315
            Left            =   1140
            TabIndex        =   16
            Top             =   960
            Width           =   6990
            _ExtentX        =   12330
            _ExtentY        =   556
            Tipo            =   2
            MaxLength       =   80
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
         Begin vbp_Ventas.ctlDataCombo ctlCboSuFijoDirecc 
            Height          =   315
            Left            =   1140
            TabIndex        =   11
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboTipoDireccion 
            Height          =   315
            Left            =   1140
            TabIndex        =   20
            Top             =   2040
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1020
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Urbanización"
            Height          =   195
            Left            =   4200
            TabIndex        =   75
            Top             =   720
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   1740
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pais"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   1380
            Width           =   300
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   4185
            TabIndex        =   71
            Top             =   1380
            Width           =   1005
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   660
            Width           =   480
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   2100
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Adicionales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   47
         Top             =   480
         Width           =   8535
         Begin vbp_Ventas.ctlTextBox txtEmail 
            Height          =   375
            Left            =   1080
            TabIndex        =   29
            Top             =   1440
            Width           =   5295
            _ExtentX        =   9340
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
         Begin vbp_Ventas.ctlTextBox txtNumDocumento 
            Height          =   315
            Left            =   5640
            TabIndex        =   24
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
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
         Begin MSComCtl2.DTPicker dtpFechaNacimento 
            Height          =   300
            Left            =   5640
            TabIndex        =   28
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   64749569
            CurrentDate     =   39469
         End
         Begin vbp_Ventas.ctlDataCombo cboTipoDocumento 
            Height          =   315
            Left            =   1080
            TabIndex        =   23
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboEstadoCivil 
            Height          =   315
            Left            =   1080
            TabIndex        =   25
            Top             =   600
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlTextBox txtNumHijos 
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Tipo            =   3
            MaxLength       =   3
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
         Begin vbp_Ventas.ctlDataCombo cboSexo 
            Height          =   315
            Left            =   5640
            TabIndex        =   26
            Top             =   600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "N° Hijos"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sexo"
            Height          =   195
            Left            =   4920
            TabIndex        =   51
            Top             =   600
            Width           =   360
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Nac."
            Height          =   195
            Left            =   4920
            TabIndex        =   50
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "E-mail"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label Label21 
            Caption         =   "Número"
            Height          =   255
            Left            =   4920
            TabIndex        =   48
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin vbp_Ventas.ctlTextBox txtApeMaterno 
      Height          =   315
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Tipo            =   2
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   615
      Left            =   6600
      Picture         =   "ctlCliente.ctx":0B4C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Modificado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4920
      TabIndex        =   37
      Top             =   4665
      Width           =   4335
      Begin vbp_Ventas.ctlTextBox txtUsuarioReg 
         Height          =   255
         Left            =   975
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   450
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
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
      Begin vbp_Ventas.ctlTextBox txtFechaReg 
         Height          =   255
         Left            =   975
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   600
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   450
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
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
      Begin vbp_Ventas.ctlTextBox txtUsuarioAct 
         Height          =   255
         Left            =   960
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   960
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   450
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
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
      Begin vbp_Ventas.ctlTextBox txtFechaAct 
         Height          =   255
         Left            =   975
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1320
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   450
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
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
      Begin VB.Label Label19 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   75
         TabIndex        =   43
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Modificado"
         Height          =   255
         Left            =   75
         TabIndex        =   42
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   45
         TabIndex        =   39
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registrado"
         Height          =   195
         Left            =   45
         TabIndex        =   38
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Telef. Asociados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   4695
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Height          =   1335
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2355
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   17
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   84
         Columns(2)._MaxComboItems=   1
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Button=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=529"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=450"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=1402"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1323"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowAddNew     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   0
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "&Juridico"
      Height          =   240
      Index           =   1
      Left            =   8400
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "&Natural"
      Height          =   240
      Index           =   0
      Left            =   7560
      TabIndex        =   3
      Top             =   360
      Value           =   -1  'True
      Width           =   855
   End
   Begin vbp_Ventas.ctlTextBox txtNombre 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Tipo            =   2
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
   Begin vbp_Ventas.ctlDataCombo cboSufijo 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox txtApellido 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      Tipo            =   2
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
   Begin VB.TextBox txtObservacion 
      Height          =   615
      Left            =   1080
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   3960
      Width           =   8175
   End
   Begin vbp_Ventas.ctlDataCombo cboLocal 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox txtCodigo 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      ColorDefault    =   -2147483634
      ColorDefault    =   -2147483634
      Enabled         =   0   'False
      Bloqueado       =   -1  'True
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
   Begin VB.CheckBox chkVerificado 
      Caption         =   "Verificado"
      Enabled         =   0   'False
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
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin vbp_Ventas.ctlDataCombo cboCia 
      Height          =   315
      Left            =   4080
      TabIndex        =   78
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Cia:"
      Height          =   255
      Left            =   3840
      TabIndex        =   79
      Top             =   60
      Width           =   375
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido Materno"
      Height          =   195
      Left            =   5040
      TabIndex        =   61
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Despacho"
      Height          =   195
      Left            =   5880
      TabIndex        =   57
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      Height          =   195
      Left            =   2040
      TabIndex        =   56
      Top             =   60
      Width           =   450
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Observación"
      Height          =   195
      Left            =   120
      TabIndex        =   55
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
      Height          =   195
      Left            =   2760
      TabIndex        =   35
      Top             =   360
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   120
      TabIndex        =   33
      Top             =   60
      Width           =   600
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   195
      Left            =   1080
      TabIndex        =   34
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "ctlCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim objCliente As New clsCliente
'I.ECASTILLO 17.12.2020
Dim objWS As New clsWebService
Dim objLocal As New clsLocal
'F.ECASTILLO 17.12.2020
Private rsCliente As oraDynaset
Private rsDireccion As oraDynaset
Dim xArrTelefono As New XArrayDB
Dim blnModo As Boolean
Dim strPais As String
Dim strDepartamento As String
Dim strProvincia As String
Dim strDistrito As String
Dim strTelefono As String
Dim XstrTelefono As String
Dim strMaxTelefono As Integer
Dim strMaxAnexo As Integer
Dim strUbigeo As String
Dim strCodDireccionCli As String
Dim bolConsultaPrecio As Boolean

Dim strFlagJuridico As String
Dim strRazonSocial As String
Dim strDireccionSocial As String
Dim strDireccionComercial As String
Dim strNumDocumentoID As String
Dim strCodDocumentoID As String
Public seleccionManualLocal As Boolean

Public XTipoFuncion As String
'Public sCia As String ' jct, 29Mar12
'Public strFechaNac As String
Public Event Change()
Public Event Click(Area As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'I.ECASTILLO 17.12.2020
Dim flg_2e_reserva
Dim rsCia As oraDynaset
Dim flgFunLocal As String
'F.ECASTILLO 17.12.2020

'Property Let Modo(mvar As Boolean)
Public Sub Modo(mvar As Boolean, Optional r As Integer = 252, Optional g As Integer = 242, Optional b As Integer = 207)
    blnModo = mvar
    If blnModo = True Then
        UserControl.Height = 10095
        UserControl.BackColor = RGB(r, g, b) 'RGB(252, 242, 207)
        UserControl.Enabled = False
        chkVerificado.BackColor = RGB(r, g, b) 'RGB(252, 242, 207)
        optTipo(0).BackColor = RGB(r, g, b) 'RGB(252, 242, 207)
        optTipo(1).BackColor = RGB(r, g, b) 'RGB(252, 242, 207)
    Else
        UserControl.Height = 6975
        UserControl.Height = 7050
    End If
End Sub

Private Sub cboCia_Change()
    Dim objLocal As New clsLocal
 
    Set cboDespacho.RowSource = objLocal.Lista(cboCia.BoundText, "", "002")
    cboDespacho.ListField = "LOCAL_DEX2"
    cboDespacho.BoundColumn = "COD_LOCAL"
    cboDespacho.BoundText = ""
 
End Sub

Private Sub cboDepartamento_Change()
    
    Set cboProvincia.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, cboDepartamento.BoundText, "[ SELECCIONAR ]")
    cboProvincia.ListField = "Descripcion"
    cboProvincia.BoundColumn = "Codigo"
    cboProvincia.BoundText = "*"
    
    If (Codigo = "" Or Trim(txtDireccion.Text) = "") And strProvincia <> "" Then
        cboProvincia.BoundText = strProvincia
    End If
    
End Sub

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboDespacho_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboDistrito_Change()
On Error GoTo CtrlErr
    Dim rs As oraDynaset
    Dim objLocal As New clsLocal
    Dim strUbigeo As String
    
    If mdiPrincipal.ctlCliente1.ConsultaPrecio = True Then
        strUbigeo = mdiPrincipal.ctlCliente1.Ubigeo
        mdiPrincipal.ctlCliente1.ConsultaPrecio = False
    Else
        strUbigeo = cboDepartamento.BoundText & cboProvincia.BoundText & cboDistrito.BoundText
    End If
    
    strUbigeo = Replace(strUbigeo, "*", "")

    Set cboUrbanizacion.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_URBANIZACION", 0, "2", strUbigeo)
    cboUrbanizacion.ListField = "DES_URBANIZACION"
    cboUrbanizacion.BoundColumn = "COD_URBANIZACION"
    
    cboUrbanizacion.BoundText = "*"
    
    'Añadido por Jahzeel Lopez para determinar el local de precio y despacho.
    Dim CodLocalPrecio As String
    Dim CodLocalReferencia As String
    If objUsuario.CodLocalCallCenter = "1DLV" Then
        Set rs = objLocal.ListaLocalPredetDLV("94", strUbigeo)
    Else
        Set rs = objLocal.ListaLocalPredetDLV(objUsuario.CodigoEmpresa, strUbigeo)
    End If
    'Set rs = objLocal.ListaLocalPredetDLV(objUsuario.CodigoEmpresa, strUbigeo)
    If Not rs.EOF Then
        CodLocalPrecio = "" & rs("COD_LOCAL_PRECIO").Value
        CodLocalReferencia = "" & rs("COD_LOCAL_REF").Value
    End If
    Set rs = Nothing

    If txtCodigo.Text = "" And strUbigeo <> "" Then
        cboLocal.BoundText = CodLocalPrecio
        cboDespacho.BoundText = CodLocalReferencia
    ElseIf strUbigeo <> "" Then
        cboLocal.BoundText = CodLocalPrecio
    End If
    'I.ECASTILLO 17.12.2020 | 2da etapa reserva | 06.01.2021
    Dim flg_ruteoA_cnv
    flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto
    End If
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    Dim sCia As String
    Dim channel As String
    Dim Cia As String
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, cboDespacho.BoundText)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    'sCia = IIf(Len(Trim(sCia)) = 0, cboCia.BoundText, sCia)
    sCia = IIf(Len(Trim(sCia)) = 0, objUsuario.CodLocalCallCenter, sCia)
    flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "1" And flgFunLocal = "1" Then
        GoTo CallServiceReserva
    End If
    
    objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, cboDespacho.BoundText, "RESERVA_STOCK_2DA")
    
    If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
    Else
CallServiceReserva:
        Dim data As New Dictionary
        Dim x, i, j As Integer
        'Dim rsCia As oraDynaset
        'Dim sCia As String
        'Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, cboDespacho.BoundText)
        'If (rsCia.RecordCount > 0) Then
        '  sCia = CStr(rsCia(1))
        'End If
        'Set rsCia = Nothing
        'sCia = IIf(Len(Trim(sCia)) = 0, cboCia.BoundText, sCia)
        'ENVIAR CANAL Y CIA
        channel = "CALL"
        If strUbigeo <> "" Then
            Select Case sCia
                Case "94", "93", "92", "1DLV"
                    Cia = "IKF"
                    Set data = objWS.listaMetodosSegmentos_Ubigeo(strUbigeo, channel, Cia)
                Case ""
                    Set data = Nothing
                Case Else
                    Cia = "MF"
                    Set data = objWS.listaMetodosSegmentosMF_Ubigeo(strUbigeo, channel, Cia)
            End Select
            If Not data Is Nothing Then
                If IsObject(data("data")) = False Then
                    Exit Sub
                End If
                x = 0
                For x = 1 To data("data").Count()
                    If IsObject(data("data")(x)) = False Then
                        CodLocalPrecio = objLocal.GetCodBTL(data("data")("localCode"))
                    Else
                        CodLocalPrecio = objLocal.GetCodBTL(data("data")(x)("localCode"))
                    End If
                    'MsgBox "CodLocalPrecio: " & CodLocalPrecio
                    Sleep 75
                    CodLocalReferencia = CodLocalPrecio
                    'MsgBox "comboPrecio: " & cboLocal.BoundText
                    Sleep 75
                    cboLocal.BoundText = CodLocalPrecio
                    'MsgBox "comboPrecio: " & cboLocal.BoundText
                    Sleep 75
                    'cboDespacho.BoundText = CodLocalReferencia
                    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, CodLocalReferencia)
                    If (rsCia.RecordCount > 0) Then
                      sCia = CStr(rsCia(1))
                    End If
                    Set rsCia = Nothing
                    'MsgBox "Cia: " & sCia
                    Sleep 75
                    'Sleep 5000 ' to sleep for 1 second
                    'MsgBox "comboCia: " & cboCia.BoundText
                    Sleep 75
                    cboCia.BoundText = sCia
                    'MsgBox "comboCia: " & cboCia.BoundText
                    'Sleep 75
                    'Sleep 5000 ' to sleep for 1 second
                    'objVenta.respetaLocal = True
                    'MsgBox "cboDespacho: " & cboDespacho.BoundText
                    Sleep 75
                    cboDespacho.BoundText = CodLocalReferencia
                    'MsgBox "cboDespacho: " & cboDespacho.BoundText
                    Sleep 75
                    objVenta.bk_codLocal = CodLocalReferencia
                    If Len(Trim(CodLocalPrecio)) > 0 Then Exit For
                Next x
            End If
        End If
    End If
cnvNoRuteaAuto:
    'F.ECASTILLO 17.12.2020
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "cboDistrito.Change", Err.Description
End Sub


Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboLocal_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cboPais_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboProvincia_Change()
    Set cboDistrito.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, cboDepartamento.BoundText, cboProvincia.BoundText, "[ SELECCIONAR ]")
    cboDistrito.ListField = "Descripcion"
    cboDistrito.BoundColumn = "Codigo"
    
    cboDistrito.BoundText = "*"
    
    
    If (Codigo = "" Or Trim(txtDireccion.Text) = "") And strDistrito <> "" Then
        cboDistrito.BoundText = strDistrito
    End If
End Sub

Public Function Cargar(Optional Telefono As String)
    
    SetGrd
    Dim objLocal As New clsLocal
    
    '*************** B JCT 27MAR12, CARGA DE CIA  ***********************************************
    Set cboCia.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCia.ListField = "Des"
    cboCia.BoundColumn = "Cod"
    If objUsuario.CodLocalCallCenter = "1DLV" Then
        cboCia.BoundText = "94"
    Else
        cboCia.BoundText = objUsuario.CodigoEmpresa
    End If
    'cboCia.BoundText = objUsuario.CodigoEmpresa
    '****************E JCT 27MAR12 **********************************************
    If objUsuario.CodLocalCallCenter = "1DLV" Then
        Set cboLocal.RowSource = objLocal.Lista("94", "")
    Else
        Set cboLocal.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
    End If
    cboLocal.ListField = "local_dex2"
    cboLocal.BoundColumn = "COD_LOCAL"
    cboLocal.BoundText = objUsuario.CodigoLocal
    
    If objUsuario.CodLocalCallCenter = "1DLV" Then
        Set cboDespacho.RowSource = objLocal.Lista("94", "", "002")
    Else
        Set cboDespacho.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "", "002")
    End If
    cboDespacho.ListField = "local_dex2"
    cboDespacho.BoundColumn = "COD_LOCAL"
    cboDespacho.BoundText = objUsuario.CodigoLocal
    
    Set cboPais.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PAIS", 0)
    cboPais.ListField = "Descripcion"
    cboPais.BoundColumn = "Codigo"
    cboPais.BoundText = "00"

    Set cboDepartamento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0, "[ SELECCIONAR ]")
    cboDepartamento.ListField = "Descripcion"
    cboDepartamento.BoundColumn = "Codigo"
    cboDepartamento.BoundText = "*"
    
    strMaxTelefono = val(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_NUM_DIG_TELEFONO"))
    strMaxAnexo = val(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_NUM_DIG_ANEXO"))

    Dim strUbigeoCarga
    If mdiPrincipal.ctlCliente1.Ubigeo = "" Then
        strUbigeoCarga = "" & objUsuario.UbigeoLocal
    Else
        strUbigeoCarga = mdiPrincipal.ctlCliente1.Ubigeo
    End If
    
    If Not strUbigeoCarga = "" Then
    On Error GoTo Y
       strDepartamento = Mid(strUbigeoCarga, 1, 2)
       strProvincia = Mid(strUbigeoCarga, 3, 2)
       strDistrito = Mid(strUbigeoCarga, 5, 2)
           
''''    Set cboUrbanizacion.RowSource = objCliente.ListaUrbanizacion(strUbigeoCarga)
''''    cboUrbanizacion.ListField = "DES_URBANIZACION"
''''    cboUrbanizacion.BoundColumn = "COD_URBANIZACION"
       
Y:
    End If
    
    If strDepartamento <> "" Then
        cboDepartamento.BoundText = strDepartamento
    End If

    'dtpFechaNacimento.Value = gclsOracle.Execute("select sysdate from dual")
    Set cboEstadoCivil.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_ESTADO_CIVIL.FN_LISTA", 0, "")
    cboEstadoCivil.ListField = "DES_ESTADO_CIVIL"
    cboEstadoCivil.BoundColumn = "COD_ESTADO_CIVIL"

    Set cboTipoDocumento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_DOCUMENTO", 0, "")
    cboTipoDocumento.BoundColumn = "COD_DOCUMENTO_IDENTIDAD"
    cboTipoDocumento.ListField = "DES_DOCUMENTO_IDENTIDAD"
    cboTipoDocumento.ListField2 = "NUM_DIGITOS"

    Set cboSexo.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_SEXO", 0)
    cboSexo.ListField = "DESCRIPCION"
    cboSexo.BoundColumn = "CODIGO"

    Set cboSufijo.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_SUFIJO", 0)
    cboSufijo.ListField = "DES_ABREVIATURA"
    cboSufijo.BoundColumn = "COD_SUFIJO"

    Set ctlCboSuFijoDirecc.RowSource = objCliente.SuFijoDirecc
    ctlCboSuFijoDirecc.ListField = "DES_ABREVIATURA_DIRECCION"
    ctlCboSuFijoDirecc.BoundColumn = "COD_SUFIJO_DIRECCION"
    ctlCboSuFijoDirecc.ListField2 = "DES_SUFIJO_DIRECCION"

    Set cboTipoDireccion.RowSource = objCliente.ListaTipoDireccionCEN
    cboTipoDireccion.ListField = "DES_TIPO_DIRECCION"
    cboTipoDireccion.BoundColumn = "COD_TIPO_DIRECCION"
    
    'grdDirecciones.
    Dim rsContacto As oraDynaset
    Dim Item As New TrueDBGrid70.ValueItem
    Set rsContacto = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_CONT", 0)
    With TDBGrid1.Columns(0).ValueItems
        While Not rsContacto.EOF
            Item.Value = rsContacto(0).Value
            Item.DisplayValue = rsContacto(1).Value
            .Add Item
            rsContacto.MoveNext
        Wend
        .Translate = True
    End With
    Set rsContacto = Nothing
    
    If Not Telefono = "" Then
        xArrTelefono.ReDim 0, -1, 0, 3
        xArrTelefono.AppendRows
        xArrTelefono(0, 0) = "001"
        xArrTelefono(0, 1) = Telefono
        xArrTelefono(0, 2) = "1"
    End If
    XstrTelefono = Telefono
    
    TDBGrid1.Array = xArrTelefono
    TDBGrid1.Rebind
    '----------------------------------
    chkPrincipal.Value = vbChecked
    cboTipoDireccion.BoundText = "001"
    SSTab1.Tab = 0
End Function

Function Grabar()
TDBGrid1.MoveNext
TDBGrid1.MovePrevious
Dim strTipoLinea As String
Dim strValorLinea As String
Dim strDafaultLinea As String
Dim strAnexo As String
Dim h As Integer
    
    'I.ECASTILLO 22.06.2021
    Dim flg_ruteoA_cnv
    flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
    Dim sCia As String
    Dim rsCia As oraDynaset
    Dim flgFunLocal As String
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    
    If objVenta.respetaLocal <> True Then
        GoTo cnvNoRuteaAuto
    End If
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto
    End If
    Set rsCia = Nothing
    sCia = ""
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, cboDespacho.BoundText)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "1" And flgFunLocal = "1" Then
        GoTo CallServiceReserva
    End If
    
    objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, objVenta.LocalDespacho, "RESERVA_STOCK_2DA")
    
    If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
        GoTo cnvNoRuteaAuto
    Else
CallServiceReserva:
        If Len(Trim(objVenta.dc_street)) <> 0 And Trim(objVenta.bk_codCliente) = Trim(txtCodigo.Text) And _
            Trim(txtDireccion.Text) <> Trim(objVenta.dc_street) Then
            If MsgBox("Ud. geolocalizó la dirección " & objVenta.dc_street & ", desea modificar la dirección original?", _
                        vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
                txtDireccion.Text = "" & Trim(objVenta.dc_street)
                objVenta.dc_street = IIf(Len(Trim(objVenta.dc_street)) = 0, "", txtDireccion.Text)
                objVenta.DireccionClienteDLV = txtDireccion.Text
                
                objVenta.DireccionCliente = objVenta.DireccionClienteDLV
                objVenta.Out_Direccion = objVenta.DireccionClienteDLV
            Else
                GoTo cnvNoRuteaAuto
            End If
        Else
cnvNoRuteaAuto:
            'txtDireccion.Text = "" & rsDireccion("DES_DIRECCION").Value
        End If
    End If
    'F.ECASTILLO 22.06.2021
    If Trim(txtDireccion.Text) = "" Then MsgBox "Ingrese la Dirección del Cliente", vbInformation: Grabar = "Error": ctlCboSuFijoDirecc.SetFocus: Exit Function
    
    While h <= xArrTelefono.UpperBound(1)
        If Abs(val(xArrTelefono(h, 2))) = "1" Then strTelefono = xArrTelefono(h, 1)
        strTipoLinea = strTipoLinea & xArrTelefono(h, 0) & "|"
        strValorLinea = strValorLinea & xArrTelefono(h, 1) & "|"
        strDafaultLinea = strDafaultLinea & Abs(val(xArrTelefono(h, 2))) & "|"
        strAnexo = strAnexo & Abs(val(xArrTelefono(h, 3))) & "|"
        h = h + 1
    Wend
    
    Dim arra As Variant
    arra = Array("1A_COD_CLIENTE", "2A_COD_LOCAL_DESPACHO", "3A_COD_ESTADO_CIVIL", _
                 "4A_DES_NOM_CLIENTE", "5A_DES_APE_CLIENTE", "6A_DES_RAZON_SOCIAL", _
                 "7A_DES_NOM_COMERCIAL", "8A_COD_DOCUMENTO_IDENTIDAD", "9A_NUM_DOCUMENTO_ID", _
                 "10A_DES_DIRECCION_SOCIAL", "11A_DES_DIRECCION_COMERCIAL", "12A_DES_EMAIL", "13A_FLG_TIPO_JURIDICA", _
                 "14A_DES_OBSERVACION", "15A_DES_CARGO", "16A_FLG_SEXO", "17A_FCH_NACIMIENTO", _
                 "18A_NUM_EMPLEADOS", "19A_NUM_HIJOS", "20A_COD_CLIENTE_REF", "21A_FCH_ULTIMA_COMPRA", _
                 "22A_FCH_ULTIMO_PEDIDO", "23A_NUM_ULTIMO_DOCUMENTO", "24A_FLG_ESTADO", "25A_COD_SUFIJO", _
                 "26A_FLG_CLIENTE_VERIFICADO", "27A_COD_USUARIO", "28A_COD_ZONA", _
                 "29A_CAD_COD_TIPO_LINEA_CONTACTO", "30A_CAD_DES_VALOR", "31A_CAD_FLG_PRINCIPAL", "32a_ubigeo", "Cod_local_precio", "Compañia", "COD_SUFIJO_DIRECCION", "COD_URBANIZACION", "DES_APE2_CLIENTE", "FLG_PRINCIPAL", "COD_TIPO_DIRECCION", "COD_DIRECCION_CLI", "COD_LOCAL_CREA")
  
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim strUbigeo As String
    strUbigeo = Replace(cboDepartamento.BoundText & cboProvincia.BoundText & cboDistrito.BoundText, "*", "")
    
    If Len(strUbigeo) < 6 Then
        Err.Raise 1, "", "No ha seleccionado todos los datos del ubigeo"
    End If
    
    If chkAñadir.Value = 1 Then objVenta.CodDireccionCli = ""

    'If dtpFechaNacimento.Value = Date Then strFechaNac = "" Else strFechaNac = dtpFechaNacimento.Value
    
    If optTipo(1).Value = True Then
        arrValores = Array(Trim(txtCodigo.Text), cboDespacho.BoundText, cboEstadoCivil.BoundText, "", "", txtNombre.Text, _
                           txtApellido.Text, cboTipoDocumento.BoundText, txtNumDocumento.Text, _
                           "", txtDireccion.Text, TxtEmail.Text, "1", txtObservacion.Text, "", cboSexo.BoundText, _
                           dtpFechaNacimento.Value, "", txtNumHijos.Text, "", "", "", "", _
                           "1", cboSufijo.BoundText, chkVerificado.Value, objUsuario.Codigo, "", strTipoLinea, strValorLinea, strDafaultLinea, strAnexo, strUbigeo, txtReferencia.Text, cboLocal.BoundText, objUsuario.CodigoEmpresa, ctlCboSuFijoDirecc.BoundText, Replace(cboUrbanizacion.BoundText, "*", ""), txtApeMaterno.Text, chkPrincipal.Value, cboTipoDireccion.BoundText, objVenta.CodDireccionCli, IIf(objUsuario.flgDeliveryProv = "1", Mid(objUsuario.NombrePC, 4, 3), objUsuario.CodigoLocal))
                        
    Else
        arrValores = Array(Trim(txtCodigo.Text), cboDespacho.BoundText, cboEstadoCivil.BoundText, txtNombre.Text, _
                           txtApellido.Text, "", "", cboTipoDocumento.BoundText, txtNumDocumento.Text, _
                           txtDireccion.Text, "", TxtEmail.Text, "0", txtObservacion.Text, "", cboSexo.BoundText, _
                           dtpFechaNacimento.Value, "", txtNumHijos.Text, "", "", "", "", _
                           "1", cboSufijo.BoundText, chkVerificado.Value, objUsuario.Codigo, "", strTipoLinea, strValorLinea, strDafaultLinea, strAnexo, strUbigeo, txtReferencia.Text, cboLocal.BoundText, objUsuario.CodigoEmpresa, ctlCboSuFijoDirecc.BoundText, Replace(cboUrbanizacion.BoundText, "*", ""), txtApeMaterno.Text, chkPrincipal.Value, cboTipoDireccion.BoundText, objVenta.CodDireccionCli, IIf(objUsuario.flgDeliveryProv = "1", Mid(objUsuario.NombrePC, 4, 3), objUsuario.CodigoLocal))
    End If
    '      MsgBox UBound(arrValores)
    arrDireccion = Array(Entrada_Salida, entrada, entrada, entrada, _
                         entrada, entrada, entrada, entrada, entrada, _
                         entrada, entrada, entrada, entrada, entrada, entrada, entrada, _
                         entrada, entrada, entrada, entrada, entrada, entrada, entrada, _
                         entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
 '   MsgBox UBound(arrDireccion)
    Dim strMensaje As String
    strMensaje = gclsOracle.SP("BTLPROD.PKG_CLIENTE.SP_GRABAR", arrValores, arrDireccion)
    
    If strMensaje = "" Then
        MsgBox "Se grabo satisfactoriamente", vbExclamation, App.ProductName
        Me.Codigo = arrValores(0)
        'I.ECASTILLO 07.07.2021
        objVenta.bk_DescSufijoDir = "" & ctlCboSuFijoDirecc.BoundText2
        objVenta.bk_AbrSufijoDir = "" & ctlCboSuFijoDirecc.Text
        If gstrFlagSufDir = "1" Then
            objVenta.bk_SufijoDir = objVenta.bk_AbrSufijoDir
        ElseIf gstrFlagSufDir = "2" Then
            objVenta.bk_SufijoDir = objVenta.bk_DescSufijoDir
        End If
        'F.ECASTILLO 07.07.2021
        'ConsultaCliente CStr(arrValores(0))
    Else
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    Grabar = strMensaje
    'I.ECASTILLO 17.12.2020
    'objVenta.respetaLocal = True
    objVenta.bk_codLocal = cboDespacho.BoundText
    'F.ECASTILLO 17.12.2020
End Function

Public Function ConsultaCliente(Codigo As String, Optional Cia As String)
    Dim strUbigeo As String
    Dim strFlagJuridico As String
    Dim strUrbanizacion As String
    Dim strCodDireccionCli As String
    Dim sCia As String
    
    'Codigo = "00000001"
    If Codigo = "" Then Exit Function
    Set rsCliente = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA", 0, Codigo, Cia)
    '''''''''''''''
    
    
       
    chkVerificado.Value = val("" & rsCliente("FLG_CLIENTE_VERIFICADO").Value)
    
    If Not rsCliente Is Nothing Then
        txtCodigo.Text = "" & rsCliente("COD_CLIENTE").Value
        LocalAsignado = "" & rsCliente("COD_LOCAL_DESPACHO").Value
        cboEstadoCivil.BoundText = "" & rsCliente("COD_ESTADO_CIVIL").Value
        strFlagJuridico = "" & rsCliente("FLG_TIPO_JURIDICA").Value
        
        If CodDireccionCli = "" Then
            strCodDireccionCli = "" & rsCliente("COD_DIRECCION_CLI").Value
        Else
            strCodDireccionCli = CodDireccionCli
        End If
    
        Set rsDireccion = objCliente.ListaDireccion(txtCodigo.Text, strCodDireccionCli)
        
        If Not rsDireccion.EOF Then
                txtDireccion.Text = "" & rsDireccion("DES_DIRECCION").Value
                ctlCboSuFijoDirecc.BoundText = "" & rsDireccion("COD_SUFIJO_DIRECCION").Value
                txtReferencia.Text = "" & rsDireccion("DES_REFERENCIA_DIRECCION").Value
                strUbigeo = "" & rsDireccion("COD_UBIGEO").Value
                cboTipoDireccion.BoundText = "" & rsDireccion("COD_TIPO_DIRECCION").Value
                strUrbanizacion = "" & rsDireccion("COD_URBANIZACION").Value
                chkPrincipal.Value = val("" & rsDireccion("FLG_PRINCIPAL").Value)
                objVenta.CodDireccionCli = "" & rsDireccion("COD_DIRECCION_CLI").Value
        End If
        
        If strFlagJuridico = "1" Then
            txtNombre.Text = "" & rsCliente("DES_RAZON_SOCIAL").Value
            txtApellido.Text = "" & rsCliente("DES_NOM_COMERCIAL").Value
            '''txtDireccion.Text = "" & rsCliente("DES_DIRECCION_COMERCIAL").Value
            optTipo(1).Value = True
            Label2.Caption = "Razón Social"
            Label3.Caption = "Razón Comercial"
        Else
            txtNombre.Text = "" & rsCliente("DES_NOM_CLIENTE").Value
            txtApellido.Text = "" & rsCliente("DES_APE_CLIENTE").Value ''''&
            txtApeMaterno.Text = "" & rsCliente("DES_APE2_CLIENTE").Value
            '''txtDireccion.Text = "" & rsCliente("DES_DIRECCION_SOCIAL").Value
            Label2.Caption = "Nombre"
            Label3.Caption = "Apellido"
            optTipo(0).Value = True
        End If
        
        cboTipoDocumento.BoundText = "" & rsCliente("COD_DOCUMENTO_IDENTIDAD").Value
        txtNumDocumento.Text = "" & rsCliente("NUM_DOCUMENTO_ID").Value
        
        txtObservacion.Text = "" & rsCliente("DES_OBSERVACION").Value
        'DireccionSocial = rsCliente("DES_DIRECCION_SOCIAL").Value
        'DireccionComercial = rsCliente("DES_DIRECCION_COMERCIAL").Value
        TxtEmail.Text = "" & rsCliente("DES_EMAIL").Value    '
        'Observacion = rsCliente("DES_OBSERVACION").Value
        'Cargo = rsCliente("DES_CARGO").Value
        cboSexo.BoundText = "" & rsCliente("FLG_SEXO").Value
        dtpFechaNacimento.Value = rsCliente("FCH_NACIMIENTO").Value 'IIf(IsNull(rsCliente("FCH_NACIMIENTO").Value), Date, rsCliente("FCH_NACIMIENTO").Value)
        txtNumHijos.Text = "" & rsCliente("NUM_HIJOS").Value
        'NumeroEmpleados = rsCliente("NUM_EMPLEADOS").Value
        'ClienteReferencia = rsCliente("COD_CLIENTE_REF").Value
        'UltimaCompra = rsCliente("FCH_ULTIMA_COMPRA").Value
        'UltimoPedido = rsCliente("FCH_ULTIMO_PEDIDO").Value
        'UltimoDocumento = rsCliente("NUM_ULTIMO_DOCUMENTO").Value
        'Estado = rsCliente("FLG_ESTADO").Value
        txtUsuarioReg.Text = "" & rsCliente("COD_USUARIO").Value
        txtFechaReg.Text = "" & rsCliente("FCH_REGISTRA").Value
        txtUsuarioAct.Text = "" & rsCliente("COD_USUARIO_ACTUALIZA").Value
        txtFechaAct.Text = "" & rsCliente("FCH_ACTUALIZA").Value
        
        If Not strUbigeo = "" Then
            On Error GoTo Y
                cboDepartamento.BoundText = Mid(strUbigeo, 1, 2)
                cboProvincia.BoundText = Mid(strUbigeo, 3, 2)
                cboDistrito.BoundText = Mid(strUbigeo, 5, 2)
                
                If strUrbanizacion <> "" Then
                    cboUrbanizacion.BoundText = strUrbanizacion
                End If
Y:
        End If
        cboSufijo.BoundText = "" & rsCliente("COD_SUFIJO").Value
        '''''''''''''''
        Dim rsCliente2 As oraDynaset
        Set rsCliente2 = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_LINEA_CONTAC", 0, Codigo)
        
        Dim h As Integer
        Dim Xaq As Boolean
        Xaq = False
        If Not rsCliente2.EOF Then
            Xaq = True
            xArrTelefono.ReDim 0, -1, 0, 3
        End If
        
        While Not rsCliente2.EOF
            xArrTelefono.AppendRows
            xArrTelefono(h, 0) = "" & rsCliente2("COD_TIPO_LINEA_CONTACTO")
            xArrTelefono(h, 1) = "" & rsCliente2("DES_VALOR")
            xArrTelefono(h, 2) = "" & rsCliente2("FLG_PRINCIPAL")
            If xArrTelefono(h, 2) = "1" Then strTelefono = "" & rsCliente2("DES_VALOR")
            xArrTelefono(h, 3) = "" & rsCliente2("NUM_ANEXO")
            h = h + 1
            rsCliente2.MoveNext
        Wend
        If XTipoFuncion = "Nuevo" Or XTipoFuncion = "Otro" Or XTipoFuncion = "Editar" Then
        Dim rsTelefono As oraDynaset
        Set rsTelefono = objCliente.TelefonoCliente(Codigo, XstrTelefono)
            If rsTelefono.EOF() And XstrTelefono <> "" Then
                If Xaq = True Then
                    xArrTelefono.AppendRows
                    xArrTelefono(h, 0) = "001"
                    xArrTelefono(h, 1) = XstrTelefono
                    xArrTelefono(h, 2) = "0"
                    xArrTelefono(h, 3) = "0"
                    h = h + 1
                End If
            End If
        End If
        TDBGrid1.Close
        TDBGrid1.Array = xArrTelefono
        TDBGrid1.Rebind
        Set rsCliente2 = Nothing
    End If
    
    
    cboLocal.BoundText = "" & rsCliente("COD_LOCAL_PRECIO").Value
    ''''''Me.LocalAsignado = "" & rsCliente("COD_LOCAL_PRECIO").Value
    
    cboDespacho.BoundText = "" & rsCliente("COD_LOCAL_DESPACHO").Value
    
    ' B JCT, el combo de cia, debe ser segun local de despacho
    sCia = ""
    sCia = "" & gclsOracle.FN_Valor("btlprod.pkg_local.fn_devuelve_cia", rsCliente("COD_LOCAL_DESPACHO").Value)
    
    If (sCia <> "") Then
        cboCia.BoundText = sCia
    Else
        cboCia.BoundText = objUsuario.CodigoEmpresa
    End If
    ' E JCT
    
    'MsgBox "## lOCAL DESPACHO EN CONSULTA CLIENTE :" + cboDespacho.BoundText
    
    ' Reasignar combo de despacho
    cboDespacho.BoundText = "" & rsCliente("COD_LOCAL_DESPACHO").Value
    
    'I.ECASTILLO 17.12.2020
    If objVenta.respetaLocal = True Then
        objVenta.bk_codLocal = "" & objVenta.bk_codLocal
        Me.reasignaLocales objVenta.bk_codLocal, objVenta.bk_codLocal
    'Else
    '    objVenta.bk_codLocal = "" & rsCliente("COD_LOCAL_DESPACHO").Value
    End If
    'F.ECASTILLO 17.12.2020
    Set rsTelefono = Nothing
End Function

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboSufijo_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cboTipoDocumento_Click(Area As Integer)
On Error GoTo handle
    If cboTipoDocumento.BoundText = "" Then Exit Sub
    txtNumDocumento.MaxLength = val(cboTipoDocumento.BoundText2)
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Property Get Verificacion() As String
    Verificacion = IIf(chkVerificado.Value = "1", "1", "0")
End Property

Property Let Verificacion(mvar As String)
    'cboEstadoCivil.BoundText = mvar
    IIf(chkVerificado.Value = True, "1", "0") = mvar
End Property

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

Private Sub cmdAñadir_Click()
Dim strMensaje As String
Dim strUbigeo As String
Dim strCodDireccionCli As String

On Error GoTo CtrlErr
    If MsgBox("Desea asignar esta dirección al cliente " & txtCodigo.Text & " " & txtNombre.Text & " " & txtApellido.Text & " " & txtApeMaterno.Text, vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
''        strUbigeo = cboDepartamento.BoundText & cboProvincia.BoundText & cboDistrito.BoundText
''
''        strMensaje = objCliente.GrabaAuxDireccionCliente(txtCodigo.Text, _
''                             cboTipoDireccion.BoundText, _
''                             txtDireccion.Text, _
''                             txtReferencia.Text, _
''                             strUbigeo, _
''                             cboUrbanizacion.BoundText, _
''                             ctlCboSuFijoDirecc.BoundText, _
''                             chkPrincipal.Value)
''        If strMensaje = "" Then
''             MsgBox "Se añadio la dirección satisfactoriamente ", vbExclamation, App.ProductName
''        Else
''             MsgBox strMensaje, vbCritical, App.ProductName
''        End If
        objVenta.CodDireccionCli = ""
    End If
Exit Sub
CtrlErr:
    Err.Raise Err.Number, "ctlCliente.Añadir", Err.Description
End Sub

Private Sub cmdBuscar_Click()
Dim objCliente As New clsCliente
On Error GoTo CtrlErr
    'Set GrdBusCliente.DataSource = objCliente.ListaCliente(txtNombre.Text & "%" & txtApellido.Text, IIf(optTipo(0).Value = True, 0, 1), "", "", "")
    Set GrdBusCliente.DataSource = objCliente.ListaCliente(txtApellido.Text & "%" & txtNombre.Text, IIf(optTipo(0).Value = True, 0, 1), "", "", "")
    Call SetGrd
    GrdBusCliente.Visible = True
    GrdBusCliente.MarqueeStyle = dbgHighlightRow
    GrdBusCliente.SetFocus
    fraLabel(0).Visible = True
    Set objCliente = Nothing
Exit Sub
CtrlErr:
    Err.Raise Err.Number, "ctlCliente.cmdBuscar", Err.Description
End Sub

Private Sub cmdVerDireccion_Click()
Dim objCliente As New clsCliente
    Set grdDirecciones.DataSource = objCliente.ListaDireccion(txtCodigo.Text)
    grdDirecciones.Visible = True
    grdDirecciones.MarqueeStyle = dbgHighlightRow
    grdDirecciones.SetFocus
    fraLabel(1).Visible = True
    Set objCliente = Nothing
End Sub

Private Sub ctlCboSuFijoDirecc_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub


Private Sub GrdBusCliente_DblClick()
On Error GoTo CtrlErr
    If GrdBusCliente.ApproxCount = 0 Then Exit Sub
    XTipoFuncion = "Otro"
    ConsultaCliente GrdBusCliente.Columns("COD_CLIENTE").Value
    GrdBusCliente.Visible = False
    fraLabel(0).Visible = False
    'SendKeys "{TAB}"
    ctlCboSuFijoDirecc.SetFocus
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "ctlCliente.DblClick", Err.Description
End Sub

Private Sub GrdBusCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            GrdBusCliente_DblClick
        Case vbKeyF3
            GrdBusCliente.Visible = False
            fraLabel(0).Visible = False
            ctlCboSuFijoDirecc.SetFocus
        Case vbKeyUp
            If GrdBusCliente.PrimerRegistro Then
                GrdBusCliente.Visible = False
                fraLabel(0).Visible = False
                txtApellido.SetFocus
            End If
       Case vbKeyEscape
            MsgBox "PARA SALIR PRESIONAR LA TECLA DE FUNCION [F3]", vbExclamation + vbOKOnly + vbDefaultButton1, App.ProductName
    End Select
End Sub

Private Sub grdDirecciones_DblClick()
    ctlCboSuFijoDirecc.BoundText = "" & grdDirecciones.Columns("COD_SUFIJO_DIRECCION").Value
    txtDireccion.Text = "" & grdDirecciones.Columns("DES_DIRECCION").Value
    txtReferencia.Text = "" & grdDirecciones.Columns("DES_REFERENCIA_DIRECCION").Value
    cboDepartamento.BoundText = Mid("" & grdDirecciones.Columns("COD_UBIGEO").Value, 1, 2)
    cboProvincia.BoundText = Mid("" & grdDirecciones.Columns("COD_UBIGEO").Value, 3, 2)
    cboDistrito.BoundText = Mid("" & grdDirecciones.Columns("COD_UBIGEO").Value, 5, 2)
    cboUrbanizacion.BoundText = "" & grdDirecciones.Columns("COD_URBANIZACION").Value
    cboTipoDireccion.BoundText = "" & grdDirecciones.Columns("COD_TIPO_DIRECCION").Value
    chkPrincipal.Value = val("" & grdDirecciones.Columns("FLG_PRINCIPAL").Value)
    objVenta.CodDireccionCli = "" & grdDirecciones.Columns("COD_DIRECCION_CLI").Value
    grdDirecciones.Visible = False
    fraLabel(1).Visible = False
    cboUrbanizacion.SetFocus

End Sub

Private Sub grdDirecciones_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdDirecciones_DblClick
        Case vbKeyF3
            grdDirecciones.Visible = False
            fraLabel(1).Visible = False
            cboUrbanizacion.SetFocus
        Case vbKeyUp
            If grdDirecciones.PrimerRegistro Then
                grdDirecciones.Visible = False
                fraLabel(1).Visible = False
                txtReferencia.SetFocus
            End If
       Case vbKeyEscape
            MsgBox "<F3> Para salir", vbExclamation + vbOKOnly + vbDefaultButton1, App.ProductName
    End Select
    
    
    
End Sub


Private Sub optTipo_Click(Index As Integer)
    If optTipo(1).Value = True Then
        Label2.Caption = "Razón Social"
        Label3.Caption = "Razón Comercial"
        Label27.Visible = False
        txtApeMaterno.Visible = False
    Else
        Label2.Caption = "Nombre"
        Label3.Caption = "Apellido"
        Label27.Visible = True
        txtApeMaterno.Visible = False
    End If
End Sub

Private Sub optTipo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    
    End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        With TDBGrid1
            If Not .Columns(.Col).Value = "0" Then
                Dim i As Integer
                While i < .ApproxCount
                    If Not .row = i Then
                        xArrTelefono(i, 2) = "0"
                    End If
                    i = i + 1
                Wend
                .RefetchCol 2
            End If
        End With
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If TDBGrid1.ApproxCount = 0 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        TDBGrid1.Delete
    End If
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)

    If Len(TDBGrid1.Columns(3).Value) >= strMaxAnexo Then
         KeyAscii = 0
            Exit Sub
    End If
        If Len(TDBGrid1.Columns(1).Value) >= strMaxTelefono Then
         KeyAscii = 0
            Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub 'BackScape
    If KeyAscii = 39 Then
        If InStr(1, TDBGrid1.Columns(TDBGrid1.Col).Value, "'") <> 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If TDBGrid1.Col = 1 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtApellido_KeyDown(KeyCode As Integer, Shift As Integer)
''    If KeyCode = vbKeyDown And GrdBusCliente.Visible = True Then
''        GrdBusCliente.MarqueeStyle = dbgHighlightRow
''        GrdBusCliente.SetFocus
''        fraLabel.Visible = True
''    End If
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)

    Dim strNombre As String
    strNombre = Trim(txtApellido.Text)
    
    
    
    If KeyAscii = 13 Then
        If GrdBusCliente.Visible Then GrdBusCliente.Visible = False: fraLabel(0).Visible = False
    Else
'        If Len(strNombre) > 3 Then
'            Set GrdBusCliente.DataSource = objCliente.ListaCliente(strNombre, "", "", "", "")  ', strctlCboTipCliente, strcboTipoDocumento, strNumeroDocumento, "") 'strflgActivo)
'            Call SetGrd
'            GrdBusCliente.Visible = True
'            fraLabel.Visible = True
'        Else
'            GrdBusCliente.Visible = False
'            fraLabel.Visible = False
'        End If
        
        If Len(strNombre) > 1 Then
            cmdBuscar.Visible = True
        Else
            cmdBuscar.Visible = False
        End If
              
    End If

End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)

    Dim strDireccion As String
    strDireccion = Trim(txtDireccion.Text)
    
    
    
    If KeyAscii = 13 Then
        If grdDirecciones.Visible Then grdDirecciones.Visible = False: fraLabel(1).Visible = False
    Else
        
        If Len(strDireccion) > 1 Then
            cmdVerDireccion.Visible = True
        Else
            cmdVerDireccion.Visible = False
        End If
        
        
        
        
    End If


End Sub

Private Sub UserControl_Initialize()
    xArrTelefono.ReDim 0, -1, 0, 3
    Set TDBGrid1.Array = xArrTelefono
    cmdBuscar.Visible = False
    cmdVerDireccion.Visible = False
    
End Sub


Private Sub UserControl_Terminate()
    Set rsCliente = Nothing
End Sub
Property Get Codigo() As String
    Codigo = txtCodigo.Text
End Property
Property Let Codigo(mvar As String)
    txtCodigo.Text = mvar
End Property

Public Sub Verificar()
    chkVerificado.Enabled = True
End Sub

Property Get LocalAsignado() As String
    If Not rsCliente Is Nothing And objVenta.respetaLocal = False Then
        LocalAsignado = rsCliente("COD_LOCAL_PRECIO").Value
        cboDespacho.BoundText = "" & rsCliente("COD_LOCAL_DESPACHO").Value
    Else
        LocalAsignado = cboLocal.BoundText
    End If
End Property
Property Let LocalAsignado(mvar As String)
'    LocalAsignado = mvar 'Codigo comentado por desarrollar
    cboLocal.BoundText = mvar
End Property

Property Get sCia() As String
    If Not rsCliente Is Nothing Then
        sCia = "" & gclsOracle.FN_Valor("btlprod.pkg_local.fn_devuelve_cia", rsCliente("COD_LOCAL_DESPACHO").Value)
    Else
        sCia = cboCia.BoundText
    End If
End Property
Property Let sCia(mvar As String)
    cboCia.BoundText = mvar
End Property

Public Sub DeshabilitaLocalPrecio()
    cboLocal.Enabled = False
End Sub

Property Get CodigoEstadoCivil() As String
    CodigoEstadoCivil = cboEstadoCivil.BoundText
End Property
Property Let CodigoEstadoCivil(mvar As String)
    cboEstadoCivil.BoundText = mvar
End Property

Property Get Telefono() As String
    Telefono = strTelefono
End Property


Property Get Nombre() As String
    Nombre = txtNombre.Text
End Property
Property Let Nombre(mvar As String)
    txtNombre.Text = mvar
End Property



Property Get Apellido() As String
    Apellido = txtApellido.Text
End Property
Property Let Apellido(mvar As String)
    txtApellido.Text = mvar
End Property

Property Get Email() As String
    Email = TxtEmail.Text
End Property
Property Let Email(mvar As String)
    TxtEmail.Text = mvar
End Property

Property Get Referencia() As String
    Referencia = txtReferencia.Text
End Property
Property Let Referencia(mvar As String)
    txtReferencia.Text = mvar
End Property

Property Get Observacion() As String
    Observacion = txtObservacion.Text
End Property
Property Let Observacion(mvar As String)
    txtObservacion.Text = mvar
End Property

Property Get RazonSocial() As String
    If Not rsCliente Is Nothing Then
        RazonSocial = "" & rsCliente("DES_RAZON_SOCIAL").Value
    Else
        RazonSocial = strRazonSocial
    End If
End Property
Property Get NombreComercial() As String
    If Not rsCliente Is Nothing Then NombreComercial = "" & rsCliente("DES_NOM_COMERCIAL").Value
End Property
Property Get CodigoDocumentoID() As String
    If Not rsCliente Is Nothing Then
        CodigoDocumentoID = "" & rsCliente("COD_DOCUMENTO_IDENTIDAD").Value
    Else
        CodigoDocumentoID = strCodDocumentoID
    End If
End Property

Property Get NumeroDocumentoID() As String
    If Not rsCliente Is Nothing Then
        NumeroDocumentoID = "" & rsCliente("NUM_DOCUMENTO_ID").Value
    Else
        NumeroDocumentoID = strNumDocumentoID
    End If
End Property
Property Get DireccionSocial() As String
    If Not rsCliente Is Nothing Then
        DireccionSocial = "" & rsCliente("DES_DIRECCION_SOCIAL").Value
    Else
        DireccionSocial = strDireccionSocial
    End If
End Property
Property Get DireccionComercial() As String
    If Not rsCliente Is Nothing Then
        DireccionComercial = "" & rsCliente("DES_DIRECCION_COMERCIAL").Value
    Else
        DireccionComercial = strDireccionComercial
    End If
End Property
Property Get FlagJuridico() As String
    If Not rsCliente Is Nothing Then
        FlagJuridico = "" & rsCliente("FLG_TIPO_JURIDICA").Value
    Else
        FlagJuridico = strFlagJuridico
    End If
End Property



Property Let FlagJuridico(ByVal lstrFlagJuridico As String)
    strFlagJuridico = lstrFlagJuridico
End Property




Property Get Cargo() As String
    If Not rsCliente Is Nothing Then Cargo = "" & rsCliente("DES_CARGO").Value
End Property
Property Get Sexo() As String
    If Not rsCliente Is Nothing Then Sexo = "" & rsCliente("FLG_SEXO").Value
End Property

Property Get FechaNacimiento() As String
'    FechaNacimiento = IIf(IsNull(rsCliente("FCH_NACIMIENTO").Value), Date, rsCliente("FCH_NACIMIENTO").Value)
End Property
Property Let FechaNacimiento(mvar As String)
    FechaNacimiento = dtpFechaNacimento.Value
End Property


Property Get NumeroHijos() As String
    If Not rsCliente Is Nothing Then NumeroHijos = "" & rsCliente("NUM_HIJOS").Value
End Property



Property Get NumeroEmpleados() As String
    If Not rsCliente Is Nothing Then NumeroEmpleados = "" & rsCliente("NUM_EMPLEADOS").Value
End Property
Property Get ClienteReferencia() As String
    If Not rsCliente Is Nothing Then ClienteReferencia = "" & rsCliente("COD_CLIENTE_REF").Value
End Property
Property Get UltimaCompra() As String
    If Not rsCliente Is Nothing Then UltimaCompra = "" & rsCliente("FCH_ULTIMA_COMPRA").Value
End Property
Property Get UltimoPedido() As String
    If Not rsCliente Is Nothing Then UltimoPedido = "" & rsCliente("FCH_ULTIMO_PEDIDO").Value
End Property
Property Get UltimoDocumento() As String
    If Not rsCliente Is Nothing Then UltimoDocumento = "" & rsCliente("NUM_ULTIMO_DOCUMENTO").Value
End Property
Property Get Estado() As String
    If Not rsCliente Is Nothing Then Estado = "" & rsCliente("FLG_ESTADO").Value
End Property
Property Get Sufijo() As String
    If Not rsCliente Is Nothing Then Sufijo = "" & rsCliente("COD_SUFIJO").Value
End Property


Public Sub Limpiar()
On Error GoTo handle
Codigo = ""
txtApeMaterno.Text = ""
chkVerificado.Value = vbUnchecked
chkPrincipal.Value = vbChecked
txtCodigo.Text = ""
cboEstadoCivil.BoundText = ""
txtNombre.Text = ""
txtApellido.Text = ""
txtDireccion.Text = ""
optTipo(1).Value = True
txtNombre.Text = ""
txtApellido.Text = ""
txtDireccion.Text = ""
cboTipoDocumento.BoundText = "001"
txtNumDocumento.Text = ""
cboLocal.BoundText = "" 'estaba comentado
cboDespacho.BoundText = "" 'estaba comentado
cboCia.BoundText = "" 'no existia
txtObservacion.Text = ""
txtReferencia.Text = ""
TxtEmail.Text = ""
cboSexo.BoundText = ""
txtNumHijos.Text = ""
txtUsuarioReg.Text = ""
txtFechaReg.Text = ""
txtUsuarioAct.Text = ""
txtFechaAct.Text = ""
cboDepartamento.BoundText = "*"
'cboProvincia.BoundText = ""
'cboDistrito.BoundText = ""
cboSufijo.BoundText = ""
'cboUrbanizacion.BoundText = ""
xArrTelefono.ReDim 0, -1, 0, 3
TDBGrid1.Close
TDBGrid1.Array = xArrTelefono
TDBGrid1.Rebind
Set rsCliente = Nothing
Set rsDireccion = Nothing
Set objCliente = Nothing
objVenta.dc_street = ""
objVenta.dc_number = ""
objVenta.dc_apartment = ""
objVenta.dc_country = ""
objVenta.dc_city = ""
objVenta.dc_district = ""
objVenta.dc_latitude = ""
objVenta.dc_longitude = ""
objVenta.bk_codCliente = ""
objVenta.dc_departamentBK = ""
objVenta.dc_referencia = ""
objVenta.dc_urbanizacion = ""
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub SetGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

           arrCampos = Array("COD_CLIENTE", "NOMBRE", "NUMERO", "FLG_TIPO_JURIDICA", "DIRECCION")
                              
           arrCaption = Array("Codigo", "Nombre", _
                              "Doc.Ident.", "Jurídica", _
                              "Dirección")
                               
           arrAncho = Array(900, 3000, _
                            1000, 800, _
                            2000)
                             
           arrAlineacion = Array(dbgCenter, dbgLeft, _
                                 dbgLeft, dbgCenter, _
                                 dbgLeft)
                                  
                                  
           GrdBusCliente.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
           GrdBusCliente.HeadLines = 0
           GrdBusCliente.MarqueeStyle = dbgNoMarquee


           arrCampos = Array("COD_CLIENTE", "DES_DIRECCION", "COD_TIPO_DIRECCION", "COD_UBIGEO", "COD_URBANIZACION", "DES_REFERENCIA_DIRECCION", "COD_SUFIJO_DIRECCION", "FLG_PRINCIPAL", "COD_DIRECCION_CLI")
           arrCaption = Array("Cliente", "Direccción", "Tipo", "Ubigeo", "Urbanizacion", "Referencia", "Sufijo", "Principal", "Item")
           arrAncho = Array(1000, 5000, 700, 800, 800, 800, 800, 800, 800)
           arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
           
           grdDirecciones.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
           grdDirecciones.HeadLines = 0
           grdDirecciones.MarqueeStyle = dbgNoMarquee
           grdDirecciones.Columns(0).Visible = False
           grdDirecciones.Columns(1).Visible = True
           grdDirecciones.Columns(2).Visible = False
           grdDirecciones.Columns(3).Visible = False
           grdDirecciones.Columns(4).Visible = False
           grdDirecciones.Columns(5).Visible = False
           grdDirecciones.Columns(6).Visible = False
           grdDirecciones.Columns(7).Visible = False
           grdDirecciones.Columns(8).Visible = False
           
           
End Sub
Public Property Get Urbanizacion() As String
    Urbanizacion = cboUrbanizacion.Text
End Property
Public Property Get Distrito() As String
    Distrito = cboDistrito.Text
End Property

Property Let LocalDespacho(mvar As String)
    cboDespacho.BoundText = mvar
End Property

Public Property Get LocalDespacho() As String
    LocalDespacho = cboDespacho.BoundText
End Property

Public Property Get CodDireccionCli() As String
    CodDireccionCli = strCodDireccionCli
End Property

Public Property Let CodDireccionCli(ByVal lstrCodDireccionCli As String)
    strCodDireccionCli = lstrCodDireccionCli
End Property

Public Property Get Ubigeo() As String
    Ubigeo = strUbigeo
End Property
Public Property Let Ubigeo(lstrUbigeo As String)
    strUbigeo = lstrUbigeo
End Property


''Property Get Ubigeo() As String
''    If Not rsCliente Is Nothing Then Ubigeo = "" & rsCliente("UBIGEO").Value
''End Property
''Property Let Ubigeo(lstrUbigeo As String)
''    strUbigeo = lstrUbigeo
''End Property



Public Property Get ConsultaPrecio() As Boolean
    ConsultaPrecio = bolConsultaPrecio
End Property

Public Property Let ConsultaPrecio(ByVal lbolConsultaPrecio As Boolean)
    bolConsultaPrecio = lbolConsultaPrecio
End Property


Public Sub CargaDatosCliente(ByVal Codigo As String, _
                            ByVal intClienteVerficado As Integer, _
                            ByVal strLocalAsignado As String, _
                            ByVal strFlgTipoJuridica As String, _
                            ByVal strDesRazonSocial As String, _
                            ByVal strDesNomComercial As String, _
                            ByVal strDesNomCliente As String, _
                            ByVal strDesApeCliente As String, _
                            ByVal strDesApe2Cliente As String, _
                            ByVal strCodLocalDespacho As String, _
                            ByVal strCodSufijo As String, _
                            ByVal strDireccSocial As String, _
                            ByVal strDireccComercial As String, _
                            ByVal strNumDocID As String, _
                            ByVal strCodDocID As String)

    Dim strUbigeo As String
    Dim strUrbanizacion As String
    Dim strCodDireccionCli As String
    Dim objLocal As New clsLocal
    
    Set cboLocal.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
    cboLocal.ListField = "local_dex2"
    cboLocal.BoundColumn = "COD_LOCAL"
    cboLocal.BoundText = objUsuario.CodigoLocal

'''    Set cboDespacho.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
'''    cboDespacho.ListField = "local_dex2"
'''    cboDespacho.BoundColumn = "COD_LOCAL"
'''    cboDespacho.BoundText = objUsuario.CodigoLocal

    Set cboSufijo.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_SUFIJO", 0)
    cboSufijo.ListField = "DES_ABREVIATURA"
    cboSufijo.BoundColumn = "COD_SUFIJO"
    cboSufijo.BoundText = strCodSufijo

    If Codigo = "" Then Exit Sub
    
    chkVerificado.Value = intClienteVerficado
    txtCodigo.Text = Codigo
    LocalAsignado = strLocalAsignado
    strFlagJuridico = strFlgTipoJuridica
    strRazonSocial = strDesRazonSocial
    strDireccionSocial = strDireccSocial
    strDireccionComercial = strDireccComercial
    strCodDocumentoID = strCodDocID
    strNumDocumentoID = strNumDocID
     
    If strFlagJuridico = "1" Then
        txtNombre.Text = strDesRazonSocial
        txtApellido.Text = strDesNomCliente
        optTipo(1).Value = True
        Label2.Caption = "Razón Social"
        Label3.Caption = "Razón Comercial"
    Else
        txtNombre.Text = strDesNomCliente
        txtApellido.Text = strDesApeCliente
        txtApeMaterno.Text = strDesApe2Cliente
        Label2.Caption = "Nombre"
        Label3.Caption = "Apellido"
        optTipo(0).Value = True
    End If
    
    '+++++ carca combo cia, jct
    Set cboCia.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCia.ListField = "Des"
    cboCia.BoundColumn = "Cod"
    'setea cia segun local de despacho
    sCia = "" & gclsOracle.FN_Valor("btlprod.pkg_local.fn_devuelve_cia", strCodLocalDespacho)
    If (sCia <> "") Then
        cboCia.BoundText = sCia
     Else
        cboCia.BoundText = objUsuario.CodigoEmpresa
    End If
    '+++++
    
    cboLocal.BoundText = strLocalAsignado
    cboDespacho.BoundText = strCodLocalDespacho

End Sub

'I.ECASTILLO 17.12.2020
Public Sub reasignaLocales(ByVal strLocalAsignado As String, _
                            ByVal strCodLocalDespacho As String)
On Error GoTo handle
    Dim rsCia As oraDynaset
    Dim Cia As String
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, strCodLocalDespacho)
    If (rsCia.RecordCount > 0) Then
      Cia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    cboCia.BoundText = Cia
    cboCia_Change
    cboLocal.BoundText = strLocalAsignado
    cboDespacho.BoundText = strCodLocalDespacho
    LocalAsignado = strLocalAsignado
    LocalDespacho = strCodLocalDespacho
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
Public Sub LimpiarLocalDespacho()
On Error GoTo handle
cboLocal.BoundText = "" 'estaba comentado
cboDespacho.BoundText = "" 'estaba comentado
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
'F.ECASTILLO
