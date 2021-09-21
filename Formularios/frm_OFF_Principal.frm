VERSION 5.00
Begin VB.Form frm_OFF_Principal 
   Caption         =   "Sistema de Contingencia de Ventas"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   Icon            =   "frm_OFF_Principal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   11775
      TabIndex        =   37
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton cmdSalirContingencia 
         Caption         =   "Finalizar Cont"
         Height          =   615
         Left            =   9585
         Picture         =   "frm_OFF_Principal.frx":1C9A2
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdProforma 
         Caption         =   "Proforma"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8520
         Picture         =   "frm_OFF_Principal.frx":1CF2C
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "&Búsqueda"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2130
         Picture         =   "frm_OFF_Principal.frx":1D4B6
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdModalidad 
         Caption         =   "&Modalidad"
         Enabled         =   0   'False
         Height          =   615
         Left            =   0
         Picture         =   "frm_OFF_Principal.frx":1F1B0
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdMantenimientos 
         Caption         =   "&Mantenimient"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1065
         Picture         =   "frm_OFF_Principal.frx":1F73A
         Style           =   1  'Graphical
         TabIndex        =   44
         Tag             =   "002"
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdDocumento 
         Caption         =   "&Documento"
         Height          =   615
         Left            =   5325
         Picture         =   "frm_OFF_Principal.frx":1FCC4
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdFormaPago 
         Caption         =   "&Forma Pago"
         Height          =   615
         Left            =   6390
         Picture         =   "frm_OFF_Principal.frx":2024E
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdAdministrador 
         Caption         =   "Administrador"
         Height          =   615
         Left            =   3195
         Picture         =   "frm_OFF_Principal.frx":207D8
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "001"
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   10650
         Picture         =   "frm_OFF_Principal.frx":20D62
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   4260
         Picture         =   "frm_OFF_Principal.frx":212EC
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         Width           =   1060
      End
      Begin VB.CommandButton cmdGrabaVenta 
         Caption         =   "&Graba Venta"
         Height          =   615
         Left            =   7455
         Picture         =   "frm_OFF_Principal.frx":21876
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Width           =   1060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   9765
         TabIndex        =   59
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   8940
         TabIndex        =   58
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1230
         TabIndex        =   57
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   7875
         TabIndex        =   56
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   4440
         TabIndex        =   55
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   2310
         TabIndex        =   54
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lblModalidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + Q"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   165
         TabIndex        =   53
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   10830
         TabIndex        =   52
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   3375
         TabIndex        =   51
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   6810
         TabIndex        =   50
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   5745
         TabIndex        =   49
         Top             =   600
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7365
      Left            =   7230
      ScaleHeight     =   7365
      ScaleWidth      =   4560
      TabIndex        =   6
      Top             =   0
      Width           =   4560
      Begin vbp_Ventas.ctlGrillaArray grdDetalleVenta 
         Height          =   4035
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   7117
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   0
         ScaleHeight     =   2745
         ScaleWidth      =   4470
         TabIndex        =   7
         Top             =   4500
         Width           =   4500
         Begin VB.Label lblRedondeo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   4005
            TabIndex        =   25
            Top             =   1440
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblVueltoRedondeado 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4005
            TabIndex        =   23
            Top             =   1725
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Redondeo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Visible         =   0   'False
            Width           =   4260
         End
         Begin VB.Label lblVuelto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   4005
            TabIndex        =   21
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "Vuelto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   4260
         End
         Begin VB.Label lblPagado 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   4005
            TabIndex        =   19
            Top             =   740
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Pagado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   18
            Top             =   740
            Width           =   4260
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   4005
            TabIndex        =   17
            Top             =   60
            Width           =   375
         End
         Begin VB.Label lblTotalaPagar 
            AutoSize        =   -1  'True
            Caption         =   "Total a Pagar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   120
            TabIndex        =   16
            Top             =   60
            Width           =   4260
         End
         Begin VB.Label lblcopago 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Left            =   4005
            TabIndex        =   14
            Top             =   400
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPctCopago 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Left            =   2385
            TabIndex        =   13
            Top             =   400
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Left            =   2880
            TabIndex        =   12
            Top             =   400
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   0
            X2              =   5160
            Y1              =   1380
            Y2              =   1380
         End
         Begin VB.Label lblTotalDolares 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   2805
            TabIndex        =   10
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label lblTipoCambio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   2805
            TabIndex        =   8
            Top             =   1995
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   4200
            MouseIcon       =   "frm_OFF_Principal.frx":21E00
            MousePointer    =   99  'Custom
            Picture         =   "frm_OFF_Principal.frx":21F52
            Top             =   2280
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgCalculadora 
            Height          =   240
            Left            =   3900
            MouseIcon       =   "frm_OFF_Principal.frx":224DC
            MousePointer    =   99  'Custom
            Picture         =   "frm_OFF_Principal.frx":2262E
            Top             =   2280
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   3600
            MousePointer    =   1  'Arrow
            Picture         =   "frm_OFF_Principal.frx":22BB8
            Stretch         =   -1  'True
            Tag             =   "Atencion de HelpDesk"
            Top             =   2280
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label10 
            BackColor       =   &H80000018&
            Caption         =   "Vuelto Redondeado === >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   120
            TabIndex        =   24
            Top             =   1725
            Width           =   4260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   120
            TabIndex        =   9
            Top             =   1995
            Width           =   3120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total en Dólares"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   2280
            Width           =   3120
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "Importe Co-Pago"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   400
            Visible         =   0   'False
            Width           =   4260
         End
      End
      Begin VB.Label lblSiguiente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   2280
         TabIndex        =   34
         Top             =   0
         Width           =   2115
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "VENTA REGULAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   30
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7365
      Left            =   0
      ScaleHeight     =   7365
      ScaleWidth      =   7185
      TabIndex        =   5
      Top             =   0
      Width           =   7190
      Begin vbp_Ventas.ctlGrillaArray grdAlternativo 
         Height          =   2115
         Left            =   120
         TabIndex        =   3
         Top             =   5160
         Width           =   7000
         _ExtentX        =   12356
         _ExtentY        =   3731
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlGrillaArray grdProducto 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   1020
         Width           =   7000
         _ExtentX        =   12356
         _ExtentY        =   6588
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlTextBox txtBuscar 
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   420
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         Tipo            =   2
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
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   4980
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F4"
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
         Left            =   2760
         TabIndex        =   36
         Top             =   4860
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Activar / Desactivar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   3120
         TabIndex        =   35
         Top             =   4875
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Productos Alternativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   420
         TabIndex        =   33
         Top             =   4875
         Width           =   2010
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F3"
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
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   4860
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   120
         Picture         =   "frm_OFF_Principal.frx":22EFA
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Búsqueda"
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
         Index           =   0
         Left            =   540
         TabIndex        =   31
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F1"
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
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   420
         Width           =   255
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
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   765
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Productos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   420
         TabIndex        =   28
         Top             =   780
         Width           =   1710
      End
   End
End
Attribute VB_Name = "frm_OFF_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsProductos As New ADODB.Recordset
Private rsAlternativos As New ADODB.Recordset

Private Const ColDescripcion As Integer = 0
Private Const ColLaboratorio As Integer = 1
Private Const ColPrecioPublico As Integer = 2
Private Const ColCodigo As Integer = 3
Private Const ColAsoSustituto As Integer = 4

Public Sub cmdAdministrador_Click()
On Error GoTo handle
    
    frm_OFF_ConsultaDocumento.Show vbModal

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdBuscar_Click()
Dim strBuscar As String
Dim objProducto As cls_OFF_Producto
'Dim cnn As ADODB.Connection
Dim strSql As String
'Dim rsProductos As New ADODB.Recordset
Dim aTmp As Variant
    
On Error GoTo handle
    
    strBuscar = Trim(txtBuscar.Text)

    If strBuscar = "" Then
        MsgBox "Debe ingresar el criterio para la busqueda", vbCritical, App.ProductName
        Exit Sub
    End If
    
    If Len(strBuscar) < 3 Then
        MsgBox "Debe ingresar al menos 3 caracteres", vbCritical, App.ProductName
        Exit Sub
    End If
    
    
    grdProducto.Limpiar
    
    'Set cnn = New ADODB.Connection
    'cnn.Open gstrConexion
    
    
    'strSql = "select * from precios.txt"
    'rsProductos.Open strSql, gstrConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    
    If IsNumeric(strBuscar) And Len(Trim(strBuscar)) = 5 Then
        rsProductos.Filter = "CODIGO = '" & strBuscar & "'"
    ElseIf IsNumeric(strBuscar) And Len(Trim(strBuscar)) >= 8 Then
        rsProductos.Filter = "CODIGO_BARRA = '" & strBuscar & "'"
    Else
        rsProductos.Filter = "DESCRIPCION LIKE '" & Trim(strBuscar) & "%'"
    End If
    
    Select Case rsProductos.RecordCount
        Case 0
            grdProducto.Array1 = New XArrayDB
            MsgBox "No se han encontrado los datos buscados", vbCritical, App.ProductName
            txtBuscar.selection
            Exit Sub
            
        Case 1
            Set objProducto = New cls_OFF_Producto
            
            With rsProductos
                grdProducto.Array1 = objProducto.AgregaProducto(.Fields("CODIGO"), _
                            Mid(UCase(.Fields("DESCRIPCION")), 1, 1) & Mid(LCase(.Fields("DESCRIPCION")), 2), _
                            .Fields("LABORATORIO"), _
                            .Fields("FLG_FRACCIONA"), _
                            .Fields("CTD_FRACCIONA"), _
                            .Fields("MONEDA"), _
                            .Fields("PRECIO_PUBLICO"), _
                            .Fields("CON_RECETA") & "", _
                            IIf(IsNull(.Fields("CODIGO_BARRA")), "", .Fields("CODIGO_BARRA")), _
                            IIf(IsNull(.Fields("PARTIDA_ARANCELARIA")), "", .Fields("PARTIDA_ARANCELARIA")), _
                            .Fields("PCT_IGV"), _
                            IIf(IsNull(.Fields("ASO_SUSTITUTO")), "", .Fields("ASO_SUSTITUTO")), _
                            .Fields("STOCK"), _
                            .Fields("DESCRIPCION_CORTA"), _
                            .Fields("PCT_DESCUENTO"), _
                            .Fields("FLG_REGALO"))
            End With
'            MsgBox grdAlternativo.Array1.Count(1)
            Set objProducto = Nothing
            
        Case Else
            grdProducto.Array1 = New XArrayDB
            grdProducto.Array1.LoadRows (rsProductos.GetRows(rsProductos.RecordCount, 0))
    End Select
    rsProductos.MoveFirst
    rsProductos.Filter = adFilterNone
    grdProducto.MoveFirst
    grdProducto.Rebind
    grdProducto.SetFocus


    
'    Set objProducto = New cls_OFF_Producto
'
'    With rsProductos
'
'        If (.BOF And .EOF) Then
'            MsgBox "No se han encontrado los datos buscados", vbCritical, App.ProductName
'            grdAlternativo.Limpiar
'        Else
'
'
'
'
'            .MoveFirst
'            Do While Not .EOF
'
'
'                grdProducto.Array1 = objProducto.AgregaProducto(.Fields("CODIGO"), _
'                            Mid(UCase(.Fields("DESCRIPCION")), 1, 1) & Mid(LCase(.Fields("DESCRIPCION")), 2), _
'                            .Fields("LABORATORIO"), _
'                            .Fields("FLG_FRACCIONA"), _
'                            .Fields("CTD_FRACCIONA"), _
'                            .Fields("MONEDA"), _
'                            .Fields("PRECIO_PUBLICO"), _
'                            .Fields("CON_RECETA"), _
'                            IIf(IsNull(.Fields("CODIGO_BARRA")), "", .Fields("CODIGO_BARRA")), _
'                            IIf(IsNull(.Fields("PARTIDA_ARANCELARIA")), "", .Fields("PARTIDA_ARANCELARIA")), _
'                            .Fields("PCT_IGV"), _
'                            IIf(IsNull(.Fields("ASO_SUSTITUTO")), "", .Fields("ASO_SUSTITUTO")), _
'                            .Fields("STOCK"), _
'                            .Fields("DESCRIPCION_CORTA"), _
'                            .Fields("PCT_DESCUENTO"), _
'                            .Fields("FLG_REGALO"))
'
'
'                .MoveNext
'            Loop
'
'
'            grdProducto.Rebind
'            grdProducto.SetFocus
'        End If
'
'
'    End With
'
'    Set objProducto = Nothing
    'rsProductos.Close
    'cnn.Close

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub





Public Sub cmdBusqueda_Click()

End Sub

Public Sub cmdCancelar_Click()
On Error GoTo handle
   If MsgBox("¿ Desea borrar todos los Datos del Documento ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        Inicio
   End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Public Sub cmdDocumento_Click()
    frm_OFF_Documento.Show vbModal
End Sub

Public Sub cmdFormaPago_Click()
    frm_OFF_FormaPago.Show vbModal
End Sub

Public Sub cmdGrabaVenta_Click()
On Error GoTo CtrlErr
    
    Dim objDocumento As cls_OFF_Documento
    Dim xTemp As New XArrayDB
    Dim i As Integer
    Dim varMsgDoc As Variant
    Dim strUltDocEmi As String
    
    'validamos los productos vendidos
    If objOFFVenta.DetalleVenta.Count(1) < 1 Then
        MsgBox "Debe ingresar Productos", vbCritical + vbOKOnly, App.ProductName
        Exit Sub
    End If
    
    'validamos la forma de pago
    If objOFFVenta.PagoVenta.Count(1) < 1 Then
        MsgBox "Debe ingresar la Forma de Pago", vbCritical + vbOKOnly, App.ProductName
        frm_OFF_FormaPago.Show vbModal
        Exit Sub
    End If
    
    'Validar monto de pagos con tarjetas
    Dim dTotalConTarjeta As Double
    dTotalConTarjeta = 0
    For i = 0 To objOFFVenta.PagoVenta.UpperBound(1)
        If objOFFVenta.PagoVenta.Value(i, objOFFVenta.ColDPCodPago) = "2" Then
            dTotalConTarjeta = dTotalConTarjeta + objOFFVenta.PagoVenta.Value(i, objOFFVenta.ColDPMtoSoles)
        End If
    Next
    If dTotalConTarjeta > objOFFVenta.TotalVenta Then
        MsgBox "El monto total del pago con tarjeta no puede ser mayor al importe total de la venta.", vbCritical, App.ProductName
        frm_OFF_FormaPago.Show vbModal 'ARTURO ESCATE
        Exit Sub
    End If
    
    objOFFVenta.LimpiaDocumentosGenerados
    objOFFVenta.Graba
    
    On Error GoTo 0
    On Error GoTo CtrlImp
    
    Set xTemp = objOFFVenta.DocumentosGenerados
    
    Set objDocumento = New cls_OFF_Documento
    
    For i = 0 To xTemp.UpperBound(1)
        varMsgDoc = varMsgDoc & xTemp(i, 0) & " " & xTemp(i, 1) & Chr(13)
    Next i
    
    MsgBox "Se realizó la transacción satisfactoriamente  - " & Chr(13) & Chr(13) & varMsgDoc, vbInformation + vbOKOnly, App.ProductName
    
    objOFFUsuario.ActualizaUltDocEmitido
    
    For i = 0 To xTemp.UpperBound(1)
        
        strUltDocEmi = objOFFUsuario.UltDocEmi
        
        If strUltDocEmi <> xTemp(i, 0) Then
                    MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                            "en posición de " & xTemp(i, 0), vbInformation, App.ProductName
        End If

        objDocumento.ImprimePorDocumento xTemp(i, 0), xTemp(i, 1)
        
        objOFFUsuario.ActualizaUltDocEmitido xTemp(i, 0)
        
        objOFFVenta.TipoDocumento = xTemp(i, 0)
        objOFFVenta.NumDocumento = xTemp(i, 1)
        
    Next i
    
    Set objDocumento = Nothing
    
    Inicio
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    Exit Sub
CtrlImp:

    MsgBox Err.Description, vbCritical, App.ProductName
    Inicio
    
End Sub

Public Sub cmdImpresion_Click()

End Sub

Public Sub cmdMantenimientos_Click()
    frm_OFF_Mantenimiento.Show vbModal
End Sub

Public Sub cmdModalidad_Click()

End Sub

Public Sub cmdProforma_Click()

End Sub

Public Sub cmdSalir_Click()
        
    
        
    Unload Me

End Sub

Private Sub cmdSalirContingencia_Click()
    Dim objIni As cls_ArchivoIni
   On Error GoTo Control
    
    If MsgBox("Desea salir del módulo de contingencia ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    
    Set objIni = New cls_ArchivoIni
    objIni.GuardarIni gstrIni, "general", "FLG_CONTINGENCIA", "0"
    
    Set objIni = Nothing
    
    Unload Me

   Exit Sub

Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub Form_Activate()


On Error GoTo CtrlErr
    cmdMantenimientos.Enabled = False
    If objOFFUsuario.CodigoPerfil = COD_PERFIL_QFI Or objOFFUsuario.CodigoPerfil = COD_PERFIL_QFII Then
        cmdMantenimientos.Enabled = True
    End If
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpCtrl As Boolean, tmpAlt As Boolean

On Error GoTo handle

    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0

    Select Case KeyCode
        Case vbKeyF1
            txtBuscar.SetFocus
        Case vbKeyF2
            grdProducto.SetFocus
        Case vbKeyF3
            If grdAlternativo.Visible Then
                grdAlternativo.SetFocus
            End If
        Case vbKeyF4
            mostrarAlternativos
        Case vbKeyF5
            cmdDocumento_Click
        Case vbKeyF6
            cmdFormaPago_Click
        Case vbKeyF7
            cmdGrabaVenta_Click
        Case vbKeyF12
            grdDetalleVenta.SetFocus
        Case tmpCtrl And vbKeyF
            cmdSalir_Click
        Case tmpCtrl And vbKeyX
            cmdSalirContingencia_Click
        Case tmpCtrl And vbKeyM
            If cmdMantenimientos.Enabled Then cmdMantenimientos_Click
            SendKeys "{BKSP}"
        Case tmpCtrl And vbKeyD
            cmdAdministrador_Click
            SendKeys "{BKSP}"
        Case tmpCtrl And vbKeyC
            cmdCancelar_Click
            SendKeys "{BKSP}"
    End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()
On Error GoTo handle

    grdAlternativo.Visible = False

    CargarProductos
    CargarAlternativos
    
    Me.Caption = Me.Caption & "-" & objOFFUsuario.NombreUsuario & "-" & objOFFUsuario.CodLocal & "- VERSION " & App.Major & App.Minor & "." & App.Revision
    
    SetGrid
    
    lblTipoCambio.Caption = Format(Val(objOFFUsuario.TipoCambio & ""), "###,###.00")
    
    objOFFVenta.TipoDocumento = objOFFVenta.CodDocDefault
    objOFFVenta.NumDocumento = objOFFVenta.NumDocDefault
    Call Siguiente
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub



Private Sub Form_Unload(Cancel As Integer)

    rsAlternativos.Close
    Set rsAlternativos = Nothing
    
    rsProductos.Close
    Set rsProductos = Nothing

    Dim X As Integer
    For X = (Forms.Count - 1) To 0 Step -1
        If TypeOf Forms(X) Is Form Then
            Unload Forms(X)
        End If
    Next X
    
End Sub

Private Sub grdAlternativo_DblClick()
On Error GoTo handle
    
    If grdAlternativo.Columns(12).Value = "0" Then
        MsgBox "El producto no tiene Stock", vbCritical + vbOKOnly, App.ProductName
        grdAlternativo.SetFocus
        Exit Sub
    End If
    
    MostrarCantidadProducto "grdAlternativo"

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub grdAlternativo_KeyPress(KeyAscii As Integer)
On Error GoTo handle

    If KeyAscii = vbKeyReturn Then
        grdAlternativo_DblClick
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdDetalleVenta_DblClick()
On Error GoTo handle
    
    MostrarCantidadProducto "grdDetalleVenta"

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdDetalleVenta_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case vbKeyDelete
            On Error GoTo CtrlErr
            If grdDetalleVenta.ApproxCount = 0 Then Exit Sub
            grdDetalleVenta.Delete
            frm_OFF_Principal.MostrarTotales
CtrlErr:
            On Error GoTo 0
    End Select

End Sub

Private Sub grdDetalleVenta_KeyPress(KeyAscii As Integer)
On Error GoTo handle

    If KeyAscii = vbKeyReturn Then
        grdDetalleVenta_DblClick
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdProducto_DblClick()
On Error GoTo handle
    
    If grdProducto.Columns(12).Value = "0" Then
        MsgBox "El producto no tiene Stock", vbCritical + vbOKOnly, App.ProductName
        grdProducto.SetFocus
        Exit Sub
    End If
    
    MostrarCantidadProducto "grdProducto"
    
    txtBuscar.SetFocus
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdProducto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handle

    If KeyCode = vbKeyReturn Then
        Call grdProducto_DblClick
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
'COMENTADO POR PHERRERA 10/09/08
'Private Sub grdProducto_KeyPress(KeyAscii As Integer)
'On Error GoTo handle
'
'    If KeyAscii = vbKeyReturn Then
'        Call grdProducto_DblClick
'    End If
'
'    Exit Sub
'handle:
'    MsgBox Err.Description, vbCritical, App.ProductName
'End Sub

Private Sub grdProducto_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo CtrlErr
    
    If grdProducto.ApproxCount > 0 Then
        If Not IsNull(grdProducto.Columns(0).Value) _
            And Not IsNull(grdProducto.Columns(0).Value) Then _
                Call BuscarAlternativo(grdProducto.Columns(0).Value, grdProducto.Columns(11).Value)
    Else
        grdAlternativo.Limpiar
    End If

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub





Private Sub txtBuscar_KeyPress(KeyAscii As Integer)


On Error GoTo handle
    
    
    
    If KeyAscii = vbKeyReturn Then
        cmdBuscar_Click
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Private Sub MostrarCantidadProducto(ByVal lstrOrigen As String)


On Error GoTo handle


    Select Case UCase(lstrOrigen)
        
        Case UCase("grdProducto")
            
            If grdProducto.ApproxCount > 0 Then
                If Not IsNull(grdProducto.Columns(0).Value) Then
                    frm_OFF_CantidadProducto.Datos Format(grdProducto.Columns(0).Value, "00000"), _
                                 grdProducto.Columns(1).Value, _
                                 grdProducto.Columns(14).Value, _
                                 grdProducto.Columns(6).Value, _
                                 grdProducto.Columns(15).Value, _
                                 grdProducto.Columns(10).Value, _
                                 grdProducto.Columns(9).Value, _
                                 grdProducto.Columns(4).Value, _
                                 grdProducto.Columns(13).Value, _
                                 grdProducto.Columns(7).Value, _
                                 grdProducto.Columns(3).Value
                End If
            End If
        Case UCase("grdAlternativo")
            If grdAlternativo.ApproxCount > 0 Then
                If Not IsNull(grdAlternativo.Columns(0).Value) Then
                    frm_OFF_CantidadProducto.Datos Format(grdAlternativo.Columns(0).Value, "00000"), _
                                 grdAlternativo.Columns(1).Value, _
                                 grdAlternativo.Columns(14).Value, _
                                 grdAlternativo.Columns(6).Value, _
                                 grdAlternativo.Columns(15).Value, _
                                 grdAlternativo.Columns(10).Value, _
                                 grdAlternativo.Columns(9).Value, _
                                 grdAlternativo.Columns(4).Value, _
                                 grdAlternativo.Columns(13).Value, _
                                 grdAlternativo.Columns(7).Value, _
                                 grdAlternativo.Columns(3).Value
                End If
            End If
        Case UCase("grdDetalleVenta")
            If grdDetalleVenta.ApproxCount > 0 Then
                If Not IsNull(grdDetalleVenta.Columns(0).Value) Then
                    frm_OFF_CantidadProducto.Datos grdDetalleVenta.Columns(0).Value, _
                             grdDetalleVenta.Columns(1).Value, _
                             grdDetalleVenta.Columns(4).Value, _
                             grdDetalleVenta.Columns(17).Value, _
                             grdDetalleVenta.Columns(11).Value, _
                             grdDetalleVenta.Columns(12).Value, _
                             grdDetalleVenta.Columns(13).Value, _
                             grdDetalleVenta.Columns(14).Value, _
                             grdDetalleVenta.Columns(16).Value, _
                             grdDetalleVenta.Columns(18).Value, _
                             grdDetalleVenta.Columns(19).Value
                End If
            End If
    End Select
    


    
    
    MostrarTotales
    

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub



Private Sub SetGrid()
On Error GoTo handle

    Dim arrCampos As Variant
    Dim arrTitulos As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim intAncho As Integer
    Dim i As Integer
    Dim columna As TrueDBGrid70.Column


    intAncho = 135
    
    'Detalle Venta
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    arrTitulos = Array("Código", "Descripción", "F", "Cant", "PctDcto", "PrcUnit", "PrcOrig", "MtoIGV", "MtoExon", "Precio", "FlgPrc", "Regalo", "PctIgv", "PartAranc", "CtdFracción", "UsuModPrecio", "DescripciónCorta", "PrecioPúblico", "ConReceta", "FlgFracciona")
    arrAncho = Array(5 * intAncho, 18 * intAncho, 2 * intAncho, 3 * intAncho, 5 * intAncho, 5 * intAncho, 7 * intAncho, 6 * intAncho, 7 * intAncho, 5 * intAncho, 6 * intAncho, 6 * intAncho, 6 * intAncho, 6 * intAncho, 6 * intAncho, 6 * intAncho, 18 * intAncho, 0, 0, 0)
    arrAlineacion = Array(0, 0, 1, 0, 1, 1, 1, 1, 1, 1, 0, 0, 1, 0, 1, 1, 0, 0, 0, 0)
    grdDetalleVenta.FormatoGrilla arrCampos, arrTitulos, arrAncho, arrAlineacion
    
    For i = 0 To objOFFVenta.ColDetalleVenta
        grdDetalleVenta.Columns(i).AllowSizing = False
    Next i
    grdDetalleVenta.Columns(objOFFVenta.ColDVMtoSubtotal).NumberFormat = "##0.00"
    grdDetalleVenta.Columns(objOFFVenta.ColDVPctDescuento).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVPrcUnitario).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVPrcOriginal).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVMtoIgv).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVMtoExonerado).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVDescripcionCorta).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVPrecioPublico).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVConReceta).Visible = False
    grdDetalleVenta.Columns(objOFFVenta.ColDVFlgFracciona).Visible = False


    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    arrTitulos = Array("Código", "Producto", "Línea", "FlgFracciona", "CtdFracciona", "Moneda", "Precio Final", "ConReceta", "CódigoBarra", "PartidaArancelaria", "PctIGV", "AsoSustituto", "Stock", "DescripciónCorta", "PctDescuento", "FlgRegalo")
    arrAncho = Array(1000, 4100, 800, 0, 0, 0, 900, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgRight, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    grdProducto.FormatoGrilla arrCampos, arrTitulos, arrAncho, arrAlineacion

    For Each columna In grdProducto.Columns
        columna.AllowSizing = False
        columna.Visible = False
    Next
    grdProducto.Columns(0).Visible = False
    grdProducto.Columns(0).NumberFormat = "00000"
    grdProducto.Columns(1).Visible = True
    grdProducto.Columns(2).Visible = True
    grdProducto.Columns(6).Visible = True
    grdProducto.Columns(6).NumberFormat = "##0.00"
    'grdProducto.RowDividerStyle = 0
    grdProducto.MarqueeStyle = dbgHighlightRow
    'grdProducto.CambiaSeleccionadoBackColor &H800000
    'grdProducto.CambiaSeleccionadoForeColor &HFFFFFF
    
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    arrTitulos = Array("Código", "Producto", "Línea", "FlgFracciona", "CtdFracciona", "Moneda", "Precio Final", "ConReceta", "CódigoBarra", "PartidaArancelaria", "PctIGV", "AsoSustituto", "Stock", "DescripciónCorta", "PctDescuento", "FlgRegalo")
    arrAncho = Array(1000, 4100, 800, 0, 0, 0, 900, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgRight, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    grdAlternativo.FormatoGrilla arrCampos, arrTitulos, arrAncho, arrAlineacion
    For Each columna In grdAlternativo.Columns
        columna.AllowSizing = False
        columna.Visible = False
    Next
    grdAlternativo.Columns(0).Visible = False
    grdAlternativo.Columns(0).NumberFormat = "00000"
    grdAlternativo.Columns(1).Visible = True
    grdAlternativo.Columns(2).Visible = True
    grdAlternativo.Columns(6).Visible = True
    grdAlternativo.Columns(6).NumberFormat = "##0.00"
    'grdAlternativo.RowDividerStyle = 0
    grdAlternativo.MarqueeStyle = dbgHighlightRow
    'grdAlternativo.CambiaSeleccionadoBackColor &H800000
    'grdAlternativo.CambiaSeleccionadoForeColor &HFFFFFF
    
    



Exit Sub
handle:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub

Public Sub MostrarTotales()
On Error GoTo handle

    objOFFVenta.Totales
    lblTotal.Caption = Format(objOFFVenta.TotalVenta, "###,###0.00")
    lblPagado.Caption = Format(objOFFVenta.TotalPago, "###,###0.00")
    lblVuelto.Caption = Format(objOFFVenta.TotalVuelto, "###,###0.00")
    lblTotalDolares.Caption = Format(objOFFVenta.TotalDolares, "###,###0.00")
    lblVueltoRedondeado.Caption = Format(objOFFVenta.TotalVuelto, "###,###0.00")
Exit Sub
handle:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub


Private Sub BuscarAlternativo(ByVal pstrCodigo As String, _
                              ByVal pstrAsoSustituto As String)
On Error GoTo CtrlErr

    Dim strBuscar As String
    Dim objProducto As cls_OFF_Producto

    If pstrAsoSustituto = "" Then Exit Sub
    If Not grdAlternativo.Visible Then Exit Sub
    
    strBuscar = "(CODIGO <> '" & pstrCodigo & "') AND (ASO_SUSTITUTO = '" & pstrAsoSustituto & "')"
    rsAlternativos.MoveFirst
    rsAlternativos.Filter = strBuscar

'    MsgBox rsAlternativos.RecordCount
    
    Select Case rsAlternativos.RecordCount
        Case 0
            grdAlternativo.Array1 = New XArrayDB
        Case 1
            Set objProducto = New cls_OFF_Producto
            
            With rsAlternativos
                grdAlternativo.Array1 = objProducto.AgregaProducto(.Fields("CODIGO"), _
                            Mid(UCase(.Fields("DESCRIPCION")), 1, 1) & Mid(LCase(.Fields("DESCRIPCION")), 2), _
                            .Fields("LABORATORIO"), _
                            .Fields("FLG_FRACCIONA"), _
                            .Fields("CTD_FRACCIONA"), _
                            .Fields("MONEDA"), _
                            .Fields("PRECIO_PUBLICO"), _
                            .Fields("CON_RECETA") & "", _
                            IIf(IsNull(.Fields("CODIGO_BARRA")), "", .Fields("CODIGO_BARRA")), _
                            IIf(IsNull(.Fields("PARTIDA_ARANCELARIA")), "", .Fields("PARTIDA_ARANCELARIA")), _
                            .Fields("PCT_IGV"), _
                            IIf(IsNull(.Fields("ASO_SUSTITUTO")), "", .Fields("ASO_SUSTITUTO")), _
                            .Fields("STOCK"), _
                            .Fields("DESCRIPCION_CORTA"), _
                            .Fields("PCT_DESCUENTO"), _
                            .Fields("FLG_REGALO"))
            End With
'            MsgBox grdAlternativo.Array1.Count(1)
            Set objProducto = Nothing
            
        Case Else
            grdAlternativo.Array1 = New XArrayDB
            grdAlternativo.Array1.LoadRows (rsAlternativos.GetRows(rsAlternativos.RecordCount, 0))
    End Select

    grdAlternativo.Rebind
    rsAlternativos.Filter = adFilterNone

'    If rsAlternativos.RecordCount > 0 Then
'        grdAlternativo.Array1 = New XArrayDB
'        grdAlternativo.Array1.LoadRows (rsAlternativos.GetRows(rsAlternativos.RecordCount, 0))
'    Else
'        grdAlternativo.Array1 = New XArrayDB
'        'grdAlternativo.Array1.ReDim -1, -1, -1, -1
'    End If
'    MsgBox grdAlternativo.Array1.Count(1)

'    strBuscar = "ASO_SUSTITUTO = '" & pstrAsoSustituto & "'"
'    rsAlternativos.MoveFirst
'    rsAlternativos.Filter = strBuscar
'    rsAlternativos.Find (strBuscar)
'    rsAlternativos.Filter = "CODIGO <> '" & pstrCodigo & "'"

'    With rsAlternativos
'        While Not .EOF
'            If .Fields("ASO_SUSTITUTO") = pstrAsoSustituto And .Fields("CODIGO") <> pstrCodigo Then
'                grdAlternativo.Array1 = objProducto.AgregaProducto(.Fields("CODIGO"), _
'                            Mid(UCase(.Fields("DESCRIPCION")), 1, 1) & Mid(LCase(.Fields("DESCRIPCION")), 2), _
'                            .Fields("LABORATORIO"), _
'                            .Fields("FLG_FRACCIONA"), _
'                            .Fields("CTD_FRACCIONA"), _
'                            .Fields("MONEDA"), _
'                            .Fields("PRECIO_PUBLICO"), _
'                            .Fields("CON_RECETA") & "", _
'                            IIf(IsNull(.Fields("CODIGO_BARRA")), "", .Fields("CODIGO_BARRA")), _
'                            IIf(IsNull(.Fields("PARTIDA_ARANCELARIA")), "", .Fields("PARTIDA_ARANCELARIA")), _
'                            .Fields("PCT_IGV"), _
'                            IIf(IsNull(.Fields("ASO_SUSTITUTO")), "", .Fields("ASO_SUSTITUTO")), _
'                            .Fields("STOCK"), _
'                            .Fields("DESCRIPCION_CORTA"), _
'                            .Fields("PCT_DESCUENTO"), _
'                            .Fields("FLG_REGALO"))
'            End If
'            .MoveNext
'        Wend
'        grdAlternativo.Rebind
'    End With

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Public Sub Siguiente()
Dim objDocumento As cls_OFF_Documento

On Error GoTo CtrlErr
    
    Set objDocumento = New cls_OFF_Documento
    objOFFVenta.NumDocDefault = objDocumento.UltimoCorrelativo(objOFFVenta.TipoDocumento)
    Set objDocumento = Nothing
    
    lblSiguiente.Caption = objOFFVenta.TipoDocumento & " - " & objOFFVenta.NumDocDefault
    
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


Public Sub Inicio()
On Error GoTo CtrlErr
    objOFFCliente.Limpia
    grdProducto.Limpiar
    grdAlternativo.Limpiar
    grdDetalleVenta.Limpiar
    objOFFVenta.LimpiaPago
    objOFFVenta.LimpiaDetalle
    frm_OFF_FormaPago.grdListaFP.Limpiar
    frm_OFF_Documento.Limpia
    MostrarTotales
    objOFFVenta.TipoDocumento = objOFFVenta.CodDocDefault
    Siguiente
    txtBuscar.Text = ""
    txtBuscar.SetFocus
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName


End Sub

Private Sub CargarProductos()
On Error GoTo CtrlErr

    Dim objProducto As New cls_OFF_Producto

    Set rsProductos = objProducto.LeerTxtProductos
    
    Set objProducto = Nothing

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub CargarAlternativos()
On Error GoTo CtrlErr

    Dim objProducto As New cls_OFF_Producto

    Set rsAlternativos = objProducto.LeerTxtAlternativos
    
    Set objProducto = Nothing

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub mostrarAlternativos()

    grdAlternativo.Visible = Not grdAlternativo.Visible
    
    If grdAlternativo.Visible Then
        If grdProducto.ApproxCount > 0 Then
            If Not IsNull(grdProducto.Columns(0).Value) Then
                Call BuscarAlternativo(grdProducto.Columns(0).Value, grdProducto.Columns(11).Value)
            End If
        Else
            grdAlternativo.Limpiar
        End If
    End If

End Sub
