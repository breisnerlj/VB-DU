VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPedido 
   BorderStyle     =   0  'None
   ClientHeight    =   8370
   ClientLeft      =   7320
   ClientTop       =   150
   ClientWidth     =   4740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPedido.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPedido.frx":05CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPedido.frx":0B63
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPedido.frx":0F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPedido.frx":14D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin vbp_Ventas.FormDragger FormDragger1 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   503
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000004&
      Height          =   3405
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   4965
      Width           =   4740
      Begin VB.Label lblPuntosRed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
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
         Left            =   4065
         TabIndex        =   53
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Puntos Redimidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   52
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lblPuntosAcum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   4065
         TabIndex        =   51
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puntos Acumulados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   2085
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "<F11>PUNTOS"
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
         Left            =   3240
         TabIndex        =   49
         Top             =   3000
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblTotalDescuento 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   270
         Left            =   4065
         TabIndex        =   43
         Top             =   0
         Width           =   420
      End
      Begin VB.Label lblDescuento 
         AutoSize        =   -1  'True
         Caption         =   "Total Descuento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label lblTotalPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   3360
         TabIndex        =   2
         Top             =   2385
         Width           =   1140
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000018&
         Caption         =   "Vuelto Redondeado ====== >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   3540
      End
      Begin VB.Image imgHelpDesk 
         Height          =   240
         Left            =   3240
         MouseIcon       =   "frmPedido.frx":18B5
         MousePointer    =   99  'Custom
         Picture         =   "frmPedido.frx":1A07
         ToolTipText     =   "Mensaje a Help Desk (Ctrl+H)"
         Top             =   3240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCalculadora 
         Height          =   240
         Left            =   4320
         MouseIcon       =   "frmPedido.frx":1E30
         MousePointer    =   99  'Custom
         Picture         =   "frmPedido.frx":1F82
         ToolTipText     =   "Calculadora"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3600
         MouseIcon       =   "frmPedido.frx":250C
         MousePointer    =   99  'Custom
         Picture         =   "frmPedido.frx":265E
         ToolTipText     =   "Créditos"
         Top             =   3240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   270
         Left            =   2760
         TabIndex        =   24
         Top             =   2685
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   23
         Top             =   2685
         Width           =   1665
      End
      Begin VB.Label lblTotalDolares 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   270
         Left            =   2760
         TabIndex        =   22
         Top             =   2985
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total en Dolares"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   2985
         Width           =   1755
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   5160
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2880
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblPctCopago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   240
         Left            =   2400
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblcopago 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   4065
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Co-Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblTotalaPagar 
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   4065
         TabIndex        =   9
         Top             =   615
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   1215
         Width           =   795
      End
      Begin VB.Label lblPagado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   4065
         TabIndex        =   7
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vuelto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label lblVuelto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
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
         Left            =   4065
         TabIndex        =   5
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Redondeo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   0
         TabIndex        =   4
         Top             =   3225
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblRedondeo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   3945
         TabIndex        =   3
         Top             =   3225
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin vbp_Ventas.ctlGrillaArray grdPedido 
      Height          =   2775
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4895
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdPrecio 
      Caption         =   "Cambi&o Precio"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Tag             =   "018"
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdProbar 
      Caption         =   "PROBAR"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Tag             =   "XXX"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1095
      Left            =   0
      TabIndex        =   26
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vale Prom. Mes"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1245
         TabIndex        =   34
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Vale Prom. Dia"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Atencio.  Dia"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2490
         TabIndex        =   32
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atencio. Mes"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3645
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   2400
         X2              =   2400
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Image ImgNumClienT 
         Height          =   435
         Left            =   3840
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblNumClienteT 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4035
         TabIndex        =   30
         Top             =   840
         Width           =   105
      End
      Begin VB.Label lblNumClienteD 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2865
         TabIndex        =   29
         Top             =   840
         Width           =   105
      End
      Begin VB.Image ImgNumClienD 
         Height          =   435
         Left            =   2680
         Top             =   120
         Width           =   480
      End
      Begin VB.Image ImgValePromD 
         Height          =   435
         Left            =   360
         Top             =   120
         Width           =   465
      End
      Begin VB.Label lblValePromD 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "S/. 0.00"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   375
         TabIndex        =   28
         Top             =   840
         Width           =   465
      End
      Begin VB.Label lblValePromT 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/. 0.00"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1530
         TabIndex        =   27
         Top             =   840
         Width           =   465
      End
      Begin VB.Image ImgValePromT 
         Height          =   435
         Left            =   1520
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.OptionButton optEfectivo 
      Caption         =   "Efectivo (Alt+E)"
      Height          =   255
      Left            =   240
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optCredito 
      Caption         =   "Tarjeta (Alt+T)"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   1575
   End
   Begin vbp_Ventas.ctlDataCombo dbcTpTarjeta 
      Height          =   315
      Left            =   2160
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2505
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.OptionButton optNinguna 
      Caption         =   "Ninguna"
      Height          =   195
      Left            =   2640
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2280
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtCMPBus 
      Height          =   255
      Left            =   600
      TabIndex        =   46
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Tipo            =   3
      MaxLength       =   10
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
   Begin VB.Label lblDniInvalido 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "DNI inválido, No Aplica Prog. Atención al Cliente"
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
      Left            =   0
      TabIndex        =   48
      Top             =   3360
      Visible         =   0   'False
      Width           =   4185
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "CMP"
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
      TabIndex        =   47
      Top             =   3960
      Width           =   405
   End
   Begin VB.Label lblMedico 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1335
      TabIndex        =   45
      Top             =   3960
      Width           =   2520
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "<Ctrl+L>"
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
      Left            =   3840
      TabIndex        =   44
      Top             =   3975
      Width           =   720
   End
   Begin VB.Label Label14 
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
      TabIndex        =   41
      Top             =   0
      Width           =   390
   End
   Begin VB.Label lblCtrl 
      AutoSize        =   -1  'True
      Caption         =   "<Ctrl+S>"
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
      Left            =   3840
      TabIndex        =   36
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lbl_Cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   3660
      Width           =   3495
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Caption         =   "VERSIÓN DE PRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   4755
   End
   Begin VB.Label lblSiguiente 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   2520
      TabIndex        =   20
      Top             =   360
      Width           =   2175
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
      Left            =   120
      TabIndex        =   12
      Top             =   337
      Width           =   390
   End
   Begin VB.Label lblModalidad 
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
      Left            =   600
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public variables used elsewhere to set values for this form's position
'and size.
' una prueba
Dim lFloatingWidth As Long
Dim lFloatingHeight As Long
Dim lFloatingLeft As Long
Dim lFloatingTop As Long
Dim bMoving As Boolean

'Private variables used to track moving/sizing etc.
Public bDocked As Boolean
Public lDockedWidth As Long
Public lDockedHeight As Long

Public pxdbDatos As New XArrayDB
Public flgF6 As Integer

Dim dblPrecio As Double

Public pstrProd As String
Public pstrCant As String
Public pstrCantFrac As String
Public pstrCtdFracc As String
Public pstrPctUnit As String
Public pstrPrcUniKairo As String
Public pstrImpuesto As String
Public pstrComision As String
Public pstrPromocion As String
Public pstrPrecio As String
Public pstrSubTot As String
Public pstrDniCli As String
Public pstrNomcli As String
Public pstrCodCliente_Ink As String
Public pstrPuntos_Ink As String 'ECASTILLO 30.04.2020

Dim objProducto As New clsProducto
Dim oraDato As oraDynaset
Dim strTipoPrecio As String

Dim objFormaPago As New clsFormaPago

Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

Public pstrdxPrcCero As String
Dim dblDescuento As Double

Private Sub cmdPrecio_Click()
On Error GoTo handle
    If grdPedido.ApproxCount = 0 Then Exit Sub
    If grdPedido.Columns(7).Value = "0" Then
        frm_VTA_CorrecionPrecios.Datos grdPedido.Columns(0).Value, _
                            grdPedido.Columns(1).Value, _
                            grdPedido.Columns(6).Value, _
                            grdPedido.Columns(3).Value, _
                            grdPedido.Columns(4).Value, _
                            grdPedido.Columns(5).Value, _
                            grdPedido.Columns(7).Value
        frm_VTA_CorrecionPrecios.Show vbModal
    Else
        MsgBox "No se puede modificar el precio de este producto" + Chr(13) + _
               " Probablemente usted desea modificar un regalo  ", vbInformation + vbOKOnly, App.ProductName
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdProbar_Click()
'''Dim i As Integer
'''Dim j As Integer
'''
'''
'''    If grdPedido.ApproxCount = 0 Then Exit Sub
'''
'''    j = 0
'''    For i = 0 To objVenta.Producto.UpperBound(1)
'''        If objVenta.Producto(i, 7) = "0" Then
'''                j = j + 1
'''        End If
'''    Next i
'''
'''    If j = 0 Then Exit Sub
'''
'''    If objUsuario.PrecioOnLine = 1 Then

Cal_Promo

End Sub



Private Sub dbcTpTarjeta_Change()
'PagoTarjeta (lblTotal.Caption)
'Cal_Montos

'frm_VTA_FormaPagoTarjeta.pstrDato = "002"
'frm_VTA_FormaPagoTarjeta.pstrDatoDes = "TARJETA"
'frm_VTA_FormaPagoTarjeta.Show
'frm_VTA_FormaPagoTarjeta.SetFocus
Cal_Promo
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    psub_KeyDownAplicacion KeyCode, Shift
    
    If KeyCode = 38 Then
'            If Me.optNinguna.Value = True Then
'                Me.optNinguna.Value = False
'                Me.optCredito.Value = True
'            Else
            If Me.optCredito.Value = True Then
                Me.optCredito.Value = False
                Me.optEfectivo.Value = True
            ElseIf Me.optEfectivo.Value = True Then
                Me.optEfectivo.Value = False
                Me.optCredito.Value = True
            End If
     ElseIf KeyCode = 40 Then
'            If Me.optNinguna.Value = True Then
'                Me.optNinguna.Value = False
'                Me.optEfectivo.Value = True
'            Else
            If Me.optEfectivo.Value = True Then
                Me.optEfectivo.Value = False
                Me.optCredito.Value = True
            ElseIf Me.optCredito.Value = True Then
                Me.optCredito.Value = False
                Me.optEfectivo.Value = True
            End If
    End If
    
    Select Case KeyCode
        Case vbKeyF11
          If objUsuario.EsDelivery = True And grdPedido.ApproxCount > 0 Then
            On Error GoTo CtrlErr1
            frm_DLV_Stock_Total.Datos grdPedido.Columns(0).Value, grdPedido.Columns(1).Value
            frm_DLV_Stock_Total.Show vbModal
            Exit Sub
CtrlErr1:
            Err.Raise Err.Number, "", App.FileDescription
            Else
            frm_VTA_Puntos.Show vbModal
          End If
    End Select
    
     Call FocusOptiones(KeyCode, Shift)
    
End Sub

Public Sub FocusOptiones(KeyCode As Integer, Shift As Integer)
    If optEfectivo.Visible And optCredito.Visible Then
            If Shift = 4 Then
                 Select Case KeyCode
                    Case 69
                            optEfectivo.Value = True
                    Case 84
                            optCredito.Value = True
                            dbcTpTarjeta.SetFocus
                End Select
            End If
    End If
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    psub_KeyDownAplicacion KeyCode, Shift
'End Sub
'AQUI!
Public Sub RefrescarGrilla()
    frmPedido.grdPedido.Array1 = objVenta.Producto
    frmPedido.grdPedido.Refresh
End Sub

Private Sub Form_Load()

On Error GoTo handle
    flgF6 = 0
    Me.lblMensaje.Visible = InStr(1, gvarTNSNAME, "DESA")
    'Initialize the positions/sizes of this form
    lDockedWidth = mdiPrincipal.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = mdiPrincipal.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    lFloatingLeft = Me.left
    lFloatingTop = Me.top
    lFloatingWidth = Me.Width
    lFloatingHeight = Me.Height
    'Start with the form docked in Picture1 on the MDI Form
    'put Form1 in the 'Dock' and position it so its resizing border is
    'hidden outside the confines of Picture1
    bDocked = True
    SetParent Me.hWnd, mdiPrincipal!Picture1.hWnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    mdiPrincipal!Picture1.Visible = True
    HabilitaPermisos
    
    '--- SETEA EL ARREGLO ---'
    psub_BeginArry
    '------------------------'
    ''grdPedido.Columns(0).Alignment = dbgCenter
    ''grdPedido.Columns(1).Alignment = dbgLeft
    ''grdPedido.Columns(2).Alignment = dbgCenter
    ''grdPedido.Columns(3).Alignment = dbgRight
    ''grdPedido.Columns(4).Alignment = dbgRight
    ''grdPedido.Columns(5).Alignment = dbgCenter
    ''grdPedido.Columns(6).Visible = True
    ''grdPedido.Columns(7).Visible = True
    ''grdPedido.MarqueeStyle = dbgHighlightRow
    ''grdPedido.AllowUpdate = False
    
    
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Código", "Descripción", "F", "Cantidad", "Precio", "T", "FlagFraccion", "Regalo", "TipDcto", "ImpDcto", "PrcAnt", "CodAutoriza", "CodUsuario", "Pct_Comi", "CtdProductoOrig", "FlgFraccionOrig", "Dato1", "Dato2", "Dato3", "Dato4", "Dato5", "FlgReceta", "NroLote", "FchVmto", "FlgFarmaco", "Pre Publico")
    ''arrAncho = Array(600, 2500, 200, 400, 800, 190, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 0, 0, 0, 0, 0, 100, 100, 100, 200, 1500)
    arrAncho = Array(0, 3100, 200, 400, 800, 190, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 0, 0, 0, 0, 0, 100, 100, 100, 200, 1500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgCenter, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgCenter, dbgRight)
    
    grdPedido.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim k As Integer
    For k = 0 To 24
        grdPedido.Columns(k).FetchStyle = True
    Next k
    ''grdPedido.Columns(1).FetchStyle = True
    ''grdPedido.Columns(7).FetchStyle = True
    grdPedido.Columns(1).WrapText = True
    PintaIndicadores
    grdPedido.CellTips = dbgFloating
    
    'visible o no el ingreso de cliente (nuevo)
    Dim rs As oraDynaset
    Dim v_Bool As Boolean
    v_Bool = objVenta.MuestraFidelizado("ACTDIDELIZ")
    lbl_Cliente.Visible = v_Bool
    lblCtrl.Visible = v_Bool
    
    'carga de datos de tipo de tarjeta de credito
    Set dbcTpTarjeta.RowSource = objFormaPago.ListaHijo2("002")
    dbcTpTarjeta.ListField = "DESCRIPCION"
    dbcTpTarjeta.BoundColumn = "CODIGO"
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Public Sub psub_BeginArry()
On Error GoTo handle
    ''pxdbDatos.ReDim 0, -1, 0, 5
    Set objVenta = Nothing
    Set objVenta = New clsVenta
    'Set objVenta = New clsVenta
    ''grdPedido.Array = objVenta.Producto
    ''grdPedido.Rebind
    'On Error GoTo ERROR
    grdPedido.Array1 = objVenta.Producto
    grdPedido.Rebind
    
    
''ERROR:
''  On Error GoTo 0
    Exit Sub
handle:
    Set objVenta = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'reset this form's owner to prevent a crash
    Call SetWindowWord(Me.hWnd, SWW_HPARENT, 0&)
End Sub

Private Sub Form_Resize()
'''
'''    If Me.WindowState <> vbMinimized Then
'''        'Update the stored Values
'''        StoreFormDimensions
'''        'position and size the listbox
'''        grdPedido.Move 3 * Screen.TwipsPerPixelX, FormDragger1.Height + (3 * Screen.TwipsPerPixelY) + 300, Me.ScaleWidth - (7 * Screen.TwipsPerPixelX) ', Me.ScaleHeight - (FormDragger1.Height + (6 * Screen.TwipsPerPixelY))
'''    End If
    
End Sub

Private Sub FormDragger1_DblClick()

    'Snap the form in or out of the dock (Picture1)
    bMoving = True 'stop the new dimensions being stored
    
    If bDocked Then
        ''Undock
        Me.Visible = False
        bDocked = False
        SetParent Me.hWnd, 0
        Me.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
        mdiPrincipal!Picture1.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hWnd, SWW_HPARENT, mdiPrincipal.hWnd)
    Else
        ''Dock
        bDocked = True
        SetParent Me.hWnd, mdiPrincipal!Picture1.hWnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        mdiPrincipal!Picture1.Visible = True
    End If
    bMoving = False

End Sub

Private Sub FormDragger1_FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    
    Dim rct As RECT

    'If over Picture1 on mdiPrincipal which we are using as a Dock, set parent
    'of this form to Picture1, and position it at -4,-4 pixels, otherwise
    'set this Form's parent to the desktop and postion it at Left,Top
    'We dont need to size the form, as the DragForm control will have done
    'this for us.
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    'Get the screen based coordinates of Picture1
    GetWindowRect mdiPrincipal!Picture1.hWnd, rct
    'Inflate the rect because we want the form to be bigger than Picture1
    'to hide it's border
    With rct
        .left = .left - 4
        .top = .top - 4
        .right = .right + 4
        .bottom = .bottom + 4
    End With
    'See if the top/left corner of this form is in Picture1's screen rectangle
    'As we have set RepositionForm to false, we are responsible for positioning the form
    If PtInRect(rct, FormLeft, FormTop) Then
        bDocked = True
        SetParent Me.hWnd, mdiPrincipal!Picture1.hWnd
        Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
        mdiPrincipal!Picture1.Visible = True
    Else
        Me.Visible = False
        bDocked = False
        SetParent Me.hWnd, 0
        Me.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
        mdiPrincipal!Picture1.Visible = False
        Me.Visible = True
        'make this form 'float' above the MDI form
        Call SetWindowWord(Me.hWnd, SWW_HPARENT, mdiPrincipal.hWnd)
    End If
    
    'reset the moving flag and store the form dimensions
    bMoving = False
    StoreFormDimensions

End Sub

Private Sub FormDragger1_FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
    
    Dim rct As RECT
    
    'Set the moving flag so we dont store the wrong dimensions
    bMoving = True

    'If over Picture1 on mdiPrincipal which we are using as a Dock, change the width to that of
    'Picture1, else change it to the 'floating width and height
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1

    'Get the screen based coordinates of Picture1
    GetWindowRect mdiPrincipal!Picture1.hWnd, rct
    'Inflate the rect because we want the form to be bigger than Picture1
    'to hide it's border
    With rct
        .left = .left - 4
        .top = .top - 4
        .right = .right + 4
        .bottom = .bottom + 4
    End With
    'See if the top/left corner of this form is in Picture1's screen rectangle

    If PtInRect(rct, FormLeft, FormTop) Then
        FormWidth = lDockedWidth / Screen.TwipsPerPixelX
        FormHeight = lDockedHeight / Screen.TwipsPerPixelY
    Else
        FormWidth = lFloatingWidth / Screen.TwipsPerPixelX
        FormHeight = lFloatingHeight / Screen.TwipsPerPixelY
    End If

End Sub

Private Sub StoreFormDimensions()

   'Store the height/width values
    If Not bMoving Then
        If bDocked Then
            lDockedWidth = Me.Width
            lDockedHeight = Me.Height
        Else
            lFloatingLeft = Me.left
            lFloatingTop = Me.top
            lFloatingWidth = Me.Width
            lFloatingHeight = Me.Height
        End If
    End If
End Sub

'Sub psub_Items(ByRef rxdb As XArrayDB, ByVal intCol%)
'MsgBox "Se supone que esto va en la clase o no????"
'    Dim i%
'    For i = "0" & grdPedido.Columns(0).Value To rxdb.UpperBound(1)
'        rxdb(i, intCol) = i + 1
'    Next i
'End Sub
'
'Public Sub AgregaLinea(ByVal Codigo As String, ByVal Producto As String, ByVal Cantidad As String, ByVal Precio As Double, Optional ByVal flgFracc As Integer = 0)
'MsgBox "Se supone que esto va en la clase o no????"
'    Dim blnAdd As Boolean
'    Dim IntRow As Integer
'    Dim i As Integer
'
'    On Error GoTo CntError
'    IntRow = pxdbDatos.Find(0, 1, Codigo)
'    On Error GoTo 0
'    GoTo Ok
'CntError:
'    IntRow = -1
'    On Error GoTo 0
'    GoTo Ok
'Ok:
'    If IntRow = -1 Then
'            pxdbDatos.AppendRows 1
'            Call psub_Items(pxdbDatos, 0)
'            pxdbDatos(pxdbDatos.UpperBound(1), 1) = Codigo
'            pxdbDatos(pxdbDatos.UpperBound(1), 2) = Producto
'            pxdbDatos(pxdbDatos.UpperBound(1), 3) = IIf(flgFracc = 0, "U", "F")
'            pxdbDatos(pxdbDatos.UpperBound(1), 4) = Cantidad
'            pxdbDatos(pxdbDatos.UpperBound(1), 5) = Format(Precio, "###,###.00")
'
'            dblPrecio = 0: pstrProd = "": pstrPrecio = "": pstrCant = "": pstrSubTot = ""
'            pstrCantFrac = "": pstrCtdFracc = "": pstrPctUnit = "": pstrPrcUniKairo = ""
'            pstrImpuesto = "": pstrComision = "": pstrPromocion = ""
'
'            For i = 0 To pxdbDatos.UpperBound(1)
'                pstrProd = pstrProd & IIf(IsNull(Trim(pxdbDatos(i, 1))) Or Trim(pxdbDatos(i, 1)) = "", "0", Trim(pxdbDatos(i, 1))) & "|"
'                pstrCant = pstrCant & IIf(IsNull(Trim(pxdbDatos(i, 4))) Or Trim(pxdbDatos(i, 4)) = "", "0", Trim(pxdbDatos(i, 4))) & "|"
'                pstrPrecio = pstrPrecio & IIf(IsNull(Trim(pxdbDatos(i, 5))) Or Trim(pxdbDatos(i, 5)) = "", "0", Trim(pxdbDatos(i, 5))) & "|"
'                pstrSubTot = pstrSubTot & pxdbDatos(i, 4) * pxdbDatos(i, 5) & "|"
'
'                pstrCantFrac = pstrCantFrac & "0|"
'                pstrCtdFracc = pstrCtdFracc & "1|"
'                pstrPctUnit = pstrPctUnit & "0|"
'                pstrPrcUniKairo = pstrPrcUniKairo & "0|"
'                pstrImpuesto = pstrImpuesto & "0|"
'                pstrComision = pstrComision & "0|"
'                pstrPromocion = pstrPromocion & "|"
'
'                dblPrecio = dblPrecio + IIf(IsNull(Trim(pxdbDatos(i, 5))) Or Trim(pxdbDatos(i, 5)) = "", "0", Trim(pxdbDatos(i, 5)))
'            Next i
'
'      Else
'
'            pxdbDatos(grdPedido.Bookmark, 4) = Cantidad
'            pxdbDatos(grdPedido.Bookmark, 5) = Format(Precio, "###,###.00")
'
'            psub_Calcula_Array_Montos
'
'    End If
'
'        psub_Cal_Montos_General
'
'        sub_ClearVarPublic
'        grdPedido.Rebind
'End Sub

Private Sub sub_ClearVarPublic()
On Error GoTo handle
    frm_VTA_RecetarioM.pstrFlgRM = ""
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Function Calcula_Redondeo(ByVal vstrRedondeo)

    On Error GoTo Err1
    Dim xRed1$, xRed2$
    If right(vstrRedondeo, 3) >= ".05" Then
         If Mid(Mid(vstrRedondeo, InStr(vstrRedondeo, "."), 2), 2, 1) = "9" Then
                If right(vstrRedondeo, 1) >= 5 Then
                       vstrRedondeo = (100 - right(vstrRedondeo, 2)) / 10
                    Else
                       vstrRedondeo = right(vstrRedondeo, 2)
                End If
                Calcula_Redondeo = vstrRedondeo / 10
            Else
                If right(vstrRedondeo, 1) >= "5" Then
                    xRed1 = Mid(vstrRedondeo, InStr(vstrRedondeo, "."), 2) + 0.1
                    xRed2 = right(vstrRedondeo, 2) / 10
                    vstrRedondeo = Round((xRed1 * 10) - (xRed2), 2)
                  Else
                    vstrRedondeo = right(vstrRedondeo, 1)
                End If
                Calcula_Redondeo = vstrRedondeo / 10
         End If
      Else
         vstrRedondeo = "0"
         Calcula_Redondeo = vstrRedondeo / 10
    End If
    
Err1:
    On Error GoTo 0
End Function

Private Function calcula_TPaga(ByVal vstrTotal#, ByVal vstrRedondeo#)

    On Error GoTo Err2
    If right(vstrTotal, 2) = "00" Then
        calcula_TPaga = Format(val(vstrTotal), "#0.00")
      Else
        If Mid(vstrRedondeo, 3, 1) >= 0 Then
            ''calcula_TPaga = Format(Val(vstrTotal), "#0") & "." & Mid(vstrRedondeo, 3, 1)
            ''calcula_TPaga = Mid(vstrTotal, 1, InStr(vstrTotal, ".") - 1) & "." & Format(Mid(vstrRedondeo, 3, 1), "00")
            calcula_TPaga = Format(vstrTotal + vstrRedondeo, "0.00")
          Else
            calcula_TPaga = Format(val(vstrTotal), "#0.00")
        End If
    End If
        
Err2:
    On Error GoTo 0
End Function

Public Sub PagoCon(ByVal vstrPagoCon#)

    On Error GoTo Err3
        lblPagado.Caption = Format(vstrPagoCon, "#0.00")
Err3:
    On Error GoTo 0
End Sub

Public Function Calcula_Vuelto(ByVal vstrTPagar#, ByVal vstrPagoCon#)

    On Error GoTo Err4
        Calcula_Vuelto = Format(vstrPagoCon - vstrTPagar, "#0.00")
Err4:
    On Error GoTo 0
End Function

Public Sub LimpiarTodo()
On Error GoTo handle
    pxdbDatos.Clear
    grdPedido.Array1 = pxdbDatos
    grdPedido.Rebind
    lblTotal.Caption = "0.00"
    lblPuntosRed.Caption = "0.00"
    lblPuntosAcum = "0"
    lblTotalDescuento.Caption = "0.00"
    lblRedondeo.Caption = "0.00"
    lblTotalPagar.Caption = "0.00"
    lblVuelto.Caption = "0.00"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


Private Sub grdPedido_AfterInsert()
    'Cal_Promo
End Sub

Private Sub grdPedido_DblClick()
On Error GoTo handle
Dim strIdFrac As String
Dim strIndicadorReceta As String
Dim strEsEspecieValorada As String


    If grdPedido.ApproxCount = 0 Then Exit Sub
        ''        strIdFrac = objProducto.ListaDevFracciona(grdPedido.Columns(0).Value)
        ''        If strIdFrac = "0" Then
        ''             frm_VTA_CantidadProducto.chkFraccionamiento.Enabled = False
        ''          Else
        ''             frm_VTA_CantidadProducto.chkFraccionamiento.Enabled = True
        ''        End If

    strEsEspecieValorada = objProducto.EsEspecieValorada(grdPedido.Columns(0).Value)
    If strEsEspecieValorada = "1" Then Exit Sub

    If objProducto.FnEsModalidad_Recetario(grdPedido.Columns(0).Value) > 0 Then
        '''frm_VTA_RecetarioM.pstrFlgRM = "1"
        frm_VTA_RecetarioM.Show
        frm_VTA_RecetarioM.SetFocus
    Else
        strIdFrac = objProducto.ListaDevFracciona(grdPedido.Columns(0).Value, objUsuario.CodigoLocal, objVenta.CodModalidadVenta)
        strIndicadorReceta = objProducto.IndicadorReceta(grdPedido.Columns(21).Value)
        If objVenta.bk_ServiceType = "RET" And grdPedido.Columns(0).Value = "09938" Then
        Else
            frm_VTA_CantidadProducto.subDatos grdPedido.Columns(0).Value, grdPedido.Columns(1).Value, strTipoPrecio, lblModalidad.Caption, grdPedido.Columns(7).Value, strIdFrac, strIndicadorReceta, grdPedido.Columns(21).Value, "", grdPedido.Columns(22).Value, grdPedido.Columns(23).Value
        End If
        '''''''''''''''Cal_Promo
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdPedido_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)


On Error GoTo handle
Dim n As Double
Dim c As Double

If Condition = 0 Then
    Select Case Col
        Case 1
          n = val(grdPedido.Columns(7).CellText(Bookmark))
          If n = 1 Then
             CellStyle.ForeColor = vbRed
             CellStyle.Font.Bold = True
          End If
          
          If n = 2 Then
             CellStyle.ForeColor = &H80FF&
             CellStyle.Font.Bold = True
          End If
            
          
            
    End Select
End If

If Condition = 2 Or Condition = 3 Then
    Select Case Col
        Case 1
          n = val(grdPedido.Columns(7).CellText(Bookmark))
          If n > 0 Then
             CellStyle.ForeColor = vbBlue
             CellStyle.Font.Bold = True
          End If
    End Select
End If

    '** Precio cero **'
    Select Case Col
        Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
          c = val(grdPedido.Columns(4).CellText(Bookmark))
          If c <= 0 Then
                CellStyle.ForeColor = RGB(255, 140, 0)
                ''CellStyle.ForeColor = vbRed
                CellStyle.Font.Bold = True
            Else
                CellStyle.ForeColor = RGB(0, 0, 0)
                CellStyle.Font.Bold = False
          End If
    End Select
    
    '** Certificado SOAT que se haga la devolución con Guia **'
    If frm_VTA_CantidadProducto.flgEspecieValorada = "1" And objVenta.ptmModalidad = Guias_Remision Then
    Select Case Col
        Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
            CellStyle.ForeColor = RGB(0, 0, 0)
            CellStyle.Font.Bold = False
    End Select
    End If
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName


End Sub

Private Sub grdPedido_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueDBGrid70.StyleDisp)
   On Error GoTo Control
   Dim strMensaje As String
    With grdPedido
        
        strMensaje = ""
        
        Select Case ColIndex
            Case 22 And "" & .Columns(22).Value <> ""
                strMensaje = "Lote :" & UCase(.Columns(22).Value)
            Case 23 And "" & .Columns(23).Value <> ""
                strMensaje = strMensaje & "  F.V :" & Format(.Columns(23).Value, "dd/yyyy")
        End Select
        
        CellTip = Trim(strMensaje)
    End With
   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub grdPedido_FirstRowChange(ByVal SplitIndex As Integer)
    ''Debug.Print "FirstRowChange"
    ''Cal_Promo
End Sub


Private Sub grdPedido_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo handle
        Select Case KeyCode
            
            Case vbKeyReturn
                Call grdPedido_DblClick
                
            Case vbKeyF1
                frm_VTA_Busqueda.txtBuscar.SetFocus
                
            Case vbKeyF2
                frm_VTA_Busqueda.grdProductos.SetFocus
            
            Case vbKeyF7
                Exit Sub
CtrlErr:
                MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
             Case vbKeyDelete
                If grdPedido.ApproxCount = 0 Then
                    frm_VTA_Busqueda.Show
                    Exit Sub
                End If
                
                
                
                Dim PromstrCodProducto As String
                Dim PromstrCodTipoVenta As TipoVenta
                Dim indicadorRegalo As String '0=no 1=si
                Dim bolActualizaPromo As Boolean
                Dim indice As Integer
                Dim i%
            
                
                PromstrCodProducto = grdPedido.Columns(0).Value
                PromstrCodTipoVenta = grdPedido.Columns(5).Value
                indicadorRegalo = grdPedido.Columns(7).Value
                
                '----INICIO JOSE MELGAR
                        For i = 0 To objVenta.Producto.Count(1) - 1
                            If objVenta.Producto(i, 0) = PromstrCodProducto Then
                                indice = i
                                GoTo salta
                            End If
                        Next
salta:
                        If indicadorRegalo <> "0" Then
                            bolActualizaPromo = False
                        Else
                            If objVenta.Producto(indice, 26) = "" Then
                                bolActualizaPromo = False
                            Else
                                bolActualizaPromo = True
                            End If
                        End If
                '----FIN JOSE MELGAR
                    Dim cantidadPedidoAnterior As Integer
                 
                ''esto es para cuando eliminan un regalo de la promocion
                objVenta.EliminaProducto PromstrCodProducto, PromstrCodTipoVenta
                grdPedido.Delete
                
           
                '****** Cuando es Recetario Magistral ******'
                objVenta.LimpiaRecetario
                Unload frm_VTA_RecetarioM
                            
                ''lblTotal.Caption = objVenta.Totales(0)
                ''Cal_Promo
                 
                If bolActualizaPromo Then
                    Cal_Promo
                End If
                         
                Cal_Montos
                
                If grdPedido.ApproxCount = 0 Then
                    frm_VTA_Busqueda.Show
                End If
                

                
''                grdPedido.Array = objVenta.EliminaProducto(strCodProducto, objVenta.CodigoTipoVenta)
''                If grdPedido.EditActive = False Then
''                    Dim IntRow%
''                    Dim strB$
''                    strB = pxdbDatos(grdPedido.BookMark, 1)
''                    grdPedido.Delete
''
''                    For IntRow = 0 To pxdbDatos.UpperBound(1)
''                        If pxdbDatos.UpperBound(1) > -1 Then
''                            strB = pxdbDatos.Find(0, 1, strB, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
''                            If strB > -1 Then
''                                pxdbDatos.DeleteRows strB
''                            End If
''                        End If
''                    Next IntRow
''
''                    Call psub_Items(pxdbDatos, 0)
''                    grdPedido.ReOpen
''
''                    If pxdbDatos.UpperBound(1) = -1 Then frm_VTA_Busqueda.txtBuscar.SetFocus: Exit Sub
''
''                    psub_Calcula_Array_Montos
''                    psub_Cal_Montos_General
''                    frmPedido.grdPedido.SetFocus
''
''                End If
        Case vbKeyF12
            If objUsuario.TipoMaquina = "003" And grdPedido.ApproxCount > 0 Then
            frm_DLV_Stock_Total.Datos grdPedido.Columns(0).Value, grdPedido.Columns(1).Value
            frm_DLV_Stock_Total.strCodigoProducto = grdPedido.Columns(0).Value
            frm_DLV_Stock_Total.strDescripcionProducto = grdPedido.Columns(1).Value
            frm_DLV_Stock_Total.strTranferencias = True
            frm_DLV_Stock_Total.Show vbModal
            End If
        End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
        
End Sub

Private Sub psub_Calcula_Array_Montos()
On Error GoTo handle
        dblPrecio = 0: pstrPrecio = "": pstrCant = "": pstrSubTot = "": dblDescuento = 0
        Dim i%
        For i = 0 To pxdbDatos.UpperBound(1)
            pstrCant = pstrCant & IIf(IsNull(Trim(pxdbDatos(i, 4))) Or Trim(pxdbDatos(i, 4)) = "", "0", Trim(pxdbDatos(i, 4))) & "|"
            pstrPrecio = pstrPrecio & IIf(IsNull(Trim(pxdbDatos(i, 5))) Or Trim(pxdbDatos(i, 5)) = "", "0", Trim(pxdbDatos(i, 5))) & "|"
            pstrSubTot = pstrSubTot & pxdbDatos(i, 4) * pxdbDatos(i, 5) & "|"
            dblPrecio = dblPrecio + IIf(IsNull(Trim(pxdbDatos(i, 5))) Or Trim(pxdbDatos(i, 5)) = "", "0", Trim(pxdbDatos(i, 5)))
            dblDescuento = dblDescuento + IIf(IsNull(Trim(pxdbDatos(i, 10))) Or Trim(pxdbDatos(i, 10)) = "", "0", Trim(pxdbDatos(i, 10)))
        Next i

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
        
End Sub

Private Sub psub_Cal_Montos_General()
On Error GoTo handle
        '--- Calcula Montos --------'
        lblTotal.Caption = Format(dblPrecio, "#0.00")
        
        lblRedondeo.Caption = Calcula_Redondeo(Format(dblPrecio, "#0.00"))
        lblTotalPagar.Caption = calcula_TPaga(Format(dblPrecio, "#0.00"), lblRedondeo.Caption)
        
        '---------------------------'
        '-- SI ES KEITO HACE ESTO --'
        If frm_VTA_FacServPrestados.pstrIDKeito = "KE" Then
            lblPagado.Caption = Format(dblPrecio, "#0.00")
            lblVuelto.Caption = Format(dblPrecio, "#0.00") - Format(dblPrecio, "#0.00")
        End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
        
End Sub


Public Sub ReCalculaPrecio(ByVal pstrTipoPrecio As String, Optional pstrCodConvenio As String = "")
On Error GoTo handle
Dim i%
Dim indicador As String

        strTipoPrecio = pstrTipoPrecio
        
        For i = 0 To objVenta.Producto.UpperBound(1)
        
                If objVenta.CodigoTipoVenta = Cobro_Responsabilidad And objVenta.Producto(i, 4) < 0 Then
                Else
                    indicador = objProducto.CodIndicadorReceta(objVenta.Producto(i, 0))
                    
                    If objUsuario.EsDelivery = True Then
                        If objUsuario.CodLocalCallCenter = "1DLV" Then 'ECASTILLO 22.06.2020
                            Set oraDato = objProducto.ListaDato("94", mdiPrincipal.ctlCliente1.LocalAsignado, strTipoPrecio, objVenta.Producto(i, 0), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), pstrCodConvenio, objUsuario.CodLocalCallCenter)
                        Else
                            Set oraDato = objProducto.ListaDato(objUsuario.CodigoEmpresa, mdiPrincipal.ctlCliente1.LocalAsignado, strTipoPrecio, objVenta.Producto(i, 0), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), pstrCodConvenio, objUsuario.CodLocalCallCenter)
                        End If
                    Else
                        If objUsuario.CodLocalCallCenter = "1DLV" Then 'ECASTILLO 22.06.2020
                            Set oraDato = objProducto.ListaDato("94", objUsuario.CodigoLocal, strTipoPrecio, objVenta.Producto(i, 0), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), pstrCodConvenio, objUsuario.CodLocalCallCenter)
                        Else
                            Set oraDato = objProducto.ListaDato(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strTipoPrecio, objVenta.Producto(i, 0), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), pstrCodConvenio, objUsuario.CodLocalCallCenter)
                        End If
                    End If
                    
                    
                    If frm_VTA_RecetarioM.pstrFlgRM <> "1" Then
                      ''If objVenta.ptmModalidad = Cobro_Responsabilidad Then Exit Sub
                        ''pstrdxPrcCero = ""
                        If oraDato(4).Value = 0 And objVenta.Producto(i, 6) = "0" Then
                           If pstrdxPrcCero = "" Then
                               pstrdxPrcCero = oraDato(1).Value
                             Else
                               pstrdxPrcCero = pstrdxPrcCero & "|" & oraDato(1).Value
                           End If
                        End If
                        
                        Dim flg_ruteoA_cnv
                        flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
                        If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
                            GoTo cnvNoRuteaAuto
                        End If
                        Dim flg_2e_reserva
                        flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
                        If flg_2e_reserva = "0" Then
cnvNoRuteaAuto:
                            objVenta.AgregaProducto objVenta.Producto(i, 0), objVenta.Producto(i, 1), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), oraDato(4).Value, objVenta.CodigoTipoVenta, objVenta.Producto(i, 7), , , , , , , , , , , oraDato(5).Value ' 06.01.2021
                        Else
                        'I.CVIERA 04.01.2021 | 05.01.2021.REVISAR
'                        Dim Buscar As String
'                        Dim Buscar1 As String
'                        Dim a As Integer
'
'                        a = 0
'                        For a = 0 To grdPedido.ApproxCount - 1
'                            Buscar = CStr(frmPedido.grdPedido.Columns(0).CellValue(a))
'                            Buscar1 = CStr(frmPedido.grdPedido.Columns(4).CellValue(a))
'                            If Buscar = "09938" Then
'                                'If CStr(frmPedido.grdPedido.Columns(0).CellValue(a)) = "09938" Then
'                                'frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, chkFraccionamiento, frm_VTA_Busqueda.grdProductos.Columns("PRECIO").Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , frm_VTA_Busqueda.grdProductos.Columns("PRECIO").Value, , , FracEnabled, , , frm_VTA_Busqueda.grdProductos.Columns("FLG_SEG").Value)
'                                objVenta.AgregaProducto objVenta.Producto(i, 0), objVenta.Producto(i, 1), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), frmPedido.grdPedido.Columns(4).CellValue(a), objVenta.CodigoTipoVenta, objVenta.Producto(i, 7), , , , , , , , , , , frmPedido.grdPedido.Columns(4).CellValue(a), , , , , , 1
'                            Else
'                                    'frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, chkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled, , , frm_VTA_Busqueda.grdProductos.Columns("FLG_SEG").Value)
'                                objVenta.AgregaProducto objVenta.Producto(i, 0), objVenta.Producto(i, 1), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), oraDato(4).Value, objVenta.CodigoTipoVenta, objVenta.Producto(i, 7), , , , , , , , , , , oraDato(5).Value
'                                'End If
'                            End If
'                            'frm_VTA_Busqueda.grdProductos.MoveNext
'                        Next a
                        'F.CVIERA 04.01.2021
                            'I.ECASTILLO 05.01.2021 | 06.01.2021
                            Dim ii
                            Dim nuevoPrecio As Double
                            For ii = 0 To objVenta.Producto.UpperBound(1)
                                If objVenta.Producto(ii, 0) = "09938" And objVenta.Producto(ii, 30) = 1 Then
                                    'debido a que el producto DLV no es fraccionable, el calculo para este caso queda pendiente
                                    nuevoPrecio = IIf(objVenta.Producto(ii, 6) = "0", objVenta.Producto(ii, 5) * CDbl(val(frm_VTA_MetodosSegmentos.strPrecioTipo)), (CDbl(val(frm_VTA_MetodosSegmentos.strPrecioTipo)) / 1) * objVenta.Producto(ii, 3))
                                    objVenta.AgregaProducto objVenta.Producto(ii, 0), objVenta.Producto(ii, 1), _
                                                            objVenta.Producto(ii, 3), objVenta.Producto(ii, 6), _
                                                            nuevoPrecio, objVenta.Producto(ii, 5), _
                                                            objVenta.Producto(ii, 7), , , , , , _
                                                            objVenta.Producto(ii, 21), objVenta.Producto(ii, 13), _
                                                            objVenta.Producto(ii, 22), objVenta.Producto(ii, 23), , _
                                                            nuevoPrecio, , , objVenta.Producto(ii, 2), , , _
                                                            1
                                Else
                                    objVenta.AgregaProducto objVenta.Producto(i, 0), objVenta.Producto(i, 1), objVenta.Producto(i, 3), IIf(objVenta.Producto(i, 2) = "U", 0, 1), oraDato(4).Value, objVenta.CodigoTipoVenta, objVenta.Producto(i, 7), , , , , , , , , , , oraDato(5).Value
                                End If
                            Next ii
                            'F.ECASTILLO 05.01.2021
                        End If
                    End If
                End If
        Next i
        'I.ECASTILLO 07.10.2020
        If gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "RCALCPRCIO") = 1 Then
            grdPedido.Rebind 'ECASTILLO 07.10.2020 - AQUI SE PIERDEN LOS PRECIOS PROMOCION
            Cal_Montos
        End If
        'F.ECASTILLO 07.10.2020
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
        
End Sub

Public Sub Cal_Montos()
On Error GoTo handle
 
If objVenta.ParametroValor("ACT_PCTCOM") = "1" Then
   If Len(Trim(objVenta.CodigoCliente)) > 0 Then
      ' busca si el DNI existe en RENIEC para indicar si va
      ' dar comision al Vendedor
      objVenta.vExisteDNI_RENIEC = objVenta.getExisteDNI_RENIEC(objVenta.CodigoCliente)
      If objVenta.vExisteDNI_RENIEC = "N" Then
         frmPedido.lblDniInvalido.Visible = True
      Else
          frmPedido.lblDniInvalido.Visible = False
      End If
    End If
End If
        


'            Dim dlbg As Double
'            Dim g As Integer
'            For g = 0 To objVenta.Servicios.UpperBound(1)
'                dlbg = dlbg + Val(objVenta.Servicios(g, 8))
'            Next g
Dim i As Integer
        Dim puntosAcum As String
        'objVenta.PctBeneficiario = 70
        If frm_VTA_FormaPago.Modificacion = True Then Exit Sub
        lblTotal.Caption = Format(objVenta.Totales(0), "###,##0.00")  '+ dlbg
        'lbltotaldescuento.Caption=
        'For i = 0 To pxdbDatos.UpperBound(1)
             'dblDescuento = dblDescuento + IIf(IsNull(Trim(pxdbDatos(i, 10))) Or Trim(pxdbDatos(i, 10)) = "", "0", Trim(pxdbDatos(i, 10)))
        'Next i
        lblTotalDescuento.Caption = Format(objVenta.Totales(12), "###,##0.00")
        lblPuntosRed.Caption = Format(objVenta.MontoRedime, "###,##0.00")
        lblRedondeo.Caption = objVenta.Totales(1)
        lblTotalPagar.Caption = objVenta.Totales(2)  '- dlbg
        lblcopago.Caption = objVenta.Totales(8)
        lblPagado.Caption = objVenta.Totales(6)
        lblVuelto.Caption = objVenta.Totales(7)  '- dlbg
        'lblVuelto.Caption = "4"  '- dlbg
        lblPctCopago.Caption = objVenta.PctBeneficiario
        lblTipoCambio.Caption = objVenta.Totales(10)
        lblTotalDolares.Caption = objVenta.Totales(11)
        If objUsuario.CodLocalCallCenter = "0DLV" Then
            lblPuntosAcum.Caption = objVenta.Totales(13)
        End If
        'lblPuntosAcum.Caption = objVenta.Totales(13)
        
        'If optEfectivo.Value = True Then
            'PagoEfectivo (lblTotal.Caption)
        'End If
       
        'If optCredito.Value = True Then
            'PagoTarjeta (lblTotal.Caption)
        'End If
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


Public Sub Cal_Promo()
On Error GoTo handle
Dim i As Integer
Dim ctdProductosNormales As Integer
Dim strCia As String ' GNIBIN 20210127 Proyecto MultiMarca
    
 '   On Error GoTo CtrlErr
    
    'CAMBIAR EN EL SP
    
    If objVenta.CodigoTipoVenta = FactServicios Or objVenta.CodigoTipoVenta = Cobro_Responsabilidad Then Exit Sub
  
    If grdPedido.ApproxCount = 0 Then Exit Sub
                
    ctdProductosNormales = 0
    For i = 0 To objVenta.Producto.UpperBound(1)
        If objVenta.Producto(i, 7) = "0" Then
                ctdProductosNormales = ctdProductosNormales + 1
        End If
    Next i
                    
    If ctdProductosNormales = 0 Then
    
        For i = 0 To objVenta.Producto.UpperBound(1)
            objVenta.Producto.DeleteRows (0)
        Next

        
        objVenta.Producto.Clear
        grdPedido.Limpiar
        Exit Sub
    End If
                
    Dim cod, desc  As String
    Dim arrValores() As String
    Dim nImporte As Double
    If objUsuario.PrecioOnLine = True Then 'Cal_Promo
        
        If dbcTpTarjeta.Visible And dbcTpTarjeta.Text <> "" Then
            arrValores = Split(dbcTpTarjeta.Text, " ")
            cod = arrValores(0)
            desc = arrValores(1)
        Else
            cod = ""
            desc = ""
        End If
        'I.ECASTILLO 28.01.2021 | se reversa los cambios de gnibin
'        'INI GNIBIN 20210127 Proyecto MultiMarca
'        strCia = mdiPrincipal.ctlCliente1.sCia
'        Select Case strCia
'            Case "94", "93", "92", "1DLV"
'                strCia = "1DLV"
'            Case Else
'                strCia = "0DLV"
'        End Select
'        'FIN GNIBIN 20210127 Proyecto MultiMarca
        
        If objUsuario.EsDelivery = True Then
            'objProducto.ProcesaXProducto gclsOracle.ODataBase, objUsuario.CodigoEmpresa, mdiPrincipal.ctlCliente1.LocalAsignado, objVenta.CodigoConvenio
            'Se cambio por el local de despacho
            'jlopez
            '06/05/2008
            
            objProducto.ProcesaXProducto gclsOracle.ODataBase, objUsuario.CodigoEmpresa, mdiPrincipal.ctlCliente1.LocalAsignado, objVenta.CodigoConvenio, mdiPrincipal.ctlCliente1.LocalDespacho, Me.optEfectivo.Value, Me.optCredito.Value, cod, lblTotal.Caption 'GNIBIN 20210127 Proyecto MultiMarca
            'objProducto.ProcesaXProducto gclsOracle.ODataBase, strCia, objUsuario.CodigoEmpresa, mdiPrincipal.ctlCliente1.LocalAsignado, objVenta.CodigoConvenio, mdiPrincipal.ctlCliente1.LocalDespacho, Me.optEfectivo.Value, Me.optCredito.Value, cod, lblTotal.Caption 'GNIBIN 20210127 Proyecto MultiMarca
        Else
            objProducto.ProcesaXProducto gclsOracle.ODataBase, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.CodigoConvenio, , Me.optEfectivo.Value, Me.optCredito.Value, cod, lblTotal.Caption 'GNIBIN 20210127 Proyecto MultiMarca
            'objProducto.ProcesaXProducto gclsOracle.ODataBase, strCia, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.CodigoConvenio, , Me.optEfectivo.Value, Me.optCredito.Value, cod, lblTotal.Caption  'GNIBIN 20210127 Proyecto MultiMarca
        End If
        'F.ECASTILLO 28.01.2021
        
        'I.ECASTILLO 07.10.2020 - AQUI DEBERÍA HABER UN PROCESO QUE MODIFIQUE IMPORTE TARJETA EN CASO OH,AGORA,INTERBANK
        '                       - YA QUE ESTO DESENCADENA PROMOCIÓN QUE MODIFICA PRECIO (IMPORTE A PAGAR)
        '                       - BUSCAR FORMA PAGO REGISTRADO, SI COINCIDE CON CRITERIO ASIGNAR NUEVO IMPORTE
        If gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "REASIMPTJT") = 1 Then
            If objVenta.CodModalidadVenta = Venta_Convenio Then
                nImporte = Format(frmPedido.lblcopago, "0.00")
            Else
                nImporte = Format(frmPedido.lblTotal, "0.00")
            End If
            objVenta.updTarjetaImporte nImporte, "002"
        End If
        'F.ECASTILLO 07.10.2020
        
        Cal_Montos
        grdPedido.Rebind
        DoEvents
        
        Dim strPromocionActual As String
        Dim IO As Integer
        'If IO < objVenta.xProductoRegalo.Count(1) Then
        While IO < objVenta.xProductoRegalo.Count(1)
            If strPromocionActual <> objVenta.xProductoRegalo(IO, 18) Then
                strPromocionActual = objVenta.xProductoRegalo(IO, 18)
                If objVenta.ObtieneCuentaMaxima(strPromocionActual) <> 0 Then 'objVenta.CUANTAMAXIMA <> 0 Then 'ECASTILLO 04.07.2020
                If val(objVenta.CUANTOREGALO(strPromocionActual)) <> val(objVenta.CUANTOTENGOREGALO(strPromocionActual)) Or val(objVenta.CUANTOREGALO(strPromocionActual)) <> val(frm_VTA_RegaloVar.carga(strPromocionActual)) Then
                    'CUANTOREGALO
                    'If objVenta.CUANTAMAXIMA <> Val(objVenta.CUANTOTENGOREGALO(strPromocionActual)) Then
                    If objVenta.ObtieneCuentaMaxima(strPromocionActual) <> val(objVenta.CUANTOTENGOREGALO(strPromocionActual)) Then 'ECASTILLO 04.07.2020
                        frm_VTA_RegaloVar.Show vbModal
                    Else
                        Dim f1 As Integer
                        f1 = 0 'ECASTILLO 04.07.2020
                        While f1 < xProductoRegaloBK.Count(1)
                            'If xProductoRegaloBK(f1, 2) > 0 Then
                            If xProductoRegaloBK(f1, 2) > 0 And xProductoRegaloBK(f1, 18) = strPromocionActual Then 'ECASTILLO 04.07.2020
                                objVenta.AgregaProducto xProductoRegaloBK(f1, 0), _
                                                        xProductoRegaloBK(f1, 1), _
                                                        xProductoRegaloBK(f1, 2), _
                                                        xProductoRegaloBK(f1, 3), _
                                                        xProductoRegaloBK(f1, 4), _
                                                        xProductoRegaloBK(f1, 5), _
                                                        xProductoRegaloBK(f1, 6), _
                                                        xProductoRegaloBK(f1, 7), _
                                                        xProductoRegaloBK(f1, 8), _
                                                        xProductoRegaloBK(f1, 9), _
                                                        xProductoRegaloBK(f1, 10), _
                                                        xProductoRegaloBK(f1, 11), _
                                                        xProductoRegaloBK(f1, 12), _
                                                        xProductoRegaloBK(f1, 13), _
                                                        xProductoRegaloBK(f1, 14), _
                                                        xProductoRegaloBK(f1, 15), _
                                                        xProductoRegaloBK(f1, 16), _
                                                        xProductoRegaloBK(f1, 17), _
                                                        xProductoRegaloBK(f1, 18), _
                                                        xProductoRegaloBK(f1, 19), _
                                                        xProductoRegaloBK(f1, 20), _
                                                        xProductoRegaloBK(f1, 21), _
                                                        xProductoRegaloBK(f1, 22)
                            End If
                            f1 = f1 + 1
                        Wend
                    End If
                Else
                    Dim f As Integer
                    While f < xProductoRegaloBK.Count(1)
                        objVenta.AgregaProducto xProductoRegaloBK(f, 0), _
                                                xProductoRegaloBK(f, 1), _
                                                xProductoRegaloBK(f, 2), _
                                                xProductoRegaloBK(f, 3), _
                                                xProductoRegaloBK(f, 4), _
                                                xProductoRegaloBK(f, 5), _
                                                xProductoRegaloBK(f, 6), _
                                                xProductoRegaloBK(f, 7), _
                                                xProductoRegaloBK(f, 8), _
                                                xProductoRegaloBK(f, 9), _
                                                xProductoRegaloBK(f, 10), _
                                                xProductoRegaloBK(f, 11), _
                                                xProductoRegaloBK(f, 12), _
                                                xProductoRegaloBK(f, 13), _
                                                xProductoRegaloBK(f, 14), _
                                                xProductoRegaloBK(f, 15), _
                                                xProductoRegaloBK(f, 16), _
                                                xProductoRegaloBK(f, 17), _
                                                xProductoRegaloBK(f, 18), _
                                                xProductoRegaloBK(f, 19), _
                                                xProductoRegaloBK(f, 20), _
                                                xProductoRegaloBK(f, 21), _
                                                xProductoRegaloBK(f, 22) _
                                                
                        f = f + 1
                    Wend
                End If
            End If
            End If 'ECASTILLO 04.07.2020
            IO = IO + 1
        Wend
        'End If
        Cal_Montos
        grdPedido.Rebind
        
        'grdPedido.Refresh
        'grdPedido.SetFocus
    End If
    Exit Sub
'CtrlErr:
'    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName


handle:
MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub


Private Sub imgCalculadora_Click()
frm_VTA_Calculadora.Show
End Sub

''arturo escate
Sub HabilitaPermisos()
On Error GoTo handle
    Dim j As Integer
    While j < Me.Controls.Count
        Select Case TypeName(Me.Controls(j))
        Case "CommandButton"
            If Not Me.Controls(j).Tag = "" Then Me.Controls(j).Enabled = False
        End Select
        j = j + 1
    Wend
    j = 0
    Dim objPermisos As New clsAutorizacion
    Dim rsPermisos As oraDynaset
        Set rsPermisos = objPermisos.ListaPermisos(objUsuario.Aplicacion, objUsuario.Codigo)
        While Not rsPermisos.EOF
            If Me.name = rsPermisos("DES_URL").Value Then
                Me.Controls(rsPermisos("DES_ICONO").Value).Enabled = True
                Me.Controls(rsPermisos("DES_ICONO").Value).Tag = rsPermisos("COD_MENU").Value
            End If
            rsPermisos.MoveNext
        Wend
    Set objPermisos = Nothing

    Exit Sub
handle:
    Set objPermisos = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub imgHelpDesk_Click()
abre
End Sub


Sub abre()
frm_VTA_HelpDesk.Show vbModal
End Sub


Sub PintaIndicadores()
ImgNumClienD.Picture = ImageList1.ListImages(4).Picture
ImgNumClienT.Picture = ImageList1.ListImages(4).Picture
ImgValePromD.Picture = ImageList1.ListImages(4).Picture
ImgValePromT.Picture = ImageList1.ListImages(4).Picture
If objUsuario.EsDelivery = False Then
Dim rs As oraDynaset

    If objUsuario.MetaValePromedio = 0 Or objUsuario.MetaNumeroCliente = 0 Then Exit Sub
    ''hay que quitar la fecha para que pueda sacar la fecha del dia
    Set rs = objVenta.ListaIndicadores(objUsuario.CodigoLocal, objUsuario.Codigo, objUsuario.sysdate, objUsuario.MetaValePromedio, objUsuario.MetaNumeroCliente)
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    If val("" & rs("AVANCE_VP").Value) >= 100 Then
        ImgValePromD.Picture = ImageList1.ListImages(1).Picture
    Else
        ImgValePromD.Picture = ImageList1.ListImages(2).Picture
    End If
    lblValePromD.Caption = "S/." & val("" & rs("VALE_PROMEDIO").Value)
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    If val("" & rs("AVANCE_NC").Value) >= 100 Then
        ImgNumClienD.Picture = ImageList1.ListImages(1).Picture
    Else
        ImgNumClienD.Picture = ImageList1.ListImages(2).Picture
    End If
    lblNumClienteD.Caption = val("" & rs("NUM_CLIENTE").Value)
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    If val("" & rs("AVANCE_VPT").Value) >= 100 Then
        ImgValePromT.Picture = ImageList1.ListImages(1).Picture
    Else
        ImgValePromT.Picture = ImageList1.ListImages(2).Picture
    End If
    lblValePromT.Caption = "S/." & val("" & rs("VALE_PROMEDIO_TOTAL").Value)
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    If val("" & rs("AVANCE_NCT").Value) >= 100 Then
        ImgNumClienT.Picture = ImageList1.ListImages(1).Picture
    Else
        ImgNumClienT.Picture = ImageList1.ListImages(2).Picture
    End If
    lblNumClienteT.Caption = val("" & rs("NUMCLIE_TOTAL").Value)
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End If
End Sub


Private Sub optCredito_Click()
'optNinguna.Enabled = False
dbcTpTarjeta.Visible = True

'frm_VTA_FormaPagoTarjeta.pstrDato = "002"
'frm_VTA_FormaPagoTarjeta.pstrDatoDes = "TARJETA"
'frm_VTA_FormaPagoTarjeta.Show
'frm_VTA_FormaPagoTarjeta.txtNroTar.SetFocus
If flgF6 = 1 Then
'flgF6 = 0
If (frm_VTA_FormaPagoEfectivo.Visible = True) Then
    frm_VTA_FormaPagoEfectivo.Hide
    frm_VTA_FormaPagoTarjeta.pstrDato = "002"
    frm_VTA_FormaPagoTarjeta.pstrDatoDes = "TARJETA"
    frm_VTA_FormaPagoTarjeta.Show
    frm_VTA_FormaPagoTarjeta.SetFocus
End If
End If

If dbcTpTarjeta.Text <> "" Then
    Cal_Promo
    Cal_Montos
End If
End Sub

Private Sub optEfectivo_Click()
'optNinguna.Enabled = False
dbcTpTarjeta.Visible = False
'PagoEfectivo (lblTotal.Caption)
'frm_VTA_FormaPagoEfectivo.pstrDato = "001"
'frm_VTA_FormaPagoEfectivo.pstrDatoDes = "EFECTIVO"
'frm_VTA_FormaPagoEfectivo.Show
'frm_VTA_FormaPagoEfectivo.grdEfectivo.SetFocus
'Cal_Montos
If flgF6 = 1 Then
'flgF6 = 0
If (frm_VTA_FormaPagoTarjeta.Visible = True) Then
    'frm_VTA_FormaPagoTarjeta.Hide
     Unload frm_VTA_FormaPagoTarjeta
     frm_VTA_FormaPagoEfectivo.carga True
    'frm_VTA_FormaPagoEfectivo.pstrDato = "001"
    'frm_VTA_FormaPagoEfectivo.pstrDatoDes = "EFECTIVO"
    'frm_VTA_FormaPagoEfectivo.Show
    'frm_VTA_FormaPagoEfectivo.grdEfectivo.SetFocus
End If
End If
Cal_Promo
Cal_Montos
End Sub

Private Sub optNinguna_Click()
dbcTpTarjeta.Visible = False
'Me.optNinguna.SetFocus
End Sub

Private Sub PagoEfectivo(Monto As Double)
'objVenta.IniFormaPago
If Monto > 0 Then
'    objVenta.AgregaFormaPago "001", _
'                                 "EFECTIVO", _
'                                 "001", _
'                                 "EFEC-SOLES", _
'                                 monto, _
'                                 "", _
'                                 "1", _
'                                 "", _
'                                 "", _
'                                 "", _
'                                 "", _
'                                 objUsuario.TipoCambio, "", _
'                                 "", "", "", "", "", "", "", _
'                                 "", "", "", "", "", "", "", _
'                                 "", "", 0, "", "", "0.0"


End If
End Sub

Private Sub PagoTarjeta(Monto As Double)
'objVenta.IniFormaPago
Dim cod, desc  As String
Dim arrValores() As String
arrValores = Split(dbcTpTarjeta.Text, " ")

If UBound(arrValores) <= 0 Then Exit Sub

cod = arrValores(0)
desc = arrValores(1)
If Monto > 0 Then
    objVenta.AgregaFormaPago "002", _
                             "TARJETA", _
                             cod, _
                             desc, _
                             Monto, _
                             "0", _
                             "1", "", _
                             "", "", _
                             "", objUsuario.TipoCambio, _
                             "", _
                             "1", _
                             "", _
                             "1", _
                             "", "", _
                             "", "", _
                             "", "", _
                             "", "", _
                             "", "", _
                             "", "", _
                             "", 0#, _
                             "", "", "", ""
End If
End Sub

Public Sub Recalcular()
If optEfectivo.Value = True Then
    PagoEfectivo (lblTotal.Caption)
    Cal_Montos
End If
If optCredito.Value = True Then
    PagoTarjeta (lblTotal.Caption)
    Cal_Montos
End If
End Sub

Public Sub loadOptions()
    frmPedido.optCredito.Visible = True
    frmPedido.optCredito.Value = False
    frmPedido.optEfectivo.Visible = True
    frmPedido.optEfectivo.Value = True
'    frmPedido.optNinguna.Visible = True
'    frmPedido.optNinguna.Value = True
End Sub

Public Sub cancelaOptions()
    frmPedido.optCredito.Visible = False
    frmPedido.optEfectivo.Visible = False
    objVenta.vExisteDNI_RENIEC = ""
    frmPedido.lblDniInvalido.Visible = False
'    frmPedido.optNinguna.Visible = False
    Me.dbcTpTarjeta.Visible = False
End Sub

Public Sub OptionsFocus()
If Me.optCredito.Visible = True And Me.optEfectivo.Visible = True Then
    If Me.optCredito.Value = True Then
        Me.optCredito.SetFocus
'    ElseIf Me.optNinguna.Value = True Then
'        Me.optNinguna.SetFocus
    ElseIf Me.optEfectivo.Value = True Then
        Me.optEfectivo.SetFocus
    End If
End If
End Sub

Private Sub txtCMPBus_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim vRes As String
 Dim vDatos As String
 Dim vArray() As String

If txtCMPBus.Text <> "" Then
   If KeyCode = vbKeyReturn Then
'        lblMedico.Caption = objVenta.NombreMedico(txtCMPBus.Text)
'        If lblMedico.Caption = "" Then
'            MsgBox "El CMP no existe", vbCritical, App.ProductName
'            Exit Sub
'        End If
'    SendKeys "{F1}"
        vRes = Trim(objMedico.buscaCMP(Trim(txtCMPBus.Text)))
           
        If vRes = "N" Then
           ' No existe el Medico
           objMedico.vCodMedico = ""
           lblMedico.Caption = ""
           objVenta.CodMedico = ""
           frmPedido.Cal_Promo
           frmPedido.Cal_Montos
           frmPedido.grdPedido.Refresh
           frmPedido.txtCMPBus.Focus
           MsgBox "El CMP no existe", vbCritical, App.ProductName
           Exit Sub
        Else
            If vRes = "M" Then
               'Varios medicos muestra pantalla de seleccion
               'MsgBox "Muestra pantalla para q seleccione", vbCritical, App.ProductName
                frm_VTA_BusquedaMedico.vNumCMP_Ingresado = Trim(txtCMPBus.Text)
                frm_VTA_BusquedaMedico.Show vbModal
                objVenta.CodMedico = objMedico.vCodMedico
            Else
               ' Es un Solo medico
               vDatos = objMedico.getDatosMedico(vRes)
               vArray = Split(vDatos, "@")
               lblMedico.Caption = "" & vArray(1)
               objMedico.vCodMedico = "" & vArray(3)
               objVenta.CodMedico = objMedico.vCodMedico
            End If
            frmPedido.Cal_Promo
            frmPedido.Cal_Montos
            frmPedido.grdPedido.Refresh
        End If
   End If
End If

End Sub

