VERSION 5.00
Begin VB.UserControl ctlDatosCliente 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   ScaleHeight     =   945
   ScaleWidth      =   9360
   Begin VB.CheckBox chkVerificado 
      Caption         =   "Verificado"
      Enabled         =   0   'False
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
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "&Natural"
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   4
      Top             =   0
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "&Juridico"
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin vbp_Ventas.ctlTextBox txtSufijo 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
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
   Begin vbp_Ventas.ctlTextBox txtDespacho 
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin vbp_Ventas.ctlTextBox txtLocal 
      Height          =   315
      Left            =   3180
      TabIndex        =   2
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin vbp_Ventas.ctlTextBox txtApeMaterno 
      Height          =   315
      Left            =   5040
      TabIndex        =   6
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
   Begin vbp_Ventas.ctlTextBox txtNombre 
      Height          =   315
      Left            =   960
      TabIndex        =   7
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
   Begin vbp_Ventas.ctlTextBox txtApellido 
      Height          =   315
      Left            =   2760
      TabIndex        =   8
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
   Begin vbp_Ventas.ctlTextBox txtCodigo 
      Height          =   315
      Left            =   600
      TabIndex        =   9
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   195
      Left            =   1200
      TabIndex        =   15
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
      Height          =   195
      Left            =   2880
      TabIndex        =   14
      Top             =   360
      Width           =   555
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido Materno"
      Height          =   195
      Left            =   5160
      TabIndex        =   13
      Top             =   360
      Width           =   1185
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
      Index           =   6
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   600
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio del Local"
      Height          =   195
      Left            =   1960
      TabIndex        =   11
      Top             =   60
      Width           =   1140
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Despacho"
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "ctlDatosCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strLocalAsignado As String
Dim strLocalDespacho As String

Public Sub Cargar(ByVal strCodigo As String, _
                    ByVal strLocal As String, _
                    ByVal strDespacho As String, _
                    ByVal intTipo As Integer, _
                    ByVal strRazonSocial As String, _
                    ByVal strNomComercial As String, _
                    ByVal strNomCliente As String, _
                    ByVal strApeCliente As String, _
                    ByVal strApe2Cliente As String, _
                    ByVal intVerificado As Integer, _
                    ByVal strSufijo As String)

    txtCodigo.Text = strCodigo
    txtLocal.Text = strLocal
    txtDespacho.Text = strDespacho
    optTipo(intTipo).Value = True
    txtApeMaterno.Visible = True
    Label27.Visible = True
    chkVerificado.Value = intVerificado
    
    If intTipo = 1 Then
        txtNombre.Text = strRazonSocial
        txtApellido.Text = strNomComercial
        Label2.Caption = "Razón Social"
        Label3.Caption = "Razón Comercial"
        txtApeMaterno.Visible = False
        Label27.Visible = False
    Else
        txtNombre.Text = strNomCliente
        txtApellido.Text = strApeCliente
        txtApeMaterno.Text = strApe2Cliente
        Label2.Caption = "Nombre"
        Label3.Caption = "Apellido"
    End If
    txtSufijo.Text = strSufijo

End Sub


Public Sub Limpiar()
    txtCodigo.Text = ""
    txtLocal.Text = ""
    txtDespacho.Text = ""
    optTipo(0).Value = True
    txtApeMaterno.Visible = True
    Label27.Visible = True
    chkVerificado.Value = 0
    txtNombre.Text = ""
    txtApellido.Text = ""
    txtApeMaterno.Text = ""
    Label2.Caption = "Nombre"
    Label3.Caption = "Apellido"
    txtSufijo.Text = ""
End Sub

Public Property Get LocalAsignado() As String
    LocalAsignado = strLocalAsignado

End Property

Public Property Let LocalAsignado(ByVal lstrLocalAsignado As String)
    strLocalAsignado = lstrLocalAsignado
End Property

Public Property Get LocalDespacho() As String
    LocalDespacho = strLocalDespacho
End Property

Public Property Let LocalDespacho(ByVal lstrLocalDespacho As String)
    strLocalDespacho = lstrLocalDespacho
End Property
