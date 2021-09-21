VERSION 5.00
Begin VB.Form frm_ADM_Transportista 
   Caption         =   "Datos del Transportista"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlTextBox txtPlaca 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Tipo            =   2
      MaxLength       =   20
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "[Esc] Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "[F11] Aceptar "
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin vbp_Ventas.ctlTextBox txtGlosa 
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   3615
      _ExtentX        =   7858
      _ExtentY        =   450
      Tipo            =   8
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
   Begin vbp_Ventas.ctlTextBox txtChofer 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
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
   Begin vbp_Ventas.ctlTextBox txtBultos 
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Tipo            =   3
      Alignment       =   2
      MaxLength       =   5
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
   Begin vbp_Ventas.ctlTextBox txtPrecintos 
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Tipo            =   3
      Alignment       =   2
      MaxLength       =   5
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
   Begin vbp_Ventas.ctlTextBox txtentrega 
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Tipo            =   3
      Alignment       =   2
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
   Begin vbp_Ventas.ctlTextBox txtIdTransportista 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Tipo            =   2
      Enabled         =   0   'False
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
   Begin vbp_Ventas.ctlTextBox txtDesTransportista 
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empresa Transporte :"
      Height          =   195
      Left            =   150
      TabIndex        =   13
      Top             =   240
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Transportista/Chofer :"
      Height          =   195
      Left            =   150
      TabIndex        =   12
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Placa Unidad :"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Cant. Bultos :"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Cant. Precintos :"
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Glosa :"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1485
   End
End
Attribute VB_Name = "frm_ADM_Transportista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strIdEntrega As String
Dim objEntrega As New clsEntrega

Public Sub carga(idEntrega As String)
    strIdEntrega = idEntrega
    'CargaCombos
    Me.Show vbModal
End Sub

'Sub CargaCombos()
'Set cboTransportista.RowSource = objEntrega.ListaTransportista("", "1", "Seleccionar")
'    cboTransportista.BoundColumn = "ID_TRANSPORTISTA"
'    cboTransportista.ListField = "DES_TRANSPORTISTA"
'    cboTransportista.Text = "Seleccionar"
'End Sub

Private Sub cmdAceptar_Click()
On Error GoTo CtrlErr
    If Me.txtIdTransportista.Text = "" And Me.txtDesTransportista.Text = "" Then
        MsgBox "Falta Seleccionar el Transportista", vbCritical, "Error"
        Me.txtDesTransportista.SetFocus
        Exit Sub
    End If
    If Trim(txtChofer.Text) = "" Then
        MsgBox "Falta Ingresar el Transportista/Chofer", vbCritical, "Error"
        txtChofer.SetFocus
        Exit Sub
    End If
    If Trim(txtPlaca.Text) = "" Then
        MsgBox "Falta Ingresar la Placa", vbCritical, "Error"
        txtPlaca.SetFocus
        Exit Sub
    End If
    If Trim(txtBultos.Text) = "" Or Trim(txtBultos.Text) = "0" Then
        MsgBox "Falta Ingresar el Numero de Bultos", vbCritical, "Error"
        txtBultos.SetFocus
        Exit Sub
    End If
    If Trim(txtPrecintos.Text) = "" Then
        MsgBox "Falta Ingresar el Numero de Precintos", vbCritical, "Error"
        txtPrecintos.SetFocus
        Exit Sub
    End If
    If txtentrega.Text = "" Then
        Dim msbo As Variant
        Dim identregax As String
        msbo = MsgBox("¿Seguro que desea Registrar el Transportista?", vbYesNo + vbInformation, App.ProductName)
        If msbo = vbYes Then
            txtentrega.Text = Graba
            identregax = Me.txtentrega.Text
            frm_ADM_Entrega.fnImprimeTransportista (identregax)
            Unload Me
            frm_ADM_TranspGuia.carga identregax
        End If
    End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Function Graba() As String
    Dim Entrega As String
    objEntrega.GrabaTransportista Entrega, objUsuario.Codigo, objUsuario.CodigoLocal, Me.txtIdTransportista.Text, Trim(txtChofer.Text), Trim(txtPlaca.Text), Val(txtBultos.Text), Trim(txtPrecintos.Text), Trim(txtGlosa.Text)
    Graba = Entrega
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 122 Then
    cmdAceptar_Click
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtDesTransportista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frm_ADM_TransportistaBuscar.Show vbModal
End If
End Sub

Public Sub recibe()
    Unload frm_ADM_TransportistaBuscar
    Me.txtChofer.SetFocus
End Sub
