VERSION 5.00
Begin VB.Form frm_ADM_AgregaCantidad2 
   Caption         =   "Agregar"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "[Enter] Aceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "[Esc] Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Tipo            =   3
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
   Begin VB.Label lblUnidad 
      Caption         =   "des_unidad"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblProducto 
      Caption         =   "des_Producto"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad : "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Unidad : "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Producto : "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ADM_AgregaCantidad2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj As New clsEntrega
Public strEntrega, codproducto, desProducto, desUnidad As String

Sub agregar()
If Me.txtCantidad.Text = "" Then
    MsgBox "Debe ingresar cantidad.", vbCritical + vbInformation, "Error"
Else
    frm_ADM_Conteo2.recibe codproducto, Trim(Me.txtCantidad.Text)
End If
End Sub

Private Sub cmdAceptar_Click()
    agregar
    'frm_ADM_Conteo2.recibe strEntrega, codproducto, Trim(Me.txtCantidad.Text)
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    agregar
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    Me.lblProducto.Caption = desProducto
    Me.lblUnidad.Caption = desUnidad
End Sub
