VERSION 5.00
Begin VB.Form frm_ADM_AgregaCantidad 
   Caption         =   "Agregar Cantidades"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "[Esc] Cerrar"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "[Enter] Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VB.Label lblCdBarras 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cod. Barras :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frm_ADM_AgregaCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codBarra As String
Public codproducto As String
Public cantidad As String
Public strEntrega As String
Dim obj As New clsEntrega

Sub aceptar()
If Me.txtCantidad.Text <> "" Then
    frm_ADM_Conteo1.recibe codproducto, Trim(Me.txtCantidad.Text)
Else
    MsgBox "Ingrese Cantidad", vbCritical, "Agregar Cantidad"
End If
'If Me.txtCantidad.Text <> "0" Then
''    frm_ADM_Conteo.ctlGrilla1.Columns("cantidad").Value = Me.txtCantidad.Text
''    frm_ADM_Conteo.ctlGrilla1.MoveLast
''    frm_ADM_Conteo.ctlGrilla1.MoveFirst
''    obj.AgregaProducto strEntrega, codProducto, Trim(Me.txtCantidad.Text), "0"
''    Unload Me
'     obj.EditaConteoAux strEntrega, codproducto, Trim(Me.txtCantidad.Text)
'     frm_ADM_Conteo1.CargaGrilla
'     Unload Me
'Else
'    Dim msbo As Variant
'    msgbo = MsgBox("¿Seguro que desea eliminar el producto?", vbYesNo + vbInformation, App.ProductName)
'    If msgbo = vbYes Then
''        frm_ADM_Conteo.eliminarDetalle
''        frm_ADM_Conteo.ctlGrilla1.Rebind
''        obj.AgregaProducto strEntrega, codProducto, "0", "0"
''        Unload Me
'         obj.EliminaConteoAux strEntrega, codproducto
'         frm_ADM_Conteo1.CargaGrilla
'         Unload Me
'    End If
'End If
End Sub

Private Sub Command1_Click()
    aceptar
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        aceptar
    ElseIf KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.lblCdBarras.Caption = codBarra
End Sub

