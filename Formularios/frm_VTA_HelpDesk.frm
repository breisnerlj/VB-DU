VERSION 5.00
Begin VB.Form frm_VTA_HelpDesk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistencia HelpDesk"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMensaje 
      Height          =   2715
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1740
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   5895
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   4320
         TabIndex        =   4
         Top             =   4500
         Width           =   1455
      End
      Begin vbp_Ventas.ctlDataCombo cboCategoria 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboSubCategoria 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   690
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboPrioridad 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1140
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Command1"
         Height          =   315
         Left            =   3300
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3600
         Width           =   15
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje :"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importancia :"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Categoria :"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Categoria :"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   765
      End
   End
End
Attribute VB_Name = "frm_VTA_HelpDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objHelpDesk As New clsHelpDesk

Private Sub cboCategoria_Change()
'   On Error GoTo Control
'cboSubCategoria.BoundText = ""
'
'Set cboSubCategoria.RowSource = objHelpDesk.ListaSubCategoria(cboCategoria.BoundText, "")
'cboSubCategoria.ListField = "DES_SUBCATEGORIA"
'cboSubCategoria.BoundColumn = "COD_SUBCATEGORIA"
'
'   Exit Sub
'
'Control:
'
'      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cboCategoria_Click(Area As Integer)
       On Error GoTo Control
cboSubCategoria.BoundText = ""

Set cboSubCategoria.RowSource = objHelpDesk.ListaSubCategoria(cboCategoria.BoundText, "")
cboSubCategoria.ListField = "DES_SUBCATEGORIA"
cboSubCategoria.BoundColumn = "COD_SUBCATEGORIA"

   Exit Sub

Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cboCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cboPrioridad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cboSubCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdGrabar_Click()
Dim strMensaje As String
    On Error GoTo handle
    If cboCategoria.BoundText = "" Then MsgBox "Debe seleccionar la categoria.", vbOKOnly + vbExclamation, "Mensaje": cboCategoria.SetFocus: Exit Sub
    If cboSubCategoria.BoundText = "" Then MsgBox "Debe seleccionar la sub-categoria.", vbOKOnly + vbExclamation, "Mensaje": cboSubCategoria.SetFocus: Exit Sub
    If cboPrioridad.BoundText = "" Then MsgBox "Debe seleccionar la importancia.", vbOKOnly + vbExclamation, "Mensaje": cboPrioridad.SetFocus: Exit Sub
    If Trim(txtMensaje.Text) = "" Then MsgBox "Debe ingresar el mensaje.", vbOKOnly + vbExclamation, "Mensaje": txtMensaje.SetFocus: Exit Sub
    Dim strNumeroTicket As String
    strMensaje = objHelpDesk.Grabar(objUsuario.CodigoLocal, cboCategoria.BoundText, cboSubCategoria.BoundText, cboPrioridad.BoundText, txtMensaje.Text, objUsuario.Codigo, strNumeroTicket)
    If Not strMensaje = "" Then
        MsgBox "Gracias" & Chr(13) & "Dentro de algunos minutos el Equipo de HelpDesk se comunicara con Usted" & Chr(13) & "Su número de ticket es : " & strMensaje, vbInformation, App.ProductName
        Unload Me
'    Else
'        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'setteaFormulario Me

   On Error GoTo Control

Set cboCategoria.RowSource = objHelpDesk.ListaCategoria("")
cboCategoria.ListField = "DES_CATEGORIA"
cboCategoria.BoundColumn = "COD_CATEGORIA"

Set cboPrioridad.RowSource = objHelpDesk.ListaPrioridad
cboPrioridad.ListField = "DESCRIPCION"
cboPrioridad.BoundColumn = "CODIGO"


Exit Sub
Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub txtMensaje_GotFocus()
txtMensaje.BackColor = cboCategoria.ColorFoco
End Sub

Private Sub txtMensaje_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMensaje_LostFocus()
txtMensaje.BackColor = cboCategoria.ColorDefault
End Sub
