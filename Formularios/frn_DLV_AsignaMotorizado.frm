VERSION 5.00
Begin VB.Form frm_DLV_AsignaMotorizado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asigna Motorizado"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "&Asignar"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin vbp_Ventas.ctlDataCombo cboMotorizado 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label Label16 
      Caption         =   "Motorizado"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm_DLV_AsignaMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strLocal As String
Public strNumProforma As String
Public strLocalPedido As String
Public CodCliente As String
Public CodDireccionCli As String
Private Sub cmdAsignar_Click()
On Error GoTo handle
Dim objProforma As New clsProforma
Dim x As String
x = objProforma.Asigna(objUsuario.CodigoEmpresa, strLocal, strNumProforma, objUsuario.Codigo, cboMotorizado.BoundText, objUsuario.NombrePC, "", strLocalPedido, CodCliente, CodDireccionCli, "", "SI", "", "1")
If x = "" Then
    MsgBox "Se grabo satisfactoriamente", vbExclamation, App.ProductName
    Unload Me
Else
    MsgBox x, vbCritical, App.ProductName
End If
Set objProforma = Nothing
Exit Sub
handle:

    Set objProforma = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()
    Dim objMotorizado As New clsMotorizado
    Set cboMotorizado.RowSource = objMotorizado.ListaDisponible(objUsuario.CodigoLocal, "", "", "", objUsuario.NombrePC)
    cboMotorizado.BoundColumn = "COD_MOTORIZADO"
    cboMotorizado.ListField = "NOMBRE"
    Set objMotorizado = Nothing
End Sub

