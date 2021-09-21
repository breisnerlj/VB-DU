VERSION 5.00
Begin VB.Form frm_DLV_ObservacionMotorizados 
   Caption         =   "Observaciones para el Motorizado"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6645
   StartUpPosition =   1  'CenterOwner
   Begin vbp_Ventas.ctlTextBox txtObservacion 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1508
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
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Observaciones de los Motorizados"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2430
   End
End
Attribute VB_Name = "frm_DLV_ObservacionMotorizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objObservacion As New clsMotorizado
Private strCodigo As String
Private strFecha As String

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let Codigo(ByVal vNewValue As String)
    strCodigo = vNewValue
End Property

Public Property Get Fecha() As String
    Fecha = strFecha
End Property

Public Property Let Fecha(ByVal vNewValue As String)
    strFecha = vNewValue
End Property

Private Sub Form_Activate()
    txtObservacion.SetFocus
End Sub

Private Sub cmdGrabar_Click()
    
    On Error GoTo CtrlErr
    
    Dim strMensaje As String
 
    If txtObservacion.Text = "" Then
        MsgBox "Debe ingresar alguna observacion", vbInformation, App.ProductName
        txtObservacion.SetFocus
        Exit Sub
    End If
    
    strMensaje = objObservacion.GrabaObservaciones(Codigo, Fecha, txtObservacion.Text)
           
    If strMensaje = "" Then
        MsgBox "Se actualizó satisfactoriamente", vbInformation, App.ProductName
        Unload Me
    Else
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub
