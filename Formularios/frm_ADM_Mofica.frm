VERSION 5.00
Begin VB.Form frm_ADM_ModificaServ 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación de Servicios"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmrGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin vbp_Ventas.ctlDataCombo cboServicio 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlDataCombo cboTipoServicio 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox txtSuministro 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   870
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "N° Operación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   870
      Width           =   1455
   End
End
Attribute VB_Name = "frm_ADM_ModificaServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Public strCodigoPadre As String
Public strCodigoHijo As String
Public strTipoDocumento As String
Public strNumeroDocumento As String

Public strNumSuministro As String

Private Sub cboServicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub cboTipoServicio_Change()
    busca_servicio cboTipoServicio.BoundText
End Sub

Private Sub cboTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdSalir_Click

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmrGrabar_Click()
Dim strMensaje As String
    strMensaje = objServicio.Modifica(cboServicio.BoundText, txtSuministro.Text, objUsuario.CodigoEmpresa, strTipoDocumento, strNumeroDocumento, objUsuario.Codigo)
    If strMensaje = "" Then
    MsgBox "Se grabo satisfactorimente ", vbInformation, App.ProductName
    frm_VTA_ConsCobServ.CmdConsultar_Click
    Unload Me
    Else
        MsgBox strMensaje, vbInformation, App.ProductName
    
    End If
End Sub

Private Sub Form_Load()
    Set cboTipoServicio.RowSource = objServicio.ListaTipo
    cboTipoServicio.ListField = "DES_TIPO_SERVICIO"
    cboTipoServicio.BoundColumn = "COD_TIPO_SERVICIO"
    busca_servicio strCodigoPadre
    ''
    cboTipoServicio.BoundText = strCodigoPadre
    cboServicio.BoundText = strCodigoHijo
    txtSuministro.Text = strNumSuministro
End Sub
Sub busca_servicio(strServicio As String)
    cboServicio.BoundText = ""
    Set cboServicio.RowSource = objServicio.Lista("", strServicio)
    cboServicio.ListField = "DES_SERVICIO"
    cboServicio.BoundColumn = "COD_SERVICIO"
End Sub
