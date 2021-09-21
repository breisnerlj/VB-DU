VERSION 5.00
Begin VB.Form frm_DLV_GrbTipoDireccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Dirección"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frm_DLV_GrbTipoDireccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmrGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   4200
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin vbp_Ventas.ctlTextBox TXTDes 
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   609
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
      Begin vbp_Ventas.ctlTextBox TxtCod 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
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
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_DLV_GrbTipoDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDirecc As New clsDireccDLV
Public pstrCod As String
Public pstrDes As String
Public pstrActivo As String
Dim strActivo As String

Private Sub cmrGrabar_Click()

    
    If TxtCod.Text = "" Then MsgBox "Ingrese su Codigo": TxtCod.SetFocus: Exit Sub
    If TXTDes.Text = "" Then MsgBox "Ingrese su Descripción": TXTDes.SetFocus: Exit Sub
    
    strActivo = IIf(chkActivo.Value = 1, "1", "0")
    
    Dim CtrlErr As String
    CtrlErr = objDirecc.Graba(TxtCod.Text, TXTDes.Text, strActivo, objUsuario.Codigo)

    If CtrlErr = "" Then
        MsgBox "Se Grabo con exito la Dirección", vbInformation, Caption
        Set frm_DLV_LstTipoDireccion.grdDireccDlv.DataSource = objDirecc.ListaDirecc
        Unload Me
    Else
        MsgBox CtrlErr, vbCritical, Caption
    End If
End Sub

Private Sub Form_Load()

    If pstrCod <> "" Then
        TxtCod.Text = pstrCod
        TxtCod.Enabled = False
        TXTDes.Text = pstrDes
        chkActivo.Value = pstrActivo
     Else
         TxtCod.Text = ""
         TxtCod.Enabled = True
         TXTDes.Text = ""
         chkActivo.Value = 0
    End If
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
    TxtCod.Tipo = Entero
End Sub

Private Sub TXTDes_KeyPress(KeyAscii As Integer)
    TXTDes.Tipo = Mayusculas
End Sub
