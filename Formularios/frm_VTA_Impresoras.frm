VERSION 5.00
Begin VB.Form frm_VTA_Impresoras 
   Caption         =   "Impresoras"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboImpresoras 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1090
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1090
   End
   Begin VB.TextBox txtNroCopias 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "99"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro de Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_Impresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intImp%
Dim strDeviceName$
Dim strCopias$

Private Sub Form_Load()
    strDeviceName = pfstr_Leer_Cadena_INI("Impresion", "Impresora", "ConfigImp.ini")
    strCopias = pfstr_Leer_Cadena_INI("Impresion", "Copias", "ConfigImp.ini")
    txtNroCopias.Text = CStr(Val(IIf(strCopias = "", "1", strCopias)))
    Call spLlenaCombo(cboImpresoras, strDeviceName)
End Sub

Private Sub cboImpresoras_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancelar_Click()
    gNroCopia = 0
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim objImpresoras As Printer
    strCopias = CStr(Val(txtNroCopias.Text))
    strDeviceName = cboImpresoras.Text
    gNroCopia = 0
    For Each objImpresoras In Printers
        If objImpresoras.Devicename = strDeviceName Then
            Set Printer = objImpresoras
            gNroCopia = Val(strCopias)
            Exit For
        End If
    Next objImpresoras
    frm_VTA_Impresoras.Hide
    Unload Me
End Sub

Private Sub txtNroCopias_GotFocus()
    psub_Selecionar_Todo txtNroCopias
    psub_Foco
End Sub

Private Sub txtNroCopias_KeyPress(KeyAscii As Integer)
    pfint_SoloNumeros KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNroCopias_LostFocus()
    psub_Foco txtNroCopias
End Sub

''''''''''''''''''''***************'''''''''''''''''''''''''''''
Private Sub spLlenaCombo(ByRef rCbo As ComboBox, ByVal vstrDevice$)
Dim intContImpr%
Dim intUbica%
Dim objImpresoras As Printer
    intContImpr = 0
    rCbo.AddItem ""
    For Each objImpresoras In Printers
        intContImpr = intContImpr + 1
        rCbo.AddItem objImpresoras.Devicename, intContImpr
        If vstrDevice <> "" Then
            If objImpresoras.Devicename = vstrDevice Then
                  intUbica = intContImpr
            End If
        End If
    Next
    If intContImpr = 0 Then
        cmdAceptar.Enabled = False
        intImp = -1
    Else
        If vstrDevice = "" Then
            rCbo.ListIndex = -1
        Else
            rCbo.ListIndex = intUbica
        End If
    End If
End Sub



