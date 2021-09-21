VERSION 5.00
Begin VB.Form frm_SERV_Impresoras 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresoras"
   ClientHeight    =   1140
   ClientLeft      =   6345
   ClientTop       =   4770
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Icon            =   "frm_SERV_Impresoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   705
      Width           =   375
   End
   Begin VB.CommandButton comCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton comAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   660
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   750
      Width           =   1215
   End
End
Attribute VB_Name = "frm_SERV_Impresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Impresora As Printer
Dim ContImpr  As Byte
Dim NameFile As String
Dim wStr, Devicename As String
Dim UbicaPrinter, StrUbica As String
Public gNroCopia As Integer

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text1.SetFocus
        SendKeys "+{home}", True
    End If
End Sub

Public Sub UbicaImpresora(ByVal StrImp As String)
    StrUbica = StrImp
End Sub

Private Sub Form_Load()
    Dim Impresoras As Printer
    Dim i As Byte
    Dim Ubica As Byte
    UbicaPrinter = StrUbica
    ContImpr = 0
    Ubica = 0
    For Each Impresoras In Printers
        Combo1.AddItem Impresoras.Devicename
        ContImpr = ContImpr + 1
        If UbicaPrinter <> "" Then
            If UbicaPrinter = Impresoras.Devicename Then
                  Ubica = ContImpr
            End If
        End If
    Next
    If ContImpr = 0 Then
        comAceptar.Enabled = False
        comCancelar.Enabled = True
    Else
        If Ubica = 0 Then
            Combo1.ListIndex = 0
        Else
            Combo1.ListIndex = Ubica - 1
            Combo1.Enabled = False
            comCancelar.Enabled = False
        End If
    End If
    StrUbica = ""
    UbicaPrinter = ""
End Sub

Private Sub comaceptar_Click()
  gNroCopia = IIf(Val(Text1.Text) = 0, 0, Val(Text1.Text))
  Devicename = Combo1.Text
  For Each Impresora In Printers
      If Impresora.Devicename = Devicename Then
      On Error GoTo handle
          'Dim antigua As Printer
'          Set antigua = Printer
          Set Printer = Impresora
           Printer.Print "."
            Printer.KillDoc
         Exit For
      End If
  Next Impresora
  
  'Call Nombre_Impresora
  'frm_SERV_Impresoras.Hide
  Unload frm_SERV_Impresoras
  Exit Sub
handle:
  MsgBox "La impresora " & Impresora.Devicename & ", no se encuentra disponible", vbCritical, App.ProductName
'  Set Printer = antigua
End Sub

Private Sub comcancelar_Click()
   gNroCopia = 0
   'frm_SERV_Impresoras.Hide
   Unload frm_SERV_Impresoras
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        comAceptar.SetFocus
    End If
End Sub


'Private Function Nombre_Impresora() As String
'    Nombre_Impresora = Trim(Devicename)
'End Function

