VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_OFF_CheckDate 
   Caption         =   "Fecha del Sistema"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3660
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   350
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "&Continuar"
      Height          =   350
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Tag             =   "Ingrese la fecha actual."
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   57147393
      CurrentDate     =   39448
      MinDate         =   36526
   End
   Begin MSComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2340
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la fecha actual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   3660
      Y1              =   550
      Y2              =   550
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BIENVENIDO AL SISTEMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CONTINGENCIA DE VENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   290
      Width           =   3615
   End
End
Attribute VB_Name = "frm_OFF_CheckDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Result As VbMsgBoxResult

Public Function CheckDate() As VbMsgBoxResult
    Me.Show vbModal
    CheckDate = Result
End Function

Private Sub cmdIngresar_Click()
    '..comprueba si el valor de value es Null
    If Not IsNull(dtpDate.Value) Then
        If CDate(Format(dtpDate.Value, "dd/mm/yyyy")) <> CDate(Format(Now, "dd/mm/yyyy")) Then
            MsgBox "La fecha actual del sistema es diferente a la fecha ingresada." & vbNewLine & _
                   "Modifique la fecha del sistema o comuníquese con soporte técnico.", vbCritical, App.ProductName
            dtpDate.SetFocus
        Else
            Result = vbYes
            Unload Me
        End If
    Else
        MsgBox "Debe seleccionar una fecha.", vbCritical, App.ProductName
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtpDate_GotFocus()
    stbPrincipal.Panels(1).Text = dtpDate.Tag
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdIngresar.SetFocus
    End If
End Sub

Private Sub dtpDate_LostFocus()
    stbPrincipal.Panels(1).Text = vbNullString
End Sub

Private Sub Form_Load()
    dtpDate.Value = Null
    Result = vbNo
End Sub
