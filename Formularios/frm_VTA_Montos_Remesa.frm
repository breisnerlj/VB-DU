VERSION 5.00
Begin VB.Form frm_VTA_Montos_Remesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Remesa"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin vbp_Ventas.ctlTextBox TxtSoles 
         Height          =   325
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin vbp_Ventas.ctlTextBox TxtDolares 
         Height          =   325
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Caption         =   "Dolares $"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Soles  S/."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   1680
      TabIndex        =   6
      Top             =   1425
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   225
   End
End
Attribute VB_Name = "frm_VTA_Montos_Remesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objRemesa As New clsRemesa
Dim dblValSol As Double
Dim dblValDol As Double

Private Sub Form_Load()
    If frm_VTA_Remesas.pColR = 5 Then
        TxtSoles.Enabled = True
        TxtSoles.Text = objLiquidacion.Remesa(frm_VTA_Remesas.grdRemesa.Row, 5)
        TxtDolares.Enabled = False
    End If
    If frm_VTA_Remesas.pColR = 6 Then
        TxtDolares.Enabled = True
        TxtDolares.Text = objLiquidacion.Remesa(frm_VTA_Remesas.grdRemesa.Row, 6)
        TxtSoles.Enabled = False
    End If
End Sub

Private Sub TxtDolares_KeyPress(KeyAscii As Integer)
    On Error GoTo CtrlErr
    TxtDolares.Tipo = Real
    If KeyAscii = 13 Then
        objLiquidacion.Remesa(frm_VTA_Remesas.grdRemesa.Row, 6) = TxtDolares.Text
                
        'frm_VTA_Remesas.LblSobreDolRemesa.Caption = objRemesa.DevSecDol(objUsuario.NombrePc, objUsuario.CodigoEmpresa)
        'Pasando al arreglo el numero de secuencia en dolares'
        'objLiquidacion.Remesa(1, 9) = objRemesa.DevSecDol(objUsuario.NombrePc, objUsuario.CodigoEmpresa)
            
        Dim j As Integer
        dblValDol = 0
        For j = 0 To objLiquidacion.Remesa.UpperBound(1)
              dblValDol = dblValDol + objLiquidacion.Remesa(j, 6)
        Next j
        frm_VTA_Remesas.grdRemesa.Columns(6).FooterText = dblValDol
            
        frm_VTA_Remesas.grdRemesa.Rebind
        Unload Me
    End If
    Exit Sub
CtrlErr:
        Err.Raise Err.Number, "Error en el dato", Err.Description
    
End Sub

Private Sub TxtSoles_KeyPress(KeyAscii As Integer)
    TxtSoles.Tipo = Real
    If KeyAscii = 13 Then
       objLiquidacion.Remesa(frm_VTA_Remesas.grdRemesa.Row, 5) = TxtSoles.Text

       'frm_VTA_Remesas.LblSobreSolRemesa.Caption = objRemesa.DevSecSol(objUsuario.NombrePc, objUsuario.CodigoEmpresa)
       'Pasando al arreglo el numero de secuencia en soles'
       'objLiquidacion.Remesa(0, 8) = objRemesa.DevSecSol(objUsuario.NombrePc, objUsuario.CodigoEmpresa)
            
       Dim i As Integer
       dblValSol = 0
       For i = 0 To objLiquidacion.Remesa.UpperBound(1)
             dblValSol = dblValSol + objLiquidacion.Remesa(i, 5)
       Next i
       frm_VTA_Remesas.grdRemesa.Columns(5).FooterText = dblValSol
       
       frm_VTA_Remesas.grdRemesa.Rebind
       Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Unload Me
    End Select
End Sub
