VERSION 5.00
Begin VB.Form frm_VTA_EspeciesValoradas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de la Especie Valorada"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   4620
      Picture         =   "frm_VTA_EspeciesValoradas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   3300
      Picture         =   "frm_VTA_EspeciesValoradas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de SOAT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      Begin vbp_Ventas.ctlTextBox txtFormulario 
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Tipo            =   8
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
      Begin vbp_Ventas.ctlTextBox txtPLaca 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Tipo            =   8
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Formulario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   427
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   907
         Width           =   525
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   11
      Left            =   5025
      TabIndex        =   8
      Top             =   2580
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   12
      Left            =   3240
      TabIndex        =   7
      Top             =   2580
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_EspeciesValoradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodigoProducto  As String
Public bolCancel As Boolean
Dim objDocumento As New clsDocumento
Dim rsAuxControlSOAT As oraDynaset

Private Sub cmdAceptar_Click()
On Error GoTo handle

    Set rsAuxControlSOAT = objDocumento.ListaControlSOAT(txtFormulario.Text, objUsuario.CodigoLocal)
    
    If rsAuxControlSOAT.EOF Then
        MsgBox "El Formulario Nº " & txtFormulario.Text & ", no existe ", vbCritical, App.ProductName
        txtFormulario.SetFocus
        Exit Sub
    Else
                
        If rsAuxControlSOAT("FLG_VENDIDO").Value = "1" Then
            MsgBox "El Formulario Nº " & txtFormulario.Text & ", esta vendido", vbInformation, App.ProductName
            txtFormulario.SetFocus
            Exit Sub
        End If
        
        If rsAuxControlSOAT("FLG_ANULADO").Value = "1" Then
            MsgBox "El Formulario Nº " & txtFormulario.Text & ", esta anulado", vbInformation, App.ProductName
            txtFormulario.SetFocus
            Exit Sub
        End If
        
        If rsAuxControlSOAT("FLG_DEVUELTO").Value = "1" Then
            MsgBox "El Formulario Nº " & txtFormulario.Text & ", esta devuelto", vbInformation, App.ProductName
            txtFormulario.SetFocus
            Exit Sub
        End If
        
        
        objVenta.AgregaEspecieValorada strCodigoProducto, txtFormulario.Text, txtPLaca.Text
        bolCancel = False
    End If
    

    Unload Me
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdCancelar_Click()
    bolCancel = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click
End Sub


Public Sub Datos(ByVal intItem As Integer)

Me.Caption = Me.Caption & Str(intItem)

End Sub


