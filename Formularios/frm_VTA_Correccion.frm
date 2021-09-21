VERSION 5.00
Begin VB.Form frm_VTA_Correccion 
   BorderStyle     =   0  'None
   Caption         =   "Administrador"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5565
      Picture         =   "frm_VTA_Correccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4260
      Picture         =   "frm_VTA_Correccion.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
   End
   Begin vbp_Ventas.ctlDataCombo dbcTipoDoc 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   2295
      _ExtentX        =   2566
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox txtDice 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2143
      _ExtentY        =   661
      Tipo            =   7
      Alignment       =   2
      MaxLength       =   11
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
   Begin vbp_Ventas.ctlTextBox txtDebe 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2143
      _ExtentY        =   661
      Tipo            =   7
      Alignment       =   2
      MaxLength       =   11
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
   Begin vbp_Ventas.ctlTextBox txtDiceFin 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2143
      _ExtentY        =   661
      Tipo            =   7
      Alignment       =   2
      MaxLength       =   11
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
   Begin vbp_Ventas.ctlTextBox txtDebeFin 
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2566
      _ExtentY        =   661
      Tipo            =   7
      Alignment       =   2
      MaxLength       =   11
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
   Begin VB.CheckBox chkhabilita 
      Caption         =   "&Habilitar un grupo de registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   6240
      Picture         =   "frm_VTA_Correccion.frx":0B14
      Top             =   120
      Width           =   240
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   240
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F3"
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
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
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
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
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
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Corrección de Correlativos"
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
      Height          =   270
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   2895
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
      Left            =   5910
      TabIndex        =   12
      Top             =   6420
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
      Left            =   4200
      TabIndex        =   11
      Top             =   6420
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DEBE DECIR"
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
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DICE"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Documento"
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
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   1785
   End
End
Attribute VB_Name = "frm_VTA_Correccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento


Private Sub chkhabilita_Click()
    If chkhabilita.Value = vbChecked Then
        txtDiceFin.Enabled = True
        txtDebeFin.Enabled = True
        txtDiceFin.SetFocus
        
    End If
    If chkhabilita.Value = vbUnchecked Then
        txtDiceFin.Enabled = False
        txtDebeFin.Enabled = False
        txtDice.SetFocus
    End If
    
End Sub

Private Sub cmdAceptar_Click()
Dim strIni As Variant
Dim strFin As Variant
Dim strDice As String
Dim strDebedecir As String
Dim strDesde As Variant
Dim incrementar As Double
Dim Correlativo As Variant
    On Error GoTo CtrlErr
    If chkhabilita.Value = vbUnchecked Then
        
        If dbcTipoDoc.BoundText = "" Then MsgBox "Debe seleccionar el tipo de documento.", vbOKOnly + vbExclamation, "Error": dbcTipoDoc.SetFocus: Exit Sub
        If Trim(txtDice.Text) = "" Then MsgBox "Debe ingresar el documento DICE.", vbOKOnly + vbExclamation, "Error": txtDice.SetFocus: Exit Sub
        If Trim(txtDebe.Text) = "" Then MsgBox "Debe ingresar el documento DEBE.", vbOKOnly + vbExclamation, "Error": txtDebe.SetFocus: Exit Sub
        
        strDice = Replace(txtDice.Text, "-", "")
        strDebedecir = Replace(txtDebe.Text, "-", "")
        Correlativo = objDocumento.Correlativo(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, dbcTipoDoc.BoundText, strDice, strDebedecir, objUsuario.Codigo)
        If Correlativo <> "" Then
            MsgBox Correlativo, vbInformation, App.ProductName
            Exit Sub
        End If
        
    End If
    
    If chkhabilita.Value = vbChecked Then
               strDebedecir = Replace(txtDebe.Text, "-", "")
               strIni = Replace(txtDice.Text, "-", "")
               strFin = Replace(txtDiceFin.Text, "-", "")
               For strDesde = strIni To strFin
                  strDice = Format(strDesde, "0000000000")
                  Correlativo = objDocumento.Correlativo(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, dbcTipoDoc.BoundText, strDice, strDebedecir, objUsuario.Codigo)
                    If Correlativo <> "" Then
                        MsgBox Correlativo, vbCritical, App.ProductName
                        Exit For
                    End If
                  incrementar = Val(strDebedecir) + 1
                  strDebedecir = Format(incrementar, "0000000000")
               Next
    End If
    
    If Correlativo = "" Then
        MsgBox "SE CAMBIO EL NUMERO DEL DOCUMENTO", vbInformation, App.ProductName
    Else
        MsgBox Correlativo, vbCritical, App.ProductName
        dbcTipoDoc.SetFocus
        Exit Sub
    End If
    
    
    'Unload Me
    
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dbcTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
                Case vbKeyEscape
                    cmdCancelar_Click

    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            dbcTipoDoc.SetFocus
        Case vbKeyF2
            txtDice.SetFocus
        Case vbKeyF3
            chkhabilita.SetFocus
        Case vbKeyF4
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyH
            If chkhabilita.Value = vbChecked Then
                    chkhabilita.Value = vbUnchecked
            Else
                    chkhabilita.Value = vbChecked
            End If
    End Select
End Sub

Private Sub Form_Load()
SetteaFormulario Me

        Me.top = 0: Me.left = 0

        Set dbcTipoDoc.RowSource = objDocumento.ListaTipo(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
        dbcTipoDoc.ListField = "DESCRIPCION"
        dbcTipoDoc.BoundColumn = "CODIGO"
        chkhabilita.Value = vbUnchecked
        txtDiceFin.Enabled = False
        txtDebeFin.Enabled = False
End Sub
