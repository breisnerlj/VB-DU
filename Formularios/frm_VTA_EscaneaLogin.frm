VERSION 5.00
Begin VB.Form frm_VTA_EscaneaLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validacion de Afiliado"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   5520
      Picture         =   "frm_VTA_EscaneaLogin.frx":0000
      ScaleHeight     =   3030
      ScaleWidth      =   1815
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1296
      ColorDefault    =   12632319
      ColorDefault    =   12632319
      Tipo            =   3
      Alignment       =   2
      PasswordChar    =   "*"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Por favor escane su tarjeta de Puntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frm_VTA_EscaneaLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tipo As String
Public ok As Boolean
Dim KeyTime As Double
Dim ACABA As Boolean
Private Sub Command1_Click()
    ok = False
    Unload Me
End Sub

Private Sub ctlTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objDocumentoPago As clsDocumentoPago
    
    If EsEscaneado(KeyCode, Shift) = True Then
        If KeyCode = vbKeyReturn Then
            If Tipo = "T" Then
                Set objDocumentoPago = New clsDocumentoPago
                If Not objDocumentoPago.buscaFun(Trim(ctlTextBox1.Text)) = "MONEDERO" Then
                    MsgBox "Tarjeta NO válida", vbCritical + vbOKOnly, "Error"
                    ctlTextBox1.selection
                    Exit Sub
                End If
                objVenta.NumTarjetaPuntos = Trim(ctlTextBox1.Text)
                ok = True
                Unload Me
            ElseIf Tipo = "D" Then
                If Len(Trim(ctlTextBox1.Text)) < 8 Or Len(Trim(ctlTextBox1.Text)) > 10 Then
                    MsgBox "Número de documento NO válido", vbCritical + vbOKOnly, "Error"
                    ctlTextBox1.selection
                    Exit Sub
                End If
                objVenta.NumDNI = Trim(ctlTextBox1.Text)
                ok = True
                Unload Me
            ElseIf Tipo = "A" Then
                If objVenta.NumTarjetaPuntos = "" Then
                    Label1.Caption = "Por favor escane su Tarjeta de Puntos"
                    If Not objDocumentoPago.buscaFun(Trim(ctlTextBox1.Text)) = "MONEDERO" Then
                        MsgBox "Tarjeta NO válida", vbCritical + vbOKOnly, "Error"
                        ctlTextBox1.selection
                        Exit Sub
                    End If
                    objVenta.NumTarjetaPuntos = Trim(ctlTextBox1.Text)
                    ctlTextBox1.Text = ""
                    ctlTextBox1.Focus
                    Exit Sub
                Else
                    If Len(Trim(ctlTextBox1.Text)) < 8 Or Len(Trim(ctlTextBox1.Text)) > 10 Then
                        MsgBox "Número de documento NO válido", vbCritical + vbOKOnly, "Error"
                        ctlTextBox1.selection
                        Exit Sub
                    End If
                    objVenta.NumDNI = Trim(ctlTextBox1.Text)
                    ok = True
                    Unload Me
                End If
            Else
                MsgBox "No tiene implementado este parametro", vbCritical, App.ProductName
                ok = False
                Unload Me
            End If
        End If
    End If
    Set objDocumentoPago = Nothing
End Sub

Private Sub Form_Load()
    Tipo = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "INDFV460", objUsuario.CodigoEmpresa)
    If Tipo = "T" Then
        Label1.Caption = "Por favor escane su Tarjeta de Puntos"
    ElseIf Tipo = "D" Then
        Label1.Caption = "Por favor escane su DNI o Carnet de Extranjeria"
    ElseIf Tipo = "A" Then
        Label1.Caption = "Por favor escane su DNI o Carnet de Extranjeria"
    Else
        MsgBox "No tiene implementado este parametro", vbCritical, App.ProductName
        ok = False
        Unload Me
    End If
End Sub

Function EsEscaneado(KeyCode As Integer, Shift As Integer) As Boolean
Dim DIF As Double
DIF = 0
    Dim ltime As Long
      ltime = timeGetTime
      
If KeyCode = vbKeyReturn Then
    If ACABA = True Then
        ACABA = False
        KeyTime = 0
        EsEscaneado = True
        Exit Function
    Else
        ACABA = False
        KeyTime = 0
        EsEscaneado = False
        Exit Function
    End If
    
End If
   If ACABA = True Then Exit Function
    
    If KeyTime = 0 Then
        KeyTime = Format(Time, "HHmmss") & ltime
        
        DIF = 0
    Else
        DIF = (Format(Time, "HHmmss") & ltime) - KeyTime
        KeyTime = Format(Time, "HHmmss") & ltime
    End If
    
    Debug.Print KeyCode & "DEMORO:" & DIF & "-->" & IIf(DIF < 40, "Escanea", "digita")


If DIF < 40 And DIF <> 0 Then
    ACABA = True
Else
    ACABA = False
End If

End Function
    
