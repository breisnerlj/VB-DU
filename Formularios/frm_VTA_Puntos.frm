VERSION 5.00
Begin VB.Form frm_VTA_Puntos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puntos a canjear"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      Tipo            =   3
      Alignment       =   2
      MaxLength       =   4
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
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usted tiene un saldo de:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ingrese los puntos a canjear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frm_VTA_Puntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim multiplo As Double, minimo As Double, conver As Double
    
    multiplo = Val(gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "INDFV682", objUsuario.CodigoEmpresa))
    minimo = Val(gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "INDFV411", objUsuario.CodigoEmpresa))
    conver = Val(gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "INDFV482", objUsuario.CodigoEmpresa))

    If KeyCode = vbKeyReturn Then
        If Val(Label3.Caption) = 0 Or Val(Label3.Caption) < Val(ctlTextBox1.Text) Then
            MsgBox "No cuenta con suficientes puntos acumulados para realizar esta acción", vbCritical + vbOKOnly, "Error: Puntos insuficientes"
            ctlTextBox1.selection
            Exit Sub
        End If
        If ((Val(ctlTextBox1.Text)) < Val((minimo / conver))) Then
            MsgBox "Como minimo debe de redimir màs de " & (minimo / conver) & " puntos", vbCritical, App.ProductName
            ctlTextBox1.selection
            Exit Sub
        End If
        If (Val(ctlTextBox1.Text) - Int(Val(ctlTextBox1.Text) / multiplo) * multiplo) <> 0 Then
            MsgBox "Debe de seleccionar en multiplos de " & (multiplo) & " puntos", vbCritical, App.ProductName
            ctlTextBox1.selection
            Exit Sub
        End If
    
        objVenta.MontoRedime = Val(ctlTextBox1.Text) / (multiplo * multiplo)
        frmPedido.lblPuntosRed.Caption = Val(ctlTextBox1.Text)
        frmPedido.Cal_Promo
        frmPedido.Cal_Montos
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Dim oFP As clsFarmaPuntos, oFPC As clsFPConstante
    Dim temp() As String, Tipo As String
    Dim vDni As String
    
    On Error GoTo ControlError
    
    Set oFP = New clsFarmaPuntos
    Set oFPC = New clsFPConstante
    
    Tipo = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "INDFV460", objUsuario.CodigoEmpresa)
    
    If Tipo = "D" Then
        temp = Split(oFP.GetEstadoTarjeta(objVenta.NumDNI, objUsuario.Codigo), "@")
        vDni = objVenta.NumDNI
    ElseIf Tipo = "T" Then
        temp = Split(oFP.GetEstadoTarjeta(objVenta.NumTarjetaPuntos, objUsuario.Codigo), "@")
        vDni = temp(1)
    ElseIf Tipo = "A" Then
        If objVenta.NumDNI <> "" Then
            temp = Split(oFP.GetEstadoTarjeta(objVenta.NumDNI, objUsuario.Codigo), "@")
            vDni = objVenta.NumDNI
        ElseIf objVenta.NumTarjetaPuntos <> "" Then
            temp = Split(oFP.GetEstadoTarjeta(objVenta.NumTarjetaPuntos, objUsuario.Codigo), "@")
            vDni = temp(1)
        End If
    Else
        Err.Raise vbObjectError, "TransactionInit", "No tiene implementada esta opción (INDFV460 = " & Tipo & ")"
    End If
    
    If temp(0) = oFPC.EstadoTarjeta.INVALIDA Or temp(0) = oFPC.EstadoTarjeta.SIN_ESTADO Then
        Label3.Caption = 0
        Err.Raise vbObjectError, "TransactionInit", "Tarjeta No válida o sin conexión"
    End If
   
    If frmPedido.pstrDniCli = "" Or frmPedido.pstrDniCli <> vDni Then
        frmPedido_Busca_Cli.CboTipoDoc.BoundText = IIf(Len(vDni) = 8, "002", "004")
        frmPedido_Busca_Cli.ctlTxtDNI.Text = vDni
        frmPedido_Busca_Cli.b_monedero = True
        frmPedido_Busca_Cli.b_afiliar = True
        frmPedido_Busca_Cli.ctlTxtDNI_KeyPress (13)
    End If
    
    objVenta.PuntosTarjetaMonedero = Val(temp(2))
    Label3.Caption = Val(objVenta.PuntosTarjetaMonedero)
    
    Exit Sub
ControlError:
    Label3.Caption = 0
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub
