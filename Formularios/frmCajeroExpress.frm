VERSION 5.00
Begin VB.Form frmCajeroExpress 
   BorderStyle     =   0  'None
   Caption         =   "Cajero Express - Scotiabank Perú"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   Icon            =   "frmCajeroExpress.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameAviso 
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   0
      Width           =   3615
      Begin VB.Label lblAviso 
         Caption         =   "Esperando"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   4920
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar Turno"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Picture         =   "frmCajeroExpress.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdActivarPOS 
      Caption         =   "&Activar POS"
      Height          =   615
      Left            =   960
      Picture         =   "frmCajeroExpress.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   5280
      Picture         =   "frmCajeroExpress.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   2400
      Picture         =   "frmCajeroExpress.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingresar datos de la transacción"
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3855
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6135
         Begin vbp_Ventas.ctlDataCombo dbcTipoOperacion 
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   120
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Dolares"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   6
            Top             =   2640
            Width           =   1095
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Soles"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   5
            Top             =   2640
            Width           =   1095
         End
         Begin vbp_Ventas.ctlTextBox txtTipoCambio 
            Height          =   375
            Left            =   1800
            TabIndex        =   15
            Top             =   3000
            Width           =   3855
            _ExtentX        =   6800
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
         Begin vbp_Ventas.ctlTextBox txtImporte 
            Height          =   375
            Left            =   1800
            TabIndex        =   16
            Top             =   2160
            Width           =   3855
            _ExtentX        =   6800
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
         Begin vbp_Ventas.ctlTextBox txtNroAutoriza 
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Top             =   1620
            Width           =   3855
            _ExtentX        =   6800
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
         Begin vbp_Ventas.ctlTextBox txtNroTarjeta 
            Height          =   375
            Left            =   1800
            TabIndex        =   18
            Top             =   600
            Width           =   3855
            _ExtentX        =   6800
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
         Begin VB.Label lblNombreTarjeta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "XXX"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1800
            TabIndex        =   13
            Top             =   1140
            Width           =   3855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio :"
            Height          =   195
            Left            =   465
            TabIndex        =   12
            Top             =   3090
            Width           =   1200
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   990
            TabIndex        =   11
            Top             =   2670
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe :"
            Height          =   195
            Left            =   1050
            TabIndex        =   10
            Top             =   2250
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operación :"
            Height          =   195
            Left            =   480
            TabIndex        =   9
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Autorización :"
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   1710
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Tarjeta :"
            Height          =   195
            Left            =   735
            TabIndex        =   7
            Top             =   690
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmCajeroExpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Dim objTarjeta As clsCajeroExpress
Dim objMoneda As New clsMoneda
'''Dim objUsuario As New clsUsuario
Dim strCodTarjeta As String
Dim intValidaTarjeta As Integer
Dim intEntrar As Integer
Dim objCajeroExpress As New clsCajeroExpress
Dim objTarjeta As New clsFormaPago

'*****************
Dim lstrMensaje As String
Dim lintTimes As Long
Dim objArchivoTexto As clsArchivoTexto
Dim gvarError As String


Private Sub cmdActivarPOS_Click()
Dim LiResult As Integer
Dim lsTrama As String
Dim intValidaCajero As Integer
Dim objDocumento As clsDocumento
Dim strCodLiquidacion As String
    On Error GoTo CtrlErr
    Set objArchivoTexto = New clsArchivoTexto
    objArchivoTexto.WriteArchivo ("Nueva Transacción : " & "Usuario -> " & objUsuario.Codigo & " Fecha y Hora -> " & objUsuario.SYSDATE & " " & Time)
    Set objDocumento = New clsDocumento
    
    
    intValidaCajero = objDocumento.MaqEsCajeroCorresponsal(objUsuario.CodigoEmpresa, objUsuario.NombrePC)
    strCodLiquidacion = objDocumento.ValidaExisteLiquidacion(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objUsuario.CodigoLocal, objUsuario.Codigo)
    objArchivoTexto.WriteArchivo ("Obteniendo Código Liquidación -> " & strCodLiquidacion)
    
    
    Set objDocumento = Nothing
    
        
    If intValidaCajero = 0 Then MsgBox "Esta Máquina No tiene Registrado el Tipo de Documento VOUCHER para trabajar como Cajero Corresponsal" + Chr(13) + "Agregar el Tipo de documento en la tabla REL_DOCUMENTO_MAQUINA ", vbExclamation + vbOKOnly, App.ProductName: Exit Sub
    
    
    
    cmdActivarPOS.Enabled = False
    cmdSalir.Enabled = False
    cmdCerrar.Enabled = False
    lsTrama = ""
    avisoResalta "Esperando", 5, True
    'frm_SERV_Wait.lblMensaje.Caption = "Esperando Transaccion en POS..."
   ' frm_SERV_Wait.Show 1
    LiResult = objCajeroExpress.b_ActivaPOS("proceso=CajeroExpress", lsTrama)
    objArchivoTexto.WriteArchivo ("Trama del POS -> " & lsTrama & "LiResult ->" & LiResult)
    'Unload frm_SERV_Wait
    avisoResalta "Finalizado", 0, False
    If LiResult = 0 Then
        dbcTipoOperacion.BoundText = objCajeroExpress.TipoOperacion
        If objCajeroExpress.TipoOperacion <> "OTROS" Then
            txtImporte.Text = objCajeroExpress.Importe
            optMoneda(0).Value = IIf(objCajeroExpress.Moneda = 0, True, False)
            optMoneda(1).Value = IIf(objCajeroExpress.Moneda = 1, True, False)
            txtNroTarjeta.Text = objCajeroExpress.NumeroTarjeta
            txtNroAutoriza.Text = objCajeroExpress.NumeroTransaccion
            objArchivoTexto.WriteArchivo ("Entrando a Grabar ")
            cmdGrabar_Click
            objArchivoTexto.WriteArchivo ("Saliendo de Grabar -> " & gvarError)
        End If
    Else
        avisoResalta "Transaccion Invalida", 5, False
        MsgBox "Transacción Invalida " & Trim(Str(LiResult)) & Chr(13) '& MGR_sGetINI("C:\Archivos de programa\Cajero Express BTL\HCC.ini", "RETORNO", CStr(LiResult), ""), vbCritical, App.ProductName
    End If
    Call Nuevo
    GoTo Cerrar
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
    'Unload frm_SERV_Wait
    Call Nuevo
    objArchivoTexto.WriteArchivo ("Capturando Error -> " & Err.Description)
    GoTo Cerrar

Cerrar:
        
        objArchivoTexto.WriteArchivo ("Termina transaccion : " & objUsuario.SYSDATE & " " & Time)
        objArchivoTexto.WriteArchivo ("-----------------------------------------------------------")
        Set objArchivoTexto = Nothing


End Sub

Private Sub cmdCerrar_Click()
Dim objCajeroExpress As New clsCajeroExpress
Dim LiResult As Integer
    
    'frm_SERV_Wait.lblMensaje.Caption = "Realizando el Cierre del Cajero Express..."
    'frm_SERV_Wait.Show
    
    If MsgBox("Esta opción cierra el Turno del Cajero Corresponsal y emite los reportes." + Chr(13) + "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    cmdActivarPOS.Enabled = False
    cmdSalir.Enabled = False
    avisoResalta "Cerrando", 5, True
    LiResult = objCajeroExpress.b_CierraTurnoPOS("proceso=CierreCE")
        
    If LiResult = 0 Then
        MsgBox "Se realizo la transacción correctamente", vbInformation, App.ProductName
    Else

        'MsgBox MGR_sGetINI("C:\Archivos de programa\Cajero Express BTL\HCC.ini", "RETORNO", CStr(LiResult), ""), vbCritical, App.ProductName
        MsgBox "Transacción Invalida " & Trim(Str(LiResult)), vbCritical, App.ProductName
        
    End If
    avisoResalta "Esperando", 0, False
    Call Nuevo

End Sub

Private Sub cmdGrabar_Click()
Dim rsTipoOperacion As oraDynaset
Dim oTipDoc As String
Dim oNumDoc As String


    On Error GoTo CtrlErr
    avisoResalta "Registrando", 5, True
    If dbcTipoOperacion.Text = "" Then MsgBox "No se ha registrado el tipo de operación", vbInformation + vbOKOnly, "Cajero Express": Exit Sub
    If txtNroTarjeta.Text = "" Then MsgBox "No se ha registrado el Nro de Tarjeta", vbInformation + vbOKOnly, "Cajero Express": Exit Sub
    If txtNroAutoriza.Text = "" Then MsgBox "No se ha registrado el Nro de autorización", vbInformation + vbOKOnly, "Cajero Express": Exit Sub
    If txtImporte.Text = "" Then MsgBox "No se ha registrado el importe de operación", vbInformation + vbOKOnly, "Cajero Express": Exit Sub
    If Not optMoneda(0).Value And Not optMoneda(1).Value Then MsgBox "No se ha registrado la moneda de la operación", vbInformation + vbOKOnly, "Cajero Express": Exit Sub

                
                
    
    
    objCajeroExpress.Moneda = IIf(optMoneda(0).Value, 1, 2)
    objCajeroExpress.NumeroTarjeta = txtNroTarjeta.Text
    objCajeroExpress.NumeroTransaccion = txtNroAutoriza.Text
    objCajeroExpress.Importe = CDbl(txtImporte.Text)
    
    gvarError = objCajeroExpress.GrabaCajeroCorresponsal(oTipDoc, oNumDoc)
                         
    If gvarError = "" Then
        avisoResalta "Transaccion exitosa", 5, False
        MsgBox "Se realizo la transacción satisfactoriamente" + Chr(13) + oTipDoc + "-" + oNumDoc, vbExclamation + vbOKOnly, "Atención"
    Else
        avisoResalta "Transaccion Invalida 5, True"
        MsgBox gvarError, vbCritical + vbOKOnly, App.ProductName
    End If
    
    Exit Sub

CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbInformation, Err.Number



End Sub

Private Sub cmdSalir_Click()
    mdiPrincipal.picComandos.Enabled = True
    Unload Me
    'objVenta.CancelarVenta
End Sub







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    ''    psub_KeyDownAplicacion KeyCode, Shift
        Select Case KeyCode
            Case vbKeyEscape
               ''' MsgBox "Entre aqui"
                
        End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
    
    
End Sub



Private Sub Form_Load()

On Error GoTo handle
setteaFormulario Me


Set dbcTipoOperacion.RowSource = objCajeroExpress.ListaTipoOperacion(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_BANCO_CAJERO_CORRES"))
            
dbcTipoOperacion.BoundColumn = "COD_TIPO_OPE_BANCO"
dbcTipoOperacion.ListField = "DES_TIPO_OPERACION"
dbcTipoOperacion.ListField2 = "COD_TIPO_OPERACION"


intEntrar = 0
Set objMoneda = New clsMoneda
txtTipoCambio.Text = objMoneda.TCambio(objUsuario.TipoCambioDefault, objUsuario.TipoCambioMonedaDefault)
Call Nuevo
Frame2.Enabled = False
cmdGrabar.Enabled = False



mdiPrincipal.picComandos.Enabled = False


handle:


End Sub

Private Sub Timer1_Timer()
DoEvents
Static bytIntentos As Long
If bytIntentos <= lintTimes Then
    bytIntentos = bytIntentos + 1
    'lblAviso.ForeColor = RGB(aleatorioEntre(1, 255), aleatorioEntre(1, 255), aleatorioEntre(1, 255))
    lblAviso = lstrMensaje & ReplicateString(".", bytIntentos)
Else
    lblAviso = lstrMensaje
    bytIntentos = 0
End If
End Sub
Function ReplicateString(Source As String, times As Long) As String
    ReplicateString = Replace$(Space$(times), " ", Source)
End Function


Private Sub txtNroAutoriza_GotFocus()
   intEntrar = 0
End Sub
Private Sub txtNroTarjeta_KeyPress(KeyAscii As Integer)
Dim strNumeroTarjeta As String

On Error GoTo CtrlError
  strNumeroTarjeta = txtNroTarjeta.Text
  If KeyAscii = 13 Then
    If Len(txtNroTarjeta.Text) = 0 Then Exit Sub
    If intEntrar = 0 And InStr(strNumeroTarjeta, "&") > 0 Then
        intEntrar = intEntrar + 1
    Else
        If InStr(strNumeroTarjeta, "&") > 0 Then
            txtNroTarjeta.Text = Mid(strNumeroTarjeta, 3, InStr(strNumeroTarjeta, "&") - 3)
        Else
            txtNroTarjeta.Text = strNumeroTarjeta
        End If
        

        
        
        
        strCodTarjeta = objTarjeta.ValidaTarjeta(txtNroTarjeta.Text, "1")
        lblNombreTarjeta.Caption = objTarjeta.ValidaTarjeta(txtNroTarjeta.Text, "2")
        
        If lblNombreTarjeta.Caption = "*" Then
            lblNombreTarjeta.Caption = "Tarjeta no registrada"
            txtNroTarjeta.Text = ""
        Else
            intEntrar = 0
        End If
        
        SendKeys "{tab}"
    End If
  End If
  Exit Sub
CtrlError:
    MsgBox Err.Description, vbInformation + vbOKOnly, Err.Number
End Sub
Private Sub Nuevo()
    dbcTipoOperacion.BoundText = ""
    txtNroTarjeta.Text = ""
    lblNombreTarjeta.Caption = ""
    txtNroAutoriza.Text = ""
    txtImporte.Text = ""
    optMoneda(0).Value = False
    optMoneda(1).Value = False
    cmdActivarPOS.Enabled = True
    cmdSalir.Enabled = True
    cmdCerrar.Enabled = True
    'lblAviso.Caption = "Esperando"
    avisoResalta "Esperando", 0, False
End Sub
Sub avisoResalta(ByVal strMensaje As String, Optional intTimes As Long = 3, Optional ByVal blnResalta As Boolean = True)
lstrMensaje = strMensaje
lintTimes = intTimes
Select Case blnResalta
    Case True
        Timer1.Interval = 250
    Case False
        Timer1.Interval = 0
        lblAviso = strMensaje
End Select
End Sub
Function aleatorioEntre(ByVal numMinimo As Long, ByVal numMaximo As Long) As Long
aleatorioEntre = Int((numMaximo - numMinimo + 1) * Rnd + numMinimo)
End Function



