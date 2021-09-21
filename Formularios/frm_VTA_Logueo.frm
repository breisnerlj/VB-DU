VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_VTA_Logueo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frm_VTA_Logueo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRecordar 
      Caption         =   "Recordar contraseña"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   810
      TabIndex        =   8
      Top             =   1680
      Width           =   2130
   End
   Begin MSComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1935
      Width           =   3750
      _ExtentX        =   6615
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
   Begin vbp_Ventas.ctlTextBox txtUsuario 
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
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
   Begin vbp_Ventas.ctlTextBox txtPassword 
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
      PasswordChar    =   "*"
      TABAuto         =   0   'False
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
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   60
      X2              =   3660
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "BIENVENIDO AL SISTEMA DE VENTAS"
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
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su password, recuerde es secreto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su usuario (codigo de Planilla)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   1980
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password :"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuario :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Width           =   945
   End
End
Attribute VB_Name = "frm_VTA_Logueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oraUsuario As oraDynaset
'Dim flgConecto As Boolean

Private Sub cmdRecordar_Click()
   

   On Error GoTo Control

    If Len(Trim(txtUsuario.Text)) < 5 Or Len(Trim(txtUsuario.Text)) >= 6 Then MsgBox "Verifique o Ingrese bien su Código de Usuario", vbInformation, App.ProductName: txtUsuario.SetFocus: Exit Sub
    frm_VTA_RecordarContraseña.Mostrar txtUsuario.Text

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objArchivoIni As cls_ArchivoIni

    On Error GoTo Handle
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF1
        Case vbKeyF2
        Case vbKeyF3
        Case vbKeyF4
        Case vbKeyReturn
            If txtUsuario.Text = "" Then Screen.MousePointer = vbDefault: Exit Sub
            If txtPassword.Text = "" Then Screen.MousePointer = vbDefault: Exit Sub
            
            If objOFFVenta.Existe(gstrIni) <> 0 Then
                Set objArchivoIni = New cls_ArchivoIni
                If txtUsuario.Text = objArchivoIni.LeerIni(gstrIni, "general", "ADMIN_USER") And txtPassword.Text = objArchivoIni.LeerIni(gstrIni, "general", "ADMIN_PASS") Then
                    OFF_Main
                    Unload Me
                    Exit Sub
                End If
                Set objArchivoIni = Nothing
            End If
    
            If MainX(txtUsuario.Text, txtPassword.Text) = False Then Exit Sub
            
            Set oraUsuario = objUsuario.Login(txtUsuario.Text, txtPassword.Text)

            If objUsuario.Conectado Then
'                flgConecto = True
                MsgBox "BIENVENIDO(A) " & Chr(13) & _
                            "    Usuario: " & objUsuario.Nombre & Chr(13) & _
                            "    Local  : " & objUsuario.NombreLocal & " BTL " & objUsuario.CodigoLocal & Chr(13) & _
                            "    Maquina: " & objUsuario.NombrePC & Chr(13), vbInformation, App.ProductName
''''               mdiPrincipal.Caption = gstrAplicacion & " * Ver: " & gstrVersion & " [" & gvarUSUARIO & "@" & gvarTNSNAME & "] " & "Local " & objUsuario.CodigoLocal & " >" & objUsuario.Empresa & "<"
''''               mdiPrincipal.HabilitaPermisos
''''               frmPedido.FormDragger1.Caption = objUsuario.Nombre


 
                gstrCodUsuario = objUsuario.Codigo
                gstrCodAreaUsuario = objUsuario.CodigoLocal
                gstrDesAreaUsuario = objUsuario.CentroCosto
                gstrPassword = objUsuario.Password
                blnEnviaMensajeDelivery = IIf(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ENVIA_MENSAJE_DELIVERY") = "1", True, False)
                '' settea la fecha
                
                Call SettearHora(Format(objUsuario.sysdate, "YYYY"), Format(objUsuario.sysdate, "MM"), Format(objUsuario.sysdate, "DD"))
                'Crear valores de configuracion por defecto del sistema de contingencia
                Call CreateContingenciaIniFile
                
                'Crear archivo de configuracion del cliente ADO
                Call CreateSchemaIniFile
                
                'Obtener el archivo de actualizacion automatica
                Call CopiaActualizar
                
                'Para delivery de provincias
                If objUsuario.flgDeliveryProv = "1" Then
                    Call CargaDatosClienteProv
                    frm_VTA_TipoMaquina.Show vbModal
                End If
    
                If objUsuario.Conectado Then
                    tipoPantalla
                End If
                Unload Me
            Else
                gclsOracle.Cerrar
            End If
                                        
    End Select
    Set objArchivoIni = Nothing
    Exit Sub
    
Handle:
    Select Case Err.Number
    Case -2147201500
        frm_VTA_CambioContraseña.Mostrar txtUsuario.Text, txtPassword.Text
        txtPassword.SetFocus
    Case Else
        
        MsgBox Err.Description, vbCritical, App.ProductName
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub CreateSchemaIniFile()
'    Dim strIniFile As String
'
'    On Error GoTo ErrorHandler
'
'    strIniFile = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "Schema.ini"
'    If Len(Dir$(strIniFile, vbHidden)) <> 0 Then
'        Call Kill(strIniFile)
'    End If
'
'    Open strIniFile For Binary As #1
'        Put #1, , CStr("[usuarios.txt]") & vbCrLf
'        Put #1, , CStr("Format=Delimited(;)") & vbCrLf
'        Put #1, , vbCrLf
'        Put #1, , CStr("[precios.txt]") & vbCrLf
'        Put #1, , CStr("Format=Delimited(;)") & vbCrLf
'        Put #1, , vbCrLf
'        Put #1, , CStr("[detalleventa.txt]") & vbCrLf
'        Put #1, , CStr("Format=Delimited(;)") & vbCrLf
'        Put #1, , vbCrLf
'        Put #1, , CStr("[pagoventa.txt]") & vbCrLf
'        Put #1, , CStr("Format=Delimited(;)") & vbCrLf
'    Close #1
'    Exit Sub
'
'ErrorHandler:
'    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub CreateContingenciaIniFile()
    Dim objArchivoIni As cls_ArchivoIni

    On Error GoTo ErrorHandler
    Set objArchivoIni = New cls_ArchivoIni
    
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "FLG_CONTINGENCIA", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FLG_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "ADMIN_USER", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_USER_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "ADMIN_PASS", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PASS_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "SEC_VENTA", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_SEC_VENTA_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "CIA", objUsuario.CodigoEmpresa)
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "LOCAL", objUsuario.CodigoLocal)
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "EMPRESA", objUsuario.Empresa)
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "RUC", objUsuario.Ruc)
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "DIRECCION", objUsuario.direccion)
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "DIRECCION_LOCAL", objUsuario.DireccionLocal)
    Call objArchivoIni.GuardarIni(gstrIni, "GLOSA", "LINE1", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LINE1_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GLOSA", "LINE2", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LINE2_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GLOSA", "LINE3", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LINE3_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GLOSA", "LINE4", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LINE4_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "GLOSA", "LINE5", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LINE5_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "MONEDA", "COD_MONEDA", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_MONEDA_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "MONEDA", "DES_MONEDA", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_DES_MONEDA_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "MONEDA", "SMB_MONEDA", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_SMB_MONEDA_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "IMPRESION", "COD_FORMATO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_FORMATO_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "IMPRESION", "DES_FORMATO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_DES_FORMATO_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "IMPRESION", "CTD_ANCHO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_CTD_ANCHO_CONTINGENCIA"))
    Call objArchivoIni.GuardarIni(gstrIni, "IMPRESION", "CTD_ALTO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_CTD_ALTO_CONTINGENCIA"))
                
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "COD_DOCUMENTO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_DOCUMENTO_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "DES_DOCUMENTO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_DES_DOCUMENTO_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "SER_DOCUMENTO", gclsOracle.FN_Valor("BTLPROD.PKG_CONTINGENCIA.FN_LISTA_SERIE_X_MAQUINA", objUsuario.NombrePC, "2"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "NUM_DOCUMENTO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_NUM_DOCUMENTO_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "NUM_LIN_DOCUMENTO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_NUM_LIN_DOCUMENTO_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "ANCHO_DOCUMENTO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ANCHO_DOCUMENTO_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "COD_FORMAPAGO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_FORMAPAGO_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "DES_FORMAPAGO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_DES_FORMAPAGO_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "COD_DOC_DEFAULT", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_DOC_DEFAULT_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "ULT_DOC_EMITIDO", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ULT_DOC_EMITIDO_CONTING"))
    'Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "COD_SERIE_ETIQ", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_SERIE_ETIQ_CONTING"))
    Call objArchivoIni.GuardarIni(gstrIni, "GENERAL", "COD_SERIE_ETIQ", gclsOracle.FN_Valor("BTLPROD.PKG_CONTINGENCIA.FN_SERIE_TICKET", objUsuario.NombrePC))
                
    Call objArchivoIni.GuardarIni(gstrIni, "UPDATEOPTIONS", "TIMECHECK", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIMECHECK"))
    Call objArchivoIni.GuardarIni(gstrIni, "UPDATEOPTIONS", "LASTUPDATE", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LASTUPDATE"))
    Call objArchivoIni.GuardarIni(gstrIni, "UPDATEOPTIONS", "NEXTUPDATE", gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_NEXTUPDATE"))
                
    Set objArchivoIni = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If flgConecto = True Then gclsOracle.Cerrar
            'End
End Sub



Private Sub txtUsuario_Change()
cmdRecordar.Enabled = Trim(txtUsuario.Text) <> ""
End Sub

Private Sub txtUsuario_GotFocus()
    stbPrincipal.Panels(1).Text = Label4.Caption
End Sub

Private Sub txtPassword_GotFocus()
    stbPrincipal.Panels(1).Text = Label5.Caption
End Sub

Private Sub CopiaActualizar()
On Error GoTo Handle
Dim strNombreActualizador As String
Dim strRutaCopia     As String
strNombreActualizador = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ACTUALIZADOR_COPIA")
strNombreActualizador = IIf(strNombreActualizador = "", strNombreActualizador, "BTLOffLine.exe")
strRutaCopia = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_RUTA_SW")

Select Case gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ACTUALIZADOR_COPIA")
    Case 0  'no hace nada, no lop actuliza ni nada
        
    Case 1
        If Dir(App.Path & IIf(right(App.Path, 1) = "\", "", "\") & strNombreActualizador) = "" Then
            FileCopy strRutaCopia & strNombreActualizador, App.Path & IIf(right(App.Path, 1) = "\", "", "\") & strNombreActualizador
        End If
    Case 2
        KillProcess (strNombreActualizador)
        FileCopy strRutaCopia & strNombreActualizador, App.Path & IIf(right(App.Path, 1) = "\", "", "\") & strNombreActualizador
End Select

'ACA VEO SI EJECUTA O NO EJECUTA
Select Case gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ACTUALIZADOR_EJECUTA")
    Case 0  'no ejecuta nada
        
    Case 1
        KillProcess (strNombreActualizador)
        Shell App.Path & IIf(right(App.Path, 1) = "\", "", "\") & strNombreActualizador, vbMinimizedNoFocus
    Case 2
        KillProcess (strNombreActualizador)
        Shell App.Path & IIf(right(App.Path, 1) = "\", "", "\") & strNombreActualizador, vbNormalNoFocus
    Case 3
        KillProcess (strNombreActualizador)
        Shell App.Path & IIf(right(App.Path, 1) = "\", "", "\") & strNombreActualizador, vbMinimizedFocus
        
    Case 4
        KillProcess (strNombreActualizador)
        Shell App.Path & IIf(right(App.Path, 1) = "\", "", "\") & strNombreActualizador, vbNormalFocus
End Select

Handle:

End Sub
Public Function CargaMetas()

End Function
