VERSION 5.00
Begin VB.Form frmPedido_Busca_Cli 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar Cliente"
   ClientHeight    =   1605
   ClientLeft      =   2925
   ClientTop       =   4185
   ClientWidth     =   4380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_Editar 
      Caption         =   "Editar"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin vbp_Ventas.ctlDataCombo CboTipoDoc 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   300
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox ctlTxtDNI 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Tipo            =   3
      MaxLength       =   15
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
   Begin VB.Label Label3 
      Caption         =   "<F1>:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Documento:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. Documento:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmPedido_Busca_Cli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim b_Encontrado As Boolean
Dim objTipoDoc As New clsTipoDocumento
Dim objClienteD As New clsClienteD
Dim v_antiguo As String
Dim v_Saldo As String
Dim v_BoolSaldo As Boolean
Public v_delfrm As String
Public b_monedero As Boolean
Public b_afiliar As Boolean

Private Sub Form_Activate()
    If Not frmPedido.lblCtrl.Visible Then
        Unload Me
    Else
        ctlTxtDNI.Focus
    End If
End Sub

Private Sub Form_Load()
    Dim v_Bool As Boolean

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    ctlTxtDNI.Text = ""
    CargarTipoDocumentos
    CboTipoDoc.Text = "002-DNI"
    ctlTxtDNI.MaxLength = CboTipoDoc.BoundText2
    b_monedero = True
    b_afiliar = False

    'visualiza el flag de edicion
    'Dim rs As oraDynaset
    v_Bool = objVenta.MuestraFidelizado("FLGEDICION")
    chk_Editar.Visible = v_Bool
    Label3.Visible = v_Bool
    chk_Editar.Value = 0

    'visualiza el saldo del cliente
    v_BoolSaldo = objVenta.MuestraFidelizado("FLGSALDO2")

End Sub

Public Sub ctlTxtDNI_KeyPress(KeyAscii As Integer)
    Dim ErrorMessage As String
    If KeyAscii = 13 Then
        Dim EliminarVales As Boolean
        If Trim(ctlTxtDNI.Text) = "" Then
            'MsgBox " Ingresar el Nro de Documento ", vbInformation, "Aviso"
            frmPedido.lbl_Cliente.Caption = ""
            objVenta.CodigoCliente = ""
            frmPedido.Cal_Promo
            frmPedido.Cal_Montos
            frmPedido.grdPedido.Refresh
            frmPedido.loadOptions
            Unload Me
            'Exit Sub
        Else
            'Valida si existe forma de pago VALE FID, de ser el caso restringe el cambio de DNI
            Dim eliminoFila As Boolean
            frm_VTA_FormaPago.GrdListaFP.MoveFirst
            While Not frm_VTA_FormaPago.GrdListaFP.EOF
                eliminoFila = False
                Debug.Print frm_VTA_FormaPago.GrdListaFP.Columns(0)
                If frm_VTA_FormaPago.GrdListaFP.Columns(0) = "011" And EliminarVales = False Then
                    If MsgBox("Al cambiar de DNI se eliminaran los vales ingresados, desea cambiar de DNI?", vbQuestion + vbYesNo) = vbYes Then
                        EliminarVales = True
                    Else
                        GoTo Salir
                    End If
                End If
                If frm_VTA_FormaPago.GrdListaFP.Columns(0) = "011" And EliminarVales = True Then
                    'frm_VTA_FormaPago.GrdListaFP.Row
                    frm_VTA_FormaPago.GrdListaFP.Delete
                    eliminoFila = True
                End If
                If eliminoFila = False Then
                    frm_VTA_FormaPago.GrdListaFP.MoveNext
                End If
            Wend
            
            If Len(ctlTxtDNI.Text) < Val(CboTipoDoc.BoundText2) And Mid(CboTipoDoc.Text, 1, 3) <> "003" Then
                MsgBox " El Nro de Documento no tiene completa la cantidad de dígitos", vbInformation, "Aviso"
                Exit Sub
            End If
            
            If objVenta.ParametroValor("FLGFIDDLV") = "1" Then
            'gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGFIDDLV") = 1 Then
                '**Proceso Original**'
                'buscar cliente con nro documento
                b_Encontrado = False
                
                gstrCodTarjetaMon = IIf(gstrCodTarjetaMon <> frm_VTA_Busqueda.txtBuscar.Text, "", gstrCodTarjetaMon)
                
                CargarClienteMonedero
    '            If Not b_Encontrado And Not b_monedero Then CargarCliente
                
                'Or chk_Editar.Value = 0
                If Not b_Encontrado And b_monedero Then
                    frmPedido.lbl_Cliente.Caption = ""
                    'FrmPedido_Ingre_Cli.Cbo_Tipo_Doc.Text = Mid(CboTipoDoc.Text, 1, 3)
                    'FrmPedido_Ingre_Cli.ctlTxtDNI.Text = ctlTxtDNI.Text
                    FrmPedido_Ingre_Cli.Caption = "Datos del cliente - " & _
                                                  IIf(b_monedero, "Programa Monedero MIFARMA", "Fidelizado")
                    FrmPedido_Ingre_Cli.b_monedero = b_monedero
                    FrmPedido_Ingre_Cli.carga Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text
                    
                    objVenta.CodigoCliente = "" & v_delfrm
                    If frmPedido.lbl_Cliente.Caption <> "" Then
                        frmPedido.pstrDniCli = ctlTxtDNI.Text
                        frmPedido.pstrNomcli = frmPedido.lbl_Cliente.Caption
                        MsgBox "Bienvenido " & frmPedido.lbl_Cliente.Caption & vbCrLf & _
                               "DNI/CE: " & frmPedido.pstrDniCli & vbCrLf & _
                               "Puntos: " & CStr(Val(objVenta.PuntosTarjetaMonedero)), vbInformation + vbOKOnly, _
                               "Programa Monedero del Ahorro"
                        frmPedido.Cal_Promo
                        frmPedido.Cal_Montos
                        frmPedido.grdPedido.Refresh
                        frmPedido.loadOptions
                        Unload Me
                    End If
    '''            Else
    '''                If Not b_monedero Then
    '''                    'busca saldo
    '''                    If v_BoolSaldo Then
    '''                        v_Saldo = vbCrLf & Space(21) & "(Saldo S/. " & Format(objVenta.Saldo_Cliente(objVenta.CodigoCliente), "##0.00") & ")"
    '''                    Else
    '''                        v_Saldo = ""
    '''                    End If
    '''
    '''                    MsgBox "Bienvenido, " & frmPedido.lbl_Cliente.Caption & "." & v_Saldo, vbInformation, "Actualización"
    '''    '                If objVenta.CodigoCliente <> v_antiguo Then
    '''                        objVenta.CodigoCliente = v_antiguo
    '''                        frmPedido.Cal_Promo
    '''                        frmPedido.Cal_Montos
    '''                        frmPedido.grdPedido.Refresh
    '''                        frmPedido.loadOptions
    '''                        'frmPedido.optEfectivo.SetFocus
    '''                        'frmPedido.grdPedido.MoveFirst
    '''     '               End If
    '''                End If
    '''                Unload Me
                Else
                    objVenta.CodigoCliente = "" & v_delfrm
                    If frmPedido.lbl_Cliente.Caption <> "" Then
                        MsgBox "Bienvenido " & frmPedido.lbl_Cliente.Caption & vbCrLf & _
                               "DNI/CE: " & frmPedido.pstrDniCli & vbCrLf & _
                               "Puntos: " & CStr(Val(objVenta.PuntosTarjetaMonedero)), vbInformation + vbOKOnly, _
                               "Programa Monedero del Ahorro"
                        frmPedido.Cal_Promo
                        frmPedido.Cal_Montos
                        frmPedido.grdPedido.Refresh
                        frmPedido.loadOptions
                        Unload Me
                    End If
                End If
                '**Fin Proceso Original**'
            Else
                '**Proceso Nuevo**'
                'Validar documento por marca
                Dim oFP As New clsFarmaPuntos
                Dim oAff As clsAfiliado
                Dim vNroCuenta As String
                Dim b_existe As Boolean
                Dim TipoDocumento As String
                Dim codClienteDNI As String
                If objUsuario.CodLocalCallCenter = "0DLV" Then 'MIFARMA
                    If Mid(CboTipoDoc.Text, 1, 3) = "002" Or Mid(CboTipoDoc.Text, 1, 3) = "004" Then
                        ' Buscar Cliente en Orbis
                        TipoDocumento = Mid(CboTipoDoc.Text, 1, 3)
                        vNroCuenta = IIf(TipoDocumento = "002", "D0", "E") & ctlTxtDNI.Text
                        Set oAff = oFP.ObtenerDatosAfiliadoSinTarjeta(vNroCuenta, objUsuario.Codigo)
                        If Not oAff Is Nothing Then b_existe = True
'                        If oAff Is Nothing Then
'                            Set oAff = objVenta.fnListaAfiliadoMonederoOff(Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text)
'                            b_existe = IIf(oAff Is Nothing, False, True)
'                        Else
'                            b_existe = oAff.DNI <> ""
'                        End If
                    End If
                ElseIf objUsuario.CodLocalCallCenter = "1DLV" Then 'INKAFARMA
                    Select Case Mid(CboTipoDoc.Text, 1, 3)
                        Case "002" 'DNI
                            TipoDocumento = "01"
                        Case "004" 'CARNET EXTRANJERIA
                            TipoDocumento = "02"
                        Case "003" 'PASAPORTE
                            TipoDocumento = "03"
                        Case "001" 'RUC
                            TipoDocumento = "04"
                    End Select
                    'tipoDocumento = Mid(CboTipoDoc.Text, 2, 2)
                    Set oAff = oFP.ObtenerAfiliadoInka(ctlTxtDNI.Text, TipoDocumento)
                    If Not oAff Is Nothing Then b_existe = True
                    'Si cliente es fidelizado entonces obtener todos sus vales sin mostrarlo en pantalla
                    
                End If
                'frmPedido_Busca_Cli.v_delfrm = vCodCliente
                'ctlCliente1.Codigo
                If b_existe Then
                    frmPedido.lbl_Cliente.Caption = oAff.Nombre & ", " & oAff.ApParterno & " " & oAff.ApMarterno
                    frmPedido.pstrDniCli = ctlTxtDNI.Text
                    frmPedido.pstrCodCliente_Ink = oAff.CodCliente
                    frmPedido.pstrPuntos_Ink = oAff.puntosDisponibles
                    frmPedido.lblPuntosAcum = IIf(frmPedido.pstrPuntos_Ink = "", 0, frmPedido.pstrPuntos_Ink)
                    objVenta.strTipoDocumento_Pts = TipoDocumento 'Mid(CboTipoDoc.Text, 1, 3)
                    objVenta.strNumDocumento_Pts = ctlTxtDNI.Text
                    MsgBox "DOCUMENTO: " & ctlTxtDNI.Text & Chr(13) & "ESTADO: AFILIADO"
                    'I.ECASTILLO 07.10.2020 - se crea nueva funcion para obtener el dato en base a tipo y numero documento
                    '                       - ya que tanto ORBIS como INKACLUB no retornan el dato esperado.
                    'objVenta.CodigoCliente = oAff.CodCliente
                    codClienteDNI = "" & gclsOracle.FN_Valor("BTLPROD.PKG_CLIENTE.FN_GET_COD_CLI", Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text, oAff.Nombre, oAff.ApParterno, oAff.ApMarterno, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.Codigo)
'                    If Len(Trim(codClienteDNI)) = }0 Then errorMessage = "Error al obtener codigo cliente fidelizado, documento " & ctlTxtDNI.Text & " no existe en la base.": GoTo Salir
                    objVenta.CodigoCliente = codClienteDNI
                    'F.ECASTILLO 07.10.2020
                Else
                    frmPedido.lbl_Cliente.Caption = ""
                    frmPedido.pstrDniCli = ""
                    frmPedido.pstrCodCliente_Ink = ""
                    frmPedido.pstrPuntos_Ink = ""
                    frmPedido.lblPuntosAcum = "0"
                    objVenta.strTipoDocumento_Pts = ""
                    objVenta.strNumDocumento_Pts = ""
                    objVenta.CodigoCliente = mdiPrincipal.ctlCliente1.Codigo 'ECASTILLO 07.10.2020
                    MsgBox "DOCUMENTO: " & ctlTxtDNI.Text & Chr(13) & "ESTADO: SIN AFILIAR"
                End If
                
                Unload Me
                '**Fin Proceso Nuevo**'
            End If
            
            If EliminarVales = True Then
                frm_VTA_FormaPago.GrdListaFP.Rebind
                
                frmPedido.Cal_Promo
                frmPedido.Cal_Montos
            End If
        End If
        
        If objVenta.ParametroValor("ACT_PCTCOM") = "1" Then
           If Len(Trim(objVenta.CodigoCliente)) > 0 Then
              ' busca si el DNI existe en RENIEC para indicar si va
              ' dar comision al Vendedor
              objVenta.vExisteDNI_RENIEC = objVenta.getExisteDNI_RENIEC(objVenta.CodigoCliente)
               If objVenta.vExisteDNI_RENIEC = "N" Then
                  frmPedido.lblDniInvalido.Visible = True
               Else
                  frmPedido.lblDniInvalido.Visible = False
               End If
            End If
        End If
        
Salir:
    Unload Me
    End If
    If Len(ErrorMessage) > 0 Then
        MsgBox ErrorMessage, vbCritical + vbOKOnly, App.ProductName
    End If
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub CargarTipoDocumentos()

  Set CboTipoDoc.RowSource = objTipoDoc.Lista
   
  CboTipoDoc.ListField = "DES_ABREVIATURA"
  CboTipoDoc.BoundColumn = "COD_DOCUMENTO_IDENTIDAD"
  CboTipoDoc.ListField2 = "NUM_DIGITOS"
    
End Sub

'''Private Sub CargarCliente()
'''    Dim vSex As String
'''    Dim rs As oraDynaset
'''
'''    Set rs = objClienteD.ListaCliente(Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text, gstrCodTarjetaFid)
'''
'''    'objVenta.CodigoCliente = "" & rs("COD_CLIENTE").Value
'''
'''    b_Encontrado = False
'''
'''    If rs.RecordCount > 0 Then
'''        frmPedido.lbl_Cliente.Caption = IIf(IsNull(rs("DES_NOM_CLIENTE").Value), "", rs("DES_NOM_CLIENTE").Value)
'''        frmPedido.lbl_Cliente.Caption = frmPedido.lbl_Cliente.Caption & ", " & IIf(IsNull(rs("DES_APE_CLIENTE").Value), "", rs("DES_APE_CLIENTE").Value)
'''        frmPedido.lbl_Cliente.Caption = frmPedido.lbl_Cliente.Caption & " " & IIf(IsNull(rs("DES_APE2_CLIENTE").Value), "", rs("DES_APE2_CLIENTE").Value)
'''
'''        If Not IsNull(rs("COD_CLIENTE").Value) Then
'''            'If (Not IsNull(rs("DES_NOM_CLIENTE").Value) And _
'''            '    Not IsNull(rs("DES_APE_CLIENTE").Value) And _
'''            '    Not IsNull(rs("FCH_NACIMIENTO").Value) And _
'''            '    Not IsNull(rs("FLG_SEXO").Value) And _
'''            '    rs("FLG_ESTADO").Value = "1") And _
'''            '    frmPedido_Busca_Cli.chk_Editar.Value = 0 _
'''            '    Then
'''                objVenta.CodigoCliente = "" & rs("COD_CLIENTE").Value
'''                b_Encontrado = True
'''                v_antiguo = "" & rs("COD_CLIENTE").Value
'''                frmPedido.pstrNomcli = rs("DES_NOM_CLIENTE") & " " & rs("DES_APE_CLIENTE") & " " & rs("DES_APE2_CLIENTE")
'''                frmPedido.pstrDniCli = rs("NUM_DOCUMENTO_ID")
'''                Unload Me
'''            'End If
'''        Else
'''            FrmPedido_Ingre_Cli.Cbo_Tipo_Doc.Text = Mid(CboTipoDoc.Text, 1, 3)
'''            FrmPedido_Ingre_Cli.ctlTxtDNI.Text = ctlTxtDNI.Text
'''            FrmPedido_Ingre_Cli.b_monedero = b_monedero
'''            FrmPedido_Ingre_Cli.CargarCliente
'''            FrmPedido_Ingre_Cli.cmdAceptar_Click
'''            'FrmPedido_Ingre_Cli.Grabar_Cliente
'''            b_Encontrado = True
'''            Unload Me
'''        End If
'''    Else
'''        b_Encontrado = False
'''        objVenta.CodigoCliente = ""
'''    End If
'''
'''    Set rs = Nothing
'''End Sub

' ******* MONEDERO MIFARMA *******
Public Sub CargarClienteMonedero()
    Dim oFP As New clsFarmaPuntos
    Dim oFPC As New clsFPConstante
    Dim oAff As clsAfiliado
    Dim vNroTarjeta As String, vNroCuenta As String, vEstadoTarjeta() As String
    Dim b_existe As Boolean, b_tiene_tarjetas As Boolean, vDniTemp As String
    Dim vUbigeo As String, vTelefono As String, vCelular As String, vCodCliente As String
    Dim oDP As New clsDocumentoPago
    Dim rs As oraDynaset
    
    On Error GoTo CtrlErr
    
    If Mid(CboTipoDoc.Text, 1, 3) = "002" Or Mid(CboTipoDoc.Text, 1, 3) = "004" Then
        ' Buscar Cliente en Orbis
        vNroCuenta = IIf(Mid(CboTipoDoc.Text, 1, 3) = "002", "D0", "E") & ctlTxtDNI.Text
        Set oAff = oFP.ObtenerDatosAfiliadoSinTarjeta(vNroCuenta, objUsuario.Codigo)
        If oAff Is Nothing Then
            Set oAff = objVenta.fnListaAfiliadoMonederoOff(Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text)
            b_existe = IIf(oAff Is Nothing, False, True)
        Else
            b_existe = oAff.DNI <> ""
        End If
        
        'Si no existe en orbis
        If Not b_existe Then
            If b_afiliar Then GoTo FlujoAfiliacion2
FlujoAfiliacion:

            'preguntar si desea afiliarse
            If MsgBox("Desea afiliarse al programa MONEDERO MIFARMA?", vbQuestion + vbYesNo) = vbYes Then
                If Not b_monedero Then b_monedero = True
                If Not b_afiliar Then b_afiliar = True
FlujoAfiliacion2:
                'solicitar tarjeta
                If gstrCodTarjetaMon = "" Then
                    vNroTarjeta = FrmPedido_Ingre_trj.ObtenerTarjeta("Afiliación nuevo cliente")
                    If vNroTarjeta = "" Or oDP.buscaFun(vNroTarjeta) <> "MONEDERO" Then
                        MsgBox "Tarjeta NO válida, NO se afiliará al programa MONEDERO MIFARMA", vbCritical + vbOKOnly, App.ProductName
                        b_Encontrado = False
                        b_monedero = False
                        objVenta.CodigoCliente = ""
                        Exit Sub
                    End If
                    gstrCodTarjetaMon = vNroTarjeta
                Else
                    vNroTarjeta = gstrCodTarjetaMon
                End If

ActualizarDatos:
                'validar tarjeta en orbis
                vEstadoTarjeta = Split(oFP.GetEstadoTarjeta(vNroTarjeta, objUsuario.Codigo), "@")
                If vEstadoTarjeta(0) = oFPC.EstadoTarjeta.SIN_ESTADO Then
                    vEstadoTarjeta = Split(objVenta.fnGetEstadoTarjetaMonedero(vNroTarjeta), "@")
                End If

                If vEstadoTarjeta(0) = oFPC.EstadoTarjeta.BLOQUEADA Then
                    MsgBox "Tarjeta bloqueada, NO se afiliará al programa MONEDERO MIFARMA", vbCritical + vbOKOnly, App.ProductName
                    b_Encontrado = False
                    b_monedero = False
                    objVenta.CodigoCliente = ""
                    'GoTo ClienteFidelizado
                    Exit Sub
                End If

                If vEstadoTarjeta(0) = oFPC.EstadoTarjeta.INVALIDA Then
                    MsgBox "Tarjeta NO válida, NO se afiliará al programa MONEDERO MIFARMA", vbCritical + vbOKOnly, App.ProductName
                    b_Encontrado = False
                    b_monedero = False
                    objVenta.CodigoCliente = ""
                    'GoTo ClienteFidelizado
                    Exit Sub
                End If

                'Tarjeta esta activa
                If vEstadoTarjeta(0) = oFPC.EstadoTarjeta.ACTIVA Or _
                   vEstadoTarjeta(0) = oFPC.EstadoTarjeta.BLOQUEADA_REDIMIR Then
                    If vEstadoTarjeta(1) <> ctlTxtDNI.Text Then
                        MsgBox "La tarjeta ingresada, esta asociada a otro DNI/CE", vbCritical + vbOKOnly, App.ProductName
                        b_Encontrado = False
                        b_monedero = False
                        objVenta.CodigoCliente = ""
                        Exit Sub
                    End If
                    If Not b_existe Then
                        Set oAff = oFP.ObtenerDatosAfiliadoSinTarjeta(vEstadoTarjeta(1), objUsuario.Codigo)
                        If oAff Is Nothing Then
                            Set oAff = objVenta.fnListaAfiliadoMonederoOff(Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text)
                        End If
                    End If

                    vUbigeo = oAff.Departamento & oAff.Provincia & oAff.Distrito
                    vUbigeo = IIf(Len(vUbigeo) < 6, "", vUbigeo)
                    vTelefono = IIf(IsNumeric(oAff.Telefono), oAff.Telefono, "")
                    vCelular = IIf(IsNumeric(oAff.Celular), oAff.Celular, "")

                    Set rs = objVenta.fnListaAfiliadoMonedero(oAff.TipoDni, oAff.DNI)
                    vCodCliente = "" & rs("COD_CLIENTE")
                    vCodCliente = objVenta.fnGrabaAfiliadoMonedero(vCodCliente, _
                                                                   oAff.TipoDni, _
                                                                   oAff.DNI, _
                                                                   oAff.Nombre, _
                                                                   oAff.ApParterno, _
                                                                   oAff.ApMarterno, _
                                                                   oAff.email, _
                                                                   IIf(oAff.Genero = "M", "1", "0"), _
                                                                   oAff.FechaNacimiento, _
                                                                   vUbigeo, _
                                                                   vNroTarjeta, _
                                                                   vTelefono, _
                                                                   vCelular, _
                                                                   oAff.TipoLugar, _
                                                                   oAff.Direccion, _
                                                                   oAff.TipoDireccion, _
                                                                   oAff.Referencias, _
                                                                   "N", _
                                                                   "S")
                    objVenta.CodigoCliente = vCodCliente
                    objVenta.EsVentaMonedero = True
                    objVenta.NumeroTarjetaMonedero = vNroTarjeta
                    objVenta.PuntosTarjetaMonedero = CDbl(Val(vEstadoTarjeta(2)))
                    frmPedido_Busca_Cli.v_delfrm = vCodCliente
                    frmPedido.lbl_Cliente.Caption = oAff.Nombre & ", " & oAff.ApParterno & " " & oAff.ApMarterno
                    frmPedido.pstrDniCli = oAff.DNI
                    frmPedido.pstrNomcli = frmPedido.lbl_Cliente.Caption
                    frmPedido.loadOptions
                    b_Encontrado = True
                    Exit Sub

                'Tarjeta no esta activa
                Else
                    'Buscar en local
                    If b_tiene_tarjetas And ctlTxtDNI.Text <> vNroTarjeta Then
                        vDniTemp = FrmPedido_Ingre_trj.ObtenerTarjeta("Registro de Tarjeta Adicional", eeDniCliente, True)
                        If vDniTemp <> ctlTxtDNI.Text Then
                            MsgBox "DNI/CE escaneado, no coincide", vbCritical + vbOKOnly, App.ProductName
                            b_Encontrado = False
                            b_monedero = False
                            objVenta.CodigoCliente = ""
                            Exit Sub
                        End If
                    End If
                    
                    Set rs = objVenta.fnListaAfiliadoMonedero(Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text)
                    vCodCliente = "" & rs("COD_CLIENTE")
                    ' si no existe solicitar datos
                    If vCodCliente = "" Then
                        b_Encontrado = False
                        b_monedero = True
                        objVenta.CodigoCliente = ""
                        Exit Sub
                    ' Si existe asociar tarjeta y cargar datos del afiliado
                    Else
                        FrmPedido_Ingre_Cli.Caption = "Programa Monedero del Ahorro - Afiliación"
                        FrmPedido_Ingre_Cli.Cbo_Tipo_Doc.Text = Mid(CboTipoDoc.Text, 1, 3)
                        FrmPedido_Ingre_Cli.ctlTxtDNI.Text = ctlTxtDNI.Text
                        FrmPedido_Ingre_Cli.b_monedero = b_monedero
                        FrmPedido_Ingre_Cli.b_afiliar = Not b_tiene_tarjetas
                        FrmPedido_Ingre_Cli.CargarCliente
                        FrmPedido_Ingre_Cli.cmdAceptar_Click
                        
                        objVenta.CodigoCliente = vCodCliente
                        objVenta.EsVentaMonedero = True
                        objVenta.NumeroTarjetaMonedero = vNroTarjeta
                        objVenta.PuntosTarjetaMonedero = CDbl(Val(vEstadoTarjeta(2)))
                        frmPedido_Busca_Cli.v_delfrm = vCodCliente
                        frmPedido.lbl_Cliente.Caption = "" & rs("DES_NOM_CLIENTE") & ", " & _
                                                        "" & rs("DES_APE_CLIENTE") & " " & _
                                                        "" & rs("DES_APE2_CLIENTE")

                        frmPedido.pstrDniCli = ctlTxtDNI.Text
                        frmPedido.pstrNomcli = frmPedido.lbl_Cliente.Caption
                        frmPedido.loadOptions
                        b_Encontrado = True
                        Exit Sub
                    End If
                End If
            Else
                b_monedero = False
                gstrCodTarjetaMon = ""
            End If

        'Si existe
        Else
            'Validar si tiene tarjetas afiliadas
            b_tiene_tarjetas = (oAff.Tarjetas.Count(1) > 0)
            
            'Si tiene tarjetas afiliadas
            If b_tiene_tarjetas Then
            
                ' Si es busqueda por DNI oDP.buscaFun(gstrCodTarjetaMon) <> "MONEDERO"
                If gstrCodTarjetaMon = "" Then
                    vNroTarjeta = ctlTxtDNI.Text 'vNroCuenta
                    gstrCodTarjetaMon = vNroTarjeta
                ' sino es una tarjeta adicional
                Else
                    vNroTarjeta = gstrCodTarjetaMon
                End If
                GoTo ActualizarDatos
            'No tiene tarjeta
            Else
                ' Si es busqueda por DNI
                If gstrCodTarjetaMon = "" Then
                    'vNroTarjeta = ctlTxtDNI.Text 'vNroCuenta
                    'gstrCodTarjetaMon = vNroTarjeta
                    GoTo FlujoAfiliacion
                ' sino es una tarjeta nueva
                Else
                    vNroTarjeta = gstrCodTarjetaMon
                    GoTo ActualizarDatos
                End If
            End If
        End If

'''        b_existe = oClienteF.CargarCliente(Mid(CboTipoDoc.Text, 1, 3), ctlTxtDNI.Text)
'''        If b_existe And oClienteF.TarjetasAsociadas.Count(1) = 0 Then
'''            oClienteF.AsociarTarjeta (FrmPedido_Ingre_trj.ObtenerTarjeta(""))
'''        End If
    End If
    ' ******* FIN MONEDERO MIFARMA *******
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub CboTipoDoc_Change()
    ctlTxtDNI.Text = Mid(ctlTxtDNI.Text, 1, CboTipoDoc.BoundText2)
    Debug.Print CboTipoDoc.BoundText2
    ctlTxtDNI.MaxLength = CboTipoDoc.BoundText2
End Sub

'Private Sub CboTipoDoc_Click(Area As Integer)
'    ctlTxtDNI.Text = Mid(ctlTxtDNI.Text, 1, CboTipoDoc.BoundText2)
'    ctlTxtDNI.MaxLength = CboTipoDoc.BoundText2
'End Sub

Private Sub chk_Editar_Click()
    ctlTxtDNI.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Label3.Visible Then
        Select Case KeyCode
            Case vbKeyF1
                If chk_Editar.Value = 0 Then
                    chk_Editar.Value = 1
                Else
                    chk_Editar.Value = 0
                End If
        End Select
    Else
        chk_Editar.Value = 0
    End If
End Sub
