VERSION 5.00
Begin VB.Form frm_VTA_Cotizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización"
   ClientHeight    =   2145
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3480
      TabIndex        =   4
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   2160
      TabIndex        =   3
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   4635
      Begin vbp_Ventas.ctlTextBox TxtNumCotizacion 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Tipo            =   7
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
      Begin VB.Label Label1 
         Caption         =   "Ingrese el número de la Proforma o Cotización en el recuadro de texto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "No. de Cotización:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frm_VTA_Cotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim objCliente As New clsCliente
'Dim objProveedor As New clsProveedor
Dim odynCab As oraDynaset
Dim odynDet As oraDynaset
Dim odynDetIns As oraDynaset
Dim odynDetFroPag As oraDynaset
Public pstrNumProf As String
Public strNombreFomularioOrigen As String
Public TipoDoc As Boolean
Dim strRucProv As String
Dim strCodProdRecMag As String

Private Sub cmdAceptar_Click()
On Error GoTo Control

    ConsultaProforma

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Load()
   TxtNumCotizacion.Text = pstrNumProf
End Sub

Private Sub TxtNumCotizacion_KeyPress(KeyAscii As Integer)
''   TxtNumCotizacion.Tipo = Entero
''   If KeyAscii = 13 Then
''        cmdAceptar_Click
''   End If
End Sub


Private Sub cmdCancelar_Click()
  frmPedido.lblTotalDescuento.Caption = "0.00"
    Unload Me
End Sub

Sub ConsultaProforma()
        Dim objProducto As New clsProducto
        Dim strCotizacion As String
        Dim Indicador As String
        Dim PctComi As Double
        Dim objConvenio As clsConvenio
        Dim objDocumento As clsDocumento
        'Dim rsReceta As oraDynaset
        Dim recCabReceta As oraDynaset
        Dim recDetReceta As oraDynaset
        
        Dim rsDatosAdicionales As oraDynaset
        Dim strDocProf As String
        Dim pxdbDatos As New XArrayDB
        Dim strDiagnostico As String
        Dim ObjDoc As clsProforma
        Dim gstrActualizaCotizacion As String
        
        'On Error GoTo CtrlErr
        
        
       
        
        Set ObjDoc = New clsProforma
        
        frmPedido.psub_BeginArry
        
        strCotizacion = Replace(Trim(TxtNumCotizacion.Text), "-", "")
        
        objVenta.NumPedidoPadre = strCotizacion
                
        Set odynCab = ObjDoc.ListaCabecera(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strCotizacion)
        Set odynDet = ObjDoc.ListaDetalle(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strCotizacion)
        Set odynDetIns = ObjDoc.ListaDetalleInsumosRM(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strCotizacion)
        Set odynDetFroPag = ObjDoc.ListaFormaPago(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strCotizacion)
        
        
        If odynCab.RecordCount = 0 Then
            MsgBox "La Proforma " & strCotizacion & " No existe, verifique ", vbInformation + vbOKOnly, App.ProductName
            TxtNumCotizacion.SetFocus
            Exit Sub
        End If
       
'''''  *********************************************************************************************************
'''''  Validacion para que en los locales no puedan cambiar el pedido de Delivery por la busqueda de la proforma
       
       If (odynCab("COD_ESTADO").Value = "005" Or odynCab("COD_ESTADO").Value = "006" Or _
          odynCab("COD_ESTADO").Value = "007" Or odynCab("COD_ESTADO").Value = "008" Or _
          odynCab("COD_ESTADO").Value = "009" Or odynCab("COD_ESTADO").Value = "010" Or _
          odynCab("COD_ESTADO").Value = "011") And (objUsuario.CodigoLocal <> "DLV") Then
          
            MsgBox "La Proforma " & strCotizacion & " de Delivery no se puede usar por esta opción, verifique ", vbInformation + vbOKOnly, App.ProductName
            Set odynCab = Nothing
            Set odynDet = Nothing
            Set odynDetIns = Nothing
            Set odynDetFroPag = Nothing
            Exit Sub
       End If
       
'''''  *********************************************************************************************************
'''''  *********************************************************************************************************

'''''''''        If objUsuario.TipoMaquina = objUsuario.TipoMaquinaCabina Then
'''''''''                mdiPrincipal.ctlCliente1.LocalAsignado = odynCab("COD_LOCAL_REF").Value
'''''''''                mdiPrincipal.ctlCliente1.Codigo = "" & odynCab("COD_CLIENTE").Value
'''''''''                'mdiPrincipal.ctlCliente1.DE = objCliente.DesSubFijo(odynCab("COD_SUFIJO").Value)
'''''''''                mdiPrincipal.ctlCliente1.Nombre = "" & odynCab("DES_NOM_CLIENTE").Value
'''''''''                mdiPrincipal.ctlCliente1.Apellido = "" & odynCab("DES_APE_CLIENTE").Value

        
        mdiPrincipal.subNuevo
        objVenta.NumPedidoPadre = strCotizacion
        objUsuario.CodTipoVenta = "" & odynCab("COD_TIPO_VENTA").Value
        
        
          If objUsuario.EsDelivery = True Then
            If MsgBox("¿Desea Actualizar los Precios? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                gstrActualizaCotizacion = "1"
            Else
                frm_VTA_ValidaCotizacion.Show vbModal
                If gstrValidaCotizacion = "2" Then
                    cmdCancelar_Click
                    Exit Sub
                Else
                    If gstrValidaCotizacion = "1" Then
                       gstrActualizaCotizacion = "0"
                       gstrValidaCotizacion = "0"
                    Else
                        MsgBox "", vbInformation + vbOKOnly, App.ProductName
                        Exit Sub
                    End If
                End If
            End If
         End If
        
        With objVenta
        ''llena la forma de pago de la proforma
            .CodigoTipoVenta = "" & odynCab("COD_MODALIDAD_VENTA").Value
            .CodModalidadVenta = "" & odynCab("COD_MODALIDAD_VENTA").Value
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            .ptmModalidad = "" & odynCab("COD_MODALIDAD_VENTA").Value
            .UbigeoEntrega = "" & odynCab("COD_UBIGEO").Value
            .bk_Ubigeo = .UbigeoEntrega
            mdiPrincipal.txtDLVTelefono.Text = "" & odynCab("DES_AUX_CLI_TLF").Value
            .CodigoClienteDLV = "" & odynCab("COD_CLIENTE_DLV").Value
            .NombreClienteDLV = "" & odynCab("DES_AUX_CLI_NOMBRE").Value
            .DireccionClienteDLV = "" & odynCab("DES_AUX_CLI_DIRECC").Value
            .DesReferenciaCli = "" & odynCab("DES_AUX_CLI_REF").Value
            .DesUrbanizacionDLV = "" & odynCab("URBANIZACION").Value
            .DesDistritoDLV = "" & odynCab("DISTRITO").Value
            .ObservacionClienteDLV = "" & odynCab("DES_OBSERVACION").Value
            .ObsNotaLocal = "" & odynCab("OBS_NOTA_LOCAL").Value
            .ObsNotaRuteo = "" & odynCab("OBS_NOTA_RUTEO").Value
            .ObsNotaMotorizado = "" & odynCab("OBS_NOTA_MOTORIZADO").Value
            .ObsNotaVerificacion = "" & odynCab("OBS_NOTA_VERIFICACION").Value
            
            .CodDireccionCli = "" & odynCab("COD_DIRECCION_CLI").Value
            .EntregaTercero = "" & odynCab("FLG_ENTREGA_TERCERO").Value
            .DesAuxRecogeNombre = "" & odynCab("DES_AUX_RECOGE_NOMBRE").Value
            .DesAuxRecogeDirecc = "" & odynCab("DES_AUX_RECOGE_DIRECC").Value
            .DesAuxRecogeRef = "" & odynCab("DES_AUX_RECOGE_REF").Value
            .DesAuxRecogeTlf = "" & odynCab("DES_AUX_RECOGE_TLF").Value
            .Ubigeo = "" & odynCab("COD_UBIGEO").Value
            'I.CVIERA 19.03.2021
            .bk_ServiceType = "" & odynCab("DELIVERY_TYPE").Value
            .bk_FechaCapacidad = "" & odynCab("FCH_HORA_PACT_ENTR").Value
            .bk_HoraCapacidad2 = "" & odynCab("HORA_SEGUNDA_PACT_ENTR").Value
            'F.CVIERA 19.03.2021
            strDocProf = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_PROF")
        
        If .CodigoTipoVenta = Venta_Convenio Then
            
            Set objConvenio = New clsConvenio
        
            .LimpiaServicio
            
            If .ptmModalidad = Guias_Remision Then
                ''frmPedido.psub_BeginArry
                mdiPrincipal.subNuevo
            End If
            frmPedido.grdPedido.Rebind
            Unload Me
            ptmTipoPrecio = Convenio
            .CodigoTipoVenta = Venta_Convenio
            frmPedido.lblModalidad.Caption = .NombreTipoVenta
            .ptmModalidad = Venta_Convenio
            frmPedido.Label4.Visible = True
            frmPedido.lblPctCopago.Visible = True
            frmPedido.Label8.Visible = True
            frmPedido.lblcopago.Visible = True
            frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = True
        
            ''''If objUsuario.EsDelivery = True Then
                .CodigoConvenio = "" & odynCab("COD_CONVENIO").Value
                .PctBeneficiario = val("" & odynCab("PCT_BENEFICIARIO").Value)
                .CodigoBeneficiario = "" & odynCab("COD_CLIENTE").Value
                .CodigoCliente = "" & odynCab("COD_CLIENTE").Value
                .NombreBeneficiario = "" & odynCab("BENEFICIARIO").Value
                .DesAuxCliNombre = "" & odynCab("BENEFICIARIO").Value
                .NombreConvenio = "" & odynCab("CONVENIO").Value
                .bk_codBeneficiario = "" & .CodigoBeneficiario
''''            Else
''''                .CodigoConvenio = frm_DLV_Pedido.grdPedidoDLV.Columns("COD_CONVENIO")
''''                .PctBeneficiario = "" & odynCab("PCT_BENEFICIARIO")
''''                .CodigoBeneficiario = frm_DLV_Pedido.grdPedidoDLV.Columns("COD_CLIENTE")
''''                .CodigoCliente = frm_DLV_Pedido.grdPedidoDLV.Columns("COD_CLIENTE")
''''                .NombreBeneficiario = frm_DLV_Pedido.grdPedidoDLV.Columns("DES_CLIENTE")
''''                .DesAuxCliNombre = frm_DLV_Pedido.grdPedidoDLV.Columns("DES_CLIENTE")
''''            End If
            
            
            
            
            Set objDocumento = New clsDocumento
            
            Set rsDatosAdicionales = objDocumento.ListaTipoCampo(objUsuario.CodigoEmpresa, strDocProf, strCotizacion, objUsuario.CodigoLocal)
            
            Set objDocumento = Nothing
            
            pxdbDatos.ReDim 0, -1, 0, 6
            Dim ultimo As Integer
            
            
            If Not rsDatosAdicionales.EOF Then
                                                    
                
                
                While Not rsDatosAdicionales.EOF
                    ultimo = pxdbDatos.Count(1)
                    pxdbDatos.AppendRows
                    pxdbDatos(ultimo, 0) = "" & rsDatosAdicionales("COD_TIPO_CAMPO").Value
                    pxdbDatos(ultimo, 1) = "" & rsDatosAdicionales("DES_TIPO_CAMPO").Value
                    pxdbDatos(ultimo, 2) = "" & rsDatosAdicionales("FLG_TIPO_DATO").Value
                    pxdbDatos(ultimo, 3) = "" & rsDatosAdicionales("CTD_LONG_MAX").Value
                    pxdbDatos(ultimo, 4) = "" & rsDatosAdicionales("CTD_LONG_MIN").Value
                    pxdbDatos(ultimo, 5) = "" & rsDatosAdicionales("DES_VALOR").Value
                    pxdbDatos(ultimo, 6) = "" & rsDatosAdicionales("FLG_EDITABLE").Value
                    rsDatosAdicionales.MoveNext
                Wend
                
                
                Set .DatosAdicional = pxdbDatos
                
                
                
            End If
            
            
            
            ''If .CodigoConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then
                Set recCabReceta = objConvenio.ListaCabeceraReceta(strDocProf, strCotizacion)
                If Not recCabReceta.EOF Then
                    .FechaReceta = "" & recCabReceta("FECHA").Value
                    .LocalReceta = "" & recCabReceta("COD_LOCAL_EMISION").Value
                    .CodMedico = "" & recCabReceta("COD_MEDICO").Value
                End If
                
                Set recDetReceta = objConvenio.ListaDetalleReceta(strDocProf, strCotizacion)
                While Not recDetReceta.EOF
                    .AgregaDiagnostico recDetReceta("COD_DIAGNOSTICO").Value, recDetReceta("DES_DIAGNOSTICO").Value
                    strDiagnostico = strDiagnostico & recDetReceta("COD_DIAGNOSTICO").Value & "|"
                    recDetReceta.MoveNext
                Wend
                .DetalleReceta = strDiagnostico
            ''Else
                .CodMotorizado = "" & odynCab("COD_MOTORIZADO_CONVENIO").Value
            ''End If
            
                If .CodMedico = "" Then
                    .CodMedico = "" & odynCab("COD_MEDICO").Value
                End If
            
            If objUsuario.CodTipoVenta = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_TIPO_VENTA_LOCAL") Then Unload frm_DLV_Pedido
            
            Set objConvenio = Nothing
            
        End If
        
        '** Permite pasarle los valores a estas propiedades **'
        '** Hecho 31/10/2007 Por Cristhian Rueda            **'
        
        ''objVenta.NomTitular = "" & odynDetFroPag("DES_NOM_TITULAR").Value
        ''objVenta.NumDNI = "" & odynDetFroPag("NUM_DOCUMENTO_IDENT").Value
        
        '*****************************************************'
        
        odynDetFroPag.MoveFirst
        While Not odynDetFroPag.EOF
            objVenta.AgregaFormaPago "" & odynDetFroPag("COD_FORMA_PAGO"), _
                                     "" & odynDetFroPag("DES_FORMA_PAGO"), _
                                     "" & odynDetFroPag("COD_HIJO"), _
                                     "" & odynDetFroPag("DES_HIJO"), _
                                     "" & odynDetFroPag("IMP_MONEDA_NAC"), _
                                     "" & odynDetFroPag("COD_TIPO_TARJETA"), _
                                     "" & odynDetFroPag("COD_MONEDA"), _
                                     "" & odynDetFroPag("COD_DOCUMENTO_PAGO"), _
                                     "" & odynDetFroPag("COD_BANCO"), _
                                     "" & odynDetFroPag("COD_DONACION"), "", _
                                     "" & odynDetFroPag("IMP_TIPO_CAMBIO"), _
                                     "" & odynDetFroPag("NUM_TARJETA"), _
                                     "" & odynDetFroPag("NUM_CUOTAS"), _
                                     "" & odynDetFroPag("FCH_VENCIMIENTO"), _
                                     "" & odynDetFroPag("FLG_CUOTA_NORMAL"), _
                                     "" & odynDetFroPag("NUM_DOCUMENTO_PAGO"), _
                                     "" & odynDetFroPag("NUM_MOVIMIENTO"), "", _
                                     "" & odynDetFroPag("FCH_MOVIMIENTO"), _
                                     "" & odynDetFroPag("NUM_AUTORIZACION"), _
                                     "" & odynDetFroPag("FCH_DOC_NOTA_CRED"), _
                                     "" & odynDetFroPag("NUM_DOC_NOTA_CRED"), "", "", _
                                     "" & odynDetFroPag("NUM_DOCUMENTO_IDENT").Value, _
                                     "", objUsuario.CodigoLocal, "", _
                                     "" & odynDetFroPag("FLG_RETIRO_EFEC"), "", "", "", _
                                     "" & odynDetFroPag("DES_NOM_TITULAR").Value  'Agredado por Cristhian Rueda

            odynDetFroPag.MoveNext
        Wend
        
        
        
        
        If odynCab Is Nothing Then MsgBox "Error en el acceso de datos", vbCritical, App.ProductName: Exit Sub
            'Carga la venta de datos del documento
        
        .NombreCliente = "" & odynCab("DES_NOM_CLIENTE").Value
        .ApellidoCliente = "" & odynCab("DES_APE_CLIENTE").Value
        .DireccionCliente = "" & odynCab("DIRECCION").Value
        .DireccionClienteSocial = "" & odynCab("DES_DIRECCION_SOCIAL").Value
        .CodigoCliente = "" & odynCab("COD_CLIENTE").Value
        
        .FchHoraPactEntr = "" & odynCab("FCH_HORA_PACT_ENTR").Value
        .HoraPactEntr = "" & odynCab("FCH_HORA_PACT_ENTR").Value
        
        .FchHoraPactRecog = "" & odynCab("FCH_HORA_PACT_RECOG").Value
        .HoraPactRecog = "" & odynCab("FCH_HORA_PACT_RECOG").Value
        
        .FlgEntregaLocal = "" & odynCab("FLG_ENTREGA_LOCAL").Value
        .FlgUrgente = "" & odynCab("FLG_URGENTE").Value
        .FlgPactado = "" & odynCab("FLG_FECHA_PACTADA").Value
        .CodigoDocumentoVenta = "" & odynCab("COD_TIPO_DOCUMENTO").Value
        
        If .CodigoDocumentoVenta = objUsuario.TipoDocBol Or .CodigoDocumentoVenta = .TipoDocTKB Then
            .NumeroDocumentoID = "" & odynCab("NUM_RUC_EMPRESA").Value
            ''comentado por pherrera 30/11/07 no imprimia nombre en boleta de DLV
            ''.NombreCliente = "" & odynCab("DES_RAZON_SOCIAL").Value
            .NombreCliente = "" & odynCab("DES_AUX_CLI_NOMBRE").Value
            .TipoCliente = "0"
            .DesAuxCliNombre = "" & .NombreClienteDLV
            .DesAuxCliTlf = "" & odynCab("DES_AUX_CLI_TLF").Value
            .DesAuxCliDirecc = "" & .DireccionClienteDLV
        End If
        
        
        If .CodigoDocumentoVenta = objUsuario.TipoDocFac Or .CodigoDocumentoVenta = .TipoDocTKF Then
            .Ruc = "" & odynCab("NUM_RUC_EMPRESA").Value
            .RazonSocial = "" & odynCab("DES_RAZON_SOCIAL").Value
            .TipoCliente = "1"
        End If
        
        
        
        
        .Out_NumeroId = "" & odynCab("NUM_RUC_EMPRESA").Value
        .Out_NombreCliente = "" & odynCab("DES_RAZON_SOCIAL").Value
        'Agregado por DJara 12/09/2008
        .Out_Direccion = "" & odynCab("DES_DIREC_SOCIAL").Value
        .Out_Tipo = .TipoCliente

        If odynCab("COD_CONVENIO").Value = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then
            .NumeroDocumentoID = ""
            .RazonSocial = ""
        End If
    
        ''** Se asigna a estos valores para la grabación de los convenios que solo emitan 1 documento                                                            ***** '
        ''** 05/02/2008 Por Cristhian Rueda                                                                                                                      ***** '
        Dim odynAuxCnv As oraDynaset
        Dim xstrTipDocEmp As String
        Dim xstrTipDocBenef As String
        
        Set odynAuxCnv = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DATOS_CONV_PROF", 0, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strCotizacion)
        If Not odynAuxCnv.EOF Then
            xstrTipDocEmp = "" & odynAuxCnv("CODTIPDOC_EMP").Value
            xstrTipDocBenef = "" & odynAuxCnv("CODTIPDOC_BENEF").Value
            .DesAuxCliDirecc = "" & odynAuxCnv("DIRECCION_X").Value
            If Trim(xstrTipDocEmp) <> "" And Trim(xstrTipDocBenef) = "" Then
                objVenta.TipoCliente = "" & odynAuxCnv("FLG_TIPO_JURIDICO_X").Value
                objVenta.Ruc = "" & odynAuxCnv("NUM_DOCUMENTO_X").Value
                objVenta.RazonSocial = "" & odynAuxCnv("RAZON_SOCIAL_X").Value
            End If
        End If
        
        '' ********************************************************************************************************************************************************** '
        
        ''COMENTADO POR JAHZEEL LOPEZ EL 19/11/2007 POR QUE SE CAMBIO POR EL CODIGO DE ARRIBA
        ''If odynCab("NUM_RUC_EMPRESA").Value <> "" Then
        ''        .Ruc = "" & odynCab("NUM_RUC_EMPRESA").Value
        ''        .TipoCliente = "1"
        ''        objVenta.Out_Tipo = "1"
        ''        objVenta.Out_NumeroId = "" & odynCab("NUM_RUC_EMPRESA").Value
        ''        objVenta.Out_NombreCliente = "" & odynCab("DES_RAZON_SOCIAL").Value
        ''Else
        ''        If odynCab("COD_CONVENIO").Value = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then
        ''            .NumeroDocumentoID = ""
        ''            .RazonSocial = ""
        ''        Else
        ''            .NumeroDocumentoID = "" & odynCab("num_documento_id").Value
        ''            .RazonSocial = "" & odynCab("NOMBRES").Value
        ''        End If
        ''        .TipoCliente = "0"
        ''        objVenta.Out_Tipo = "0"
        ''End If
        
        End With
        

        
        
        
  If objVenta.CodigoTipoVenta = Recetario Then
        If odynDetIns.RecordCount > 0 Then
                strRucProv = ObjDoc.Dev_RucProv_RecMag_x_Btl(objUsuario.CodigoLocal)
                
                
                '-- Proveedor de recetario magistral --'
''                Set frm_VTA_RecetarioM.ctlCboProveedor.RowSource = objProveedor.ListaRegMagistral(strRucProv, "1", objUsuario.CodigoLocal)
''                frm_VTA_RecetarioM.ctlCboProveedor.ListField = "NOM_PROVEEDOR"
''                frm_VTA_RecetarioM.ctlCboProveedor.BoundColumn = "RUC_PROVEEDOR"
''                frm_VTA_RecetarioM.ctlCboProveedor.BoundText = objVenta.ProvPreDeterminadoRM
                '-- xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx --'
                
                frm_VTA_RecetarioM.txtProveedor.Text = "" & Trim(strRucProv)
                frm_VTA_RecetarioM.txtCliente.Text = "" & odynCab("COD_CLIENTE").Value
                frm_VTA_RecetarioM.lblCliente.Caption = "" & odynCab("NOMBRES").Value
                frm_VTA_RecetarioM.TxtMedico.Text = "" & odynCab("COD_MEDICO").Value
                frm_VTA_RecetarioM.lblMedico.Caption = "" & odynCab("MEDICO").Value
                
                odynDetIns.MoveFirst
                While Not odynDetIns.EOF
                    Call frm_VTA_RecetarioM.psub_Agrega_Insumo(odynDetIns("COD_TIPO_INSUMO").Value, _
                                                               odynDetIns("DES_TIPO_INSUMO").Value, _
                                                               odynDetIns("COD_PRODUCTO_INS").Value, _
                                                               odynDetIns("DES_PRODUCTO").Value, _
                                                               odynDetIns("DES_UND_CAPACIDAD_ABREV").Value, _
                                                               odynDetIns("PCT_BASE").Value, _
                                                               odynDetIns("CTD_BASE").Value, _
                                                               odynDetIns("CTD_PRODUCTO").Value, _
                                                               odynDetIns("IMP_PRECIO").Value, _
                                                               odynDetIns("SUBTOTAL").Value, _
                                                               odynDetIns("COD_PRODUCTO").Value)
                    odynDetIns.MoveNext
                Wend
                frm_VTA_RecetarioM.GrdInsumos.Rebind
                ''frm_VTA_RecetarioM.Hide
                ''frm_VTA_RecetarioM.Show
                              
                                    
                Dim k%
                k = 0
                strCodProdRecMag = objProducto.ListaDevRM(objUsuario.CodigoLocal, strRucProv)

                For k = 0 To frm_VTA_RecetarioM.pxdbInsumos.UpperBound(1)
                    ''If k = 0 Then
                    objVenta.AgregaRecetarioM frm_VTA_RecetarioM.pxdbInsumos(k, 11), _
                                              frm_VTA_RecetarioM.pxdbInsumos(k, 3), _
                                              frm_VTA_RecetarioM.pxdbInsumos(k, 7), _
                                              "0", _
                                              frm_VTA_RecetarioM.pxdbInsumos(k, 9), _
                                              objVenta.CodigoTipoVenta, _
                                              frm_VTA_RecetarioM.pxdbInsumos(k, 6), _
                                              frm_VTA_RecetarioM.pxdbInsumos(k, 8), _
                                              frm_VTA_RecetarioM.pxdbInsumos(k, 12)
                    ''End If
                Next k
                frmPedido.Cal_Promo
                frmPedido.Cal_Montos
                frmPedido.grdPedido.Rebind
                
                
                Unload Me
        End If
  End If
        odynDet.MoveFirst
        While Not odynDet.EOF
            Indicador = objProducto.CodIndicadorReceta(odynDet("COD_PRODUCTO").Value)
            PctComi = objProducto.pctComision(odynDet("COD_PRODUCTO").Value, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
            If val(odynDet("FLG_REGALO").Value) = 0 Then
                objVenta.AgregaProducto odynDet("COD_PRODUCTO").Value, odynDet("DES_PRODUCTO").Value, odynDet("CTD_PRODUCTO").Value, odynDet("FLG_FRACCIONAMIENTO").Value, odynDet("MTO_SUBTOTAL").Value, odynDet("COD_MODALIDAD").Value, Producto_Normal, , , , , , Indicador, PctComi, , , , , "" & odynDet("CAD_PROMOCION").Value, "" & odynDet("CAD_PROMOCION_DCTO").Value
            Else
                objVenta.AgregaProducto odynDet("COD_PRODUCTO").Value, odynDet("DES_PRODUCTO").Value, odynDet("CTD_PRODUCTO").Value, odynDet("FLG_FRACCIONAMIENTO").Value, odynDet("MTO_SUBTOTAL").Value, odynDet("COD_MODALIDAD").Value, Producto_Regalo, , , , , , Indicador, PctComi, , , , , "" & odynDet("CAD_PROMOCION").Value, "" & odynDet("CAD_PROMOCION_DCTO").Value
            End If
            odynDet.MoveNext
        Wend
        
        Set objProducto = Nothing
        
        frmPedido.grdPedido.Rebind
        frmPedido.Cal_Montos
        Unload Me
  
        If objUsuario.EsDelivery = True Then
            Set objDocumento = New clsDocumento
            frmPedido.lblSiguiente.Caption = objVenta.CodigoDocumentoVenta & " - " & objDocumento.ListaNumeroDisponible(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objVenta.CodigoDocumentoVenta)
            frm_VTA_Documento.blnTipoDoc = True
            frm_VTA_Documento.strDlvDocumento = "Usted Efectuara el Pago con" & "  " & objVenta.CodigoDocumentoVenta
            Set objDocumento = Nothing
            
            objVenta.DesAuxCliTlf = "" & odynCab("DES_AUX_CLI_TLF").Value
            
            
            
            mdiPrincipal.ctlCliente1.Cargar objVenta.DesAuxCliTlf
            mdiPrincipal.ctlCliente1.CodDireccionCli = objVenta.CodDireccionCli
            If objUsuario.CodLocalCallCenter = "1DLV" Then 'ECASTILLO 12.06.2020 - para que siempre muestre locales inka si es ejecutado por call inka
                mdiPrincipal.ctlCliente1.ConsultaCliente ("" & odynCab("COD_CLIENTE_DLV").Value), "94"
            Else
                mdiPrincipal.ctlCliente1.ConsultaCliente ("" & odynCab("COD_CLIENTE_DLV").Value)
            End If
            mdiPrincipal.txtDireccion.Text = mdiPrincipal.ctlCliente1.DireccionSocial
            
'            If MsgBox("Desea Actualizar los Precios .. ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'                frmPedido.Cal_Promo
'            Else
'                frm_VTA_ValidaCotizacion.Show vbModal
'                If gstrValidaCotizacion = "1" Then
'                   gstrValidaCotizacion = "0"
'                   'frmPedido.lblTotalDescuento.Caption = "0.00"
'                Else
'                    MsgBox "", vbInformation + vbOKOnly, App.ProductName
'                    Exit Sub
'                End If
'            End If

            If gstrActualizaCotizacion = "1" Then
                frmPedido.Cal_Promo
            End If
            
            
            'frmPedido.Cal_Montos
        End If
  '-- Carga todo los pedidos de delivery --'
  If objVenta.CodigoTipoVenta = Pedido_DLV Then
        pstrNumProf = Trim(TxtNumCotizacion.Text)
        frm_DLV_Pedido.Show
  End If
  
''  mdiPrincipal.TipoDoc = ""
  
  If frm_DLV_Pedido.blnActivaPed = True Then
    If TipoDoc = True Then
        objVenta.NumDocRef = pstrNumProf
        objVenta.CodDocRef = "PRO"
        
        Dim objConfig As New clsConfig
        
        frm_VTA_GuiaRemision.cboTipoDev.BoundText = objConfig.Valor(4, "TIPO_VTA")
        
        
        frm_VTA_GuiaRemision.cboMotivoDev.BoundText = objConfig.Valor(4, "MOTIVO_DLV")
        
        Set objConfig = Nothing
        
        frm_VTA_GuiaRemision.cboOrigen.BoundText = frm_DLV_Pedido.grdPedidoDLV.Columns("COD_LOCAL_REF")
        frm_VTA_GuiaRemision.cmdAceptar_Click
        
        If frm_VTA_GuiaRemision.FlgErr = False Then
            Call ObjDoc.CambiaEstado(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, pstrNumProf, "005", objUsuario.Codigo)
        Else
            Unload frm_VTA_GuiaRemision
        End If
    Else
        ''mdiPrincipal.xCodLocal = Mid(frm_DLV_Pedido.grdPedidoDLV.Columns("COD_MODALIDAD_VENTA"), 1, 3)
        ''mdiPrincipal.xNumProforma = pstrNumProf
        ''mdiPrincipal.TipoDoc = Mid(frm_DLV_Pedido.grdPedidoDLV.Columns("COD_MODALIDAD_VENTA"), 1, 3)
        mdiPrincipal.strNombreFomularioOrigen = Me.strNombreFomularioOrigen
        mdiPrincipal.cmdGrabaVenta_Click
    End If
        Dim objProforma As clsProforma
        Set objProforma = New clsProforma
        Set frm_DLV_Pedido.odynPedido = objProforma.ListaPedidoDLV(objUsuario.CodigoEmpresa, _
                                                                   objUsuario.CodigoLocal, odynCab("FCH_REGISTRA").Value, odynCab("FCH_REGISTRA").Value, odynCab("COD_ESTADO").Value)
        Set objProforma = Nothing
        Set frm_DLV_Pedido.grdPedidoDLV.DataSource = frm_DLV_Pedido.odynPedido
  End If
  
    Set ObjDoc = Nothing

    frmPedido.Cal_Montos
    
    Set odynCab = Nothing
    Set odynDet = Nothing
    Set odynDetIns = Nothing
    Set odynDetFroPag = Nothing
    Set recCabReceta = Nothing
    Set recDetReceta = Nothing
    Set rsDatosAdicionales = Nothing
    
'CtrlErr:
   ' MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub



'Private Sub LimpiaVariables()
'
'    Set objVenta = Nothing
'
'End Sub


