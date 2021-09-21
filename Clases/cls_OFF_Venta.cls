VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_OFF_Venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private xDetalleVenta As New XArrayDB
Private xPagoVenta As New XArrayDB
Private xInaFectos As New XArrayDB
Private xAfectos As New XArrayDB
Private strTipoDocumento As String
Private strNumDocumento As String
Private intSecuencia As Integer
Private strCodDocDefault As String
Private strNumDocDefault As String
Private dblTotalIgv As Double
Private dblTotalExonerado As Double
Private dblTotalVenta As Double
Private dblTotalPago As Double
Private dblTotalVuelto As Double
Private dblTotalDolares As Double
Private xDocumentosGenerados As New XArrayDB

Public Property Get ColDPTipoCuota() As Integer
    ColDPTipoCuota = 10
End Property

Public Property Get ColDPNumAutorizacion() As Integer
    ColDPNumAutorizacion = 11
End Property

Public Property Get ColDVDescripcionCorta() As Integer
    ColDVDescripcionCorta = 16
End Property

Public Property Get ColDVPrecioPublico() As Integer
    ColDVPrecioPublico = 17
End Property

Public Property Get ColDVConReceta() As Integer
    ColDVConReceta = 18
End Property

Public Property Get ColDVFlgFracciona() As Integer
    ColDVFlgFracciona = 19
End Property

Public Property Get ColDetalleVenta() As Integer
    ColDetalleVenta = 19
End Property

Public Property Get ColDVCodigo() As Integer
    ColDVCodigo = 0
End Property

Public Property Get ColDVDescripcion() As Integer
    ColDVDescripcion = 1
End Property

Public Property Get ColDVCantidad() As Integer
    ColDVCantidad = 3
End Property

Public Property Get ColDVFlgFraccion() As Integer
    ColDVFlgFraccion = 2
End Property

Public Property Get ColDVPctDescuento() As Integer
    ColDVPctDescuento = 4
End Property

Public Property Get ColDVPrcUnitario() As Integer
    ColDVPrcUnitario = 5
End Property

Public Property Get ColDVPrcOriginal() As Integer
    ColDVPrcOriginal = 6
End Property

Public Property Get ColDVMtoIgv() As Integer
    ColDVMtoIgv = 7
End Property

Public Property Get ColDVMtoExonerado() As Integer
    ColDVMtoExonerado = 8
End Property

Public Property Get ColDVMtoSubtotal() As Integer
    ColDVMtoSubtotal = 9
End Property

Public Property Get ColDVFlgModPrecio() As Integer
    ColDVFlgModPrecio = 10
End Property

Public Property Get ColDVFlgRegalo() As Integer
    ColDVFlgRegalo = 11
End Property

Public Property Get ColDVPctIgv() As Integer
    ColDVPctIgv = 12
End Property

Public Property Get ColDVPartidaArancelaria() As Integer
    ColDVPartidaArancelaria = 13
End Property

Public Property Get ColDVCtdFraccion() As Integer
    ColDVCtdFraccion = 14
End Property

Public Property Get ColDVUsuModPrecio() As Integer
    ColDVUsuModPrecio = 15
End Property

Public Property Get ColDetallePago() As Integer
    ColDetallePago = 11
End Property

Public Property Get ColDPCodPago() As Integer
    ColDPCodPago = 0
End Property

Public Property Get ColDPCodMoneda() As Integer
    ColDPCodMoneda = 2
End Property

Public Property Get ColDPMtoImporte() As Integer
    ColDPMtoImporte = 3
End Property

Public Property Get ColDPMtoSoles() As Integer
    ColDPMtoSoles = 4
End Property

Public Property Get ColDPMtoVuelto() As Integer
    ColDPMtoVuelto = 5
End Property

Public Property Get ColDPNumTarjeta() As Integer
    ColDPNumTarjeta = 6
End Property

Public Property Get ColDPFchVencimiento() As Integer
    ColDPFchVencimiento = 7
End Property

Public Property Get ColDPNumCuotas() As Integer
    ColDPNumCuotas = 8
End Property

Public Property Get ColDPTipoCambio() As Integer
    ColDPTipoCambio = 9
End Property

Public Property Get ColDPDescripcion() As Integer
    ColDPDescripcion = 1
End Property

Public Property Get TotalIgv() As Double
    TotalIgv = dblTotalIgv
End Property

Public Property Let TotalIgv(ByVal newValue As Double)
    dblTotalIgv = newValue
End Property

Public Property Get TotalExonerado() As Double
    TotalExonerado = dblTotalExonerado
End Property

Public Property Let TotalExonerado(ByVal newValue As Double)
    dblTotalExonerado = newValue
End Property

Public Property Get TotalVenta() As Double
    TotalVenta = dblTotalVenta
End Property

Public Property Let TotalVenta(ByVal newValue As Double)
    dblTotalVenta = newValue
End Property

Public Property Get TotalPago() As Double
    TotalPago = dblTotalPago
End Property

Public Property Let TotalPago(ByVal newValue As Double)
    dblTotalPago = newValue
End Property

Public Property Get TotalVuelto() As Double
    TotalVuelto = dblTotalVuelto
End Property

Public Property Let TotalVuelto(ByVal newValue As Double)
    dblTotalVuelto = newValue
End Property

Public Sub Graba()
    Dim i As Integer
    Dim strSql As String
    Dim rsDetalleVenta As New ADODB.Recordset
    Dim rsPagoVenta As New ADODB.Recordset
    
    Dim dblMtoSubtotal As Double
    Dim dblMtoSoles As Double
    Dim bolNoHayTC As Boolean
    Dim intLineas As Integer
    Dim xTempAfectos As XArrayDB
    Dim xTempInaFectos As XArrayDB
    Dim dblVuelto As Double
'    Dim cnn As ADODB.Connection

    On Error GoTo CtrlErr

'    Set cnn = New ADODB.Connection
'    cnn.Open gstrConexion
        
    If DetalleVenta.Count(1) = 0 Then
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "Se debe seleccionar al menos un producto"
    End If
        
    If PagoVenta.Count(1) = 0 Then
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "Se debe indicar al menos una forma de pago"
    End If
        
    Secuencia = SecuenciaVta
    If Secuencia < 1 Then
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "No se encontró la secuencia de venta"
    End If
        
    dblVuelto = 0
    For i = 0 To PagoVenta.UpperBound(1)
        If PagoVenta(i, ColDPCodPago) = COD_FPAGO_TARJETA Then
            dblVuelto = dblVuelto + PagoVenta(i, ColDPMtoVuelto)
            PagoVenta(i, ColDPMtoVuelto) = 0
        End If
    Next i

    If dblVuelto > 0 Then
        For i = 0 To PagoVenta.UpperBound(1)
            If PagoVenta(i, ColDPCodPago) = COD_FPAGO_EFE_SOLES Or PagoVenta(i, ColDPCodPago) = COD_FPAGO_EFE_DOLAR Then
                PagoVenta(i, ColDPMtoVuelto) = PagoVenta(i, ColDPMtoVuelto) + dblVuelto
                Exit For
            End If
        Next i
    End If
        
    dblMtoSubtotal = 0
        
    For i = 0 To DetalleVenta.UpperBound(1)
        dblMtoSubtotal = dblMtoSubtotal + DetalleVenta(i, 9)
    Next i
        
    dblMtoSoles = 0
    bolNoHayTC = False
    For i = 0 To PagoVenta.UpperBound(1)
        dblMtoSoles = dblMtoSoles + PagoVenta(i, ColDPMtoSoles)
        If PagoVenta(i, ColDPTipoCambio) = 0 Then bolNoHayTC = True
    Next i
        
    If bolNoHayTC Then
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "Tipo de cambio inválido"
    End If
        
    If Round(dblMtoSoles, 2) < Round(dblMtoSubtotal, 2) Then
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "El importe de la forma de pago es menor al total del documento"
    End If
        
    If TipoDocumento = COD_TIPO_BOL And dblMtoSubtotal > 1500 And objOFFCliente.Nombre = "" Then
        'frm_OFF_Documento.Show vbModal
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "Se debe ingresar el nombre del cliente"
    End If
    
    If TipoDocumento = COD_TIPO_FAC Then
        If objOFFCliente.ValidaRuc(objOFFCliente.Ruc) <> 0 Then
            Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "El número de RUC no es Válido"
        End If
    End If
        
'        strSql = "select * from pagoventa.txt"
'        rsPagoVenta.Open strSql, cnn, adOpenStatic, adLockOptimistic, adCmdText
    rsPagoVenta.CursorLocation = adUseClient
    rsPagoVenta.Open strPagoVentaXML, gstrConexion, adOpenDynamic, adLockOptimistic
        
    rsPagoVenta.Filter = "SEC_VENTA = " & Secuencia
    If rsPagoVenta.RecordCount > 0 Then
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "Documento ya tiene registrada una forma de pago"
    End If
    rsPagoVenta.Filter = adFilterNone
    
'        strSql = "select * from detalleventa.txt"
'        rsDetalleVenta.Open strSql, cnn, adOpenStatic, adLockOptimistic, adCmdText
    rsDetalleVenta.CursorLocation = adUseClient
    rsDetalleVenta.Open strDetalleVentaXML, gstrConexion, adOpenDynamic, adLockOptimistic
    
    LimpiaAfectos
    LimpiaInafectos
        
'        cnn.BeginTrans
        
    For i = 0 To DetalleVenta.UpperBound(1)
        If DetalleVenta(i, 8) > 0 Then
            Set xTempInaFectos = AgregaInafectos(DetalleVenta(i, ColDVCodigo), _
                                                 DetalleVenta(i, ColDVDescripcionCorta), _
                                                 DetalleVenta(i, ColDVCantidad), _
                                                 DetalleVenta(i, ColDVFlgFraccion), _
                                                 DetalleVenta(i, ColDVPctDescuento), _
                                                 DetalleVenta(i, ColDVPrcUnitario), _
                                                 DetalleVenta(i, ColDVPrcOriginal), _
                                                 DetalleVenta(i, ColDVMtoIgv), _
                                                 DetalleVenta(i, ColDVMtoExonerado), _
                                                 DetalleVenta(i, ColDVMtoSubtotal), _
                                                 DetalleVenta(i, ColDVFlgModPrecio), _
                                                 DetalleVenta(i, ColDVFlgRegalo), _
                                                 DetalleVenta(i, ColDVPctIgv), _
                                                 DetalleVenta(i, ColDVPartidaArancelaria), _
                                                 DetalleVenta(i, ColDVCtdFraccion), _
                                                 DetalleVenta(i, ColDVUsuModPrecio))
        Else
            Set xTempAfectos = AgregaAfectos(DetalleVenta(i, ColDVCodigo), _
                                             DetalleVenta(i, ColDVDescripcionCorta), _
                                             DetalleVenta(i, ColDVCantidad), _
                                             DetalleVenta(i, ColDVFlgFraccion), _
                                             DetalleVenta(i, ColDVPctDescuento), _
                                             DetalleVenta(i, ColDVPrcUnitario), _
                                             DetalleVenta(i, ColDVPrcOriginal), _
                                             DetalleVenta(i, ColDVMtoIgv), _
                                             DetalleVenta(i, ColDVMtoExonerado), _
                                             DetalleVenta(i, ColDVMtoSubtotal), _
                                             DetalleVenta(i, ColDVFlgModPrecio), _
                                             DetalleVenta(i, ColDVFlgRegalo), _
                                             DetalleVenta(i, ColDVPctIgv), _
                                             DetalleVenta(i, ColDVPartidaArancelaria), _
                                             DetalleVenta(i, ColDVCtdFraccion), _
                                             DetalleVenta(i, ColDVUsuModPrecio))
        End If
    Next i
            
    '----- Genera documentos afectos
    If Afectos.Count(1) > 0 Then
        GeneraDoc Secuencia, xTempAfectos, rsDetalleVenta
    End If
    
    '----- Genera documentos inafectos
    If Inafectos.Count(1) > 0 Then
        GeneraDoc Secuencia, xTempInaFectos, rsDetalleVenta
    End If

    For i = 0 To PagoVenta.UpperBound(1)
        rsPagoVenta.AddNew
        rsPagoVenta!SEC_VENTA = Secuencia
        rsPagoVenta!NUM_ITEM = i + 1
        rsPagoVenta!COD_PAGO = PagoVenta(i, ColDPCodPago)
        rsPagoVenta!COD_MONEDA = PagoVenta(i, ColDPCodMoneda)
        rsPagoVenta!MTO_IMPORTE = PagoVenta(i, ColDPMtoImporte)
        rsPagoVenta!MTO_SOLES = PagoVenta(i, ColDPMtoSoles)
        rsPagoVenta!MTO_VUELTO.Value = Trim(PagoVenta(i, ColDPMtoVuelto))
        rsPagoVenta!NUM_TARJETA = PagoVenta(i, ColDPNumTarjeta)
        rsPagoVenta!FCH_VENCIMIENTO = PagoVenta(i, ColDPFchVencimiento)
        rsPagoVenta!NUM_CUOTAS = PagoVenta(i, ColDPNumCuotas)
        rsPagoVenta!TIPO_CAMBIO = PagoVenta(i, ColDPTipoCambio)
        rsPagoVenta!TIPO_CUOTA = PagoVenta(i, ColDPTipoCuota)
        rsPagoVenta!NUM_AUTORIZACION = PagoVenta(i, ColDPNumAutorizacion)
        rsPagoVenta.Update
    Next i

    'rsPagoVenta.UpdateBatch
    'cnn.CommitTrans
        
    Call ActualizaSecuenciaVta
        
    rsDetalleVenta.Save strDetalleVentaXML, adPersistXML
    rsDetalleVenta.Close
    
    rsPagoVenta.Save strPagoVentaXML, adPersistXML
    rsPagoVenta.Close
'    cnn.Close
    Exit Sub
                        
CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Venta", Err.Description
'    cnn.RollbackTrans
End Sub

Public Property Get TipoDocumento() As String
    TipoDocumento = strTipoDocumento
End Property

Public Property Let TipoDocumento(ByVal lstrTipoDocumento As String)
    strTipoDocumento = lstrTipoDocumento
End Property

Public Property Get NumDocumento() As String
    NumDocumento = strNumDocumento
End Property

Public Property Let NumDocumento(ByVal lstrNumDocumento As String)
    strNumDocumento = lstrNumDocumento
End Property

Public Property Get Secuencia() As Integer
    Secuencia = intSecuencia
End Property

Public Property Let Secuencia(ByVal lintSecuencia As Integer)
    intSecuencia = lintSecuencia
End Property

Private Sub Class_Initialize()
    xDetalleVenta.ReDim 0, -1, 0, ColDetalleVenta
    xPagoVenta.ReDim 0, -1, 0, ColDetallePago
    xInaFectos.ReDim 0, -1, 0, 11
    xAfectos.ReDim 0, -1, 0, 11
    xDocumentosGenerados.ReDim 0, -1, 0, 1
End Sub

Public Function LimpiaDocumentosGenerados()
    xDocumentosGenerados.ReDim 0, -1, 0, 1
End Function

Public Function LimpiaDetalle()
    xDetalleVenta.ReDim 0, -1, 0, ColDetalleVenta

End Function

Public Function LimpiaPago()
    xPagoVenta.ReDim 0, -1, 0, ColDetallePago

End Function

Public Function LimpiaAfectos()
    xAfectos.ReDim 0, -1, 0, 15
End Function

Public Function LimpiaInafectos()
    xInaFectos.ReDim 0, -1, 0, 15
End Function

Public Property Get DetalleVenta() As XArrayDB
    Set DetalleVenta = xDetalleVenta
End Property

Public Property Get PagoVenta() As XArrayDB
    Set PagoVenta = xPagoVenta
End Property

Public Property Get Inafectos() As XArrayDB
    Set Inafectos = xInaFectos
End Property

Public Property Get Afectos() As XArrayDB
    Set Afectos = xAfectos
End Property

Public Function AgregaDetalleVenta(Codigo As String, _
                            Descripcion As String, _
                            FlagFraccion As String, _
                            Cantidad As Integer, _
                            PctDescuento As Double, _
                            PrcUnitario As Double, _
                            PrcOriginal As Double, _
                            MtoIgv As Double, _
                            MtoExonerado As Double, _
                            MtoSubTotal As Double, _
                            FlgModPrecio As String, _
                            FlgRegalo As String, _
                            PctIgv As Double, _
                            partidaArancel As String, _
                            CtdFraccion As Integer, _
                            UsuModPrecio As String, _
                            DescripcionCorta As String, _
                            PrecioPublico As Double, _
                            ConReceta As String, _
                            FlgFracciona As String) As XArrayDB
    Dim ultimo As Integer
    Dim aux As Integer
    If xDetalleVenta.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    encontro = False
    
    aux = xDetalleVenta.Count(1)
    While i < aux
        If xDetalleVenta(i, ColDVCodigo) = Codigo And xDetalleVenta(i, ColDVFlgRegalo) = FlgRegalo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xDetalleVenta.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xDetalleVenta.AppendRows
    End If
j:
    If xDetalleVenta.Count(1) = 0 Then ultimo = 0: xDetalleVenta.AppendRows
    
    
    xDetalleVenta(ultimo, ColDVCodigo) = Codigo
    xDetalleVenta(ultimo, ColDVDescripcion) = Descripcion
    xDetalleVenta(ultimo, ColDVFlgFraccion) = FlagFraccion
    xDetalleVenta(ultimo, ColDVCantidad) = Cantidad
    xDetalleVenta(ultimo, ColDVPctDescuento) = PctDescuento
    xDetalleVenta(ultimo, ColDVPrcUnitario) = PrcUnitario
    xDetalleVenta(ultimo, ColDVPrcOriginal) = PrcOriginal
    xDetalleVenta(ultimo, ColDVMtoIgv) = MtoIgv
    xDetalleVenta(ultimo, ColDVMtoExonerado) = MtoExonerado
    xDetalleVenta(ultimo, ColDVMtoSubtotal) = MtoSubTotal
    xDetalleVenta(ultimo, ColDVFlgModPrecio) = FlgModPrecio
    xDetalleVenta(ultimo, ColDVFlgRegalo) = FlgRegalo
    xDetalleVenta(ultimo, ColDVPctIgv) = PctIgv
    xDetalleVenta(ultimo, ColDVPartidaArancelaria) = partidaArancel
    xDetalleVenta(ultimo, ColDVCtdFraccion) = CtdFraccion
    xDetalleVenta(ultimo, ColDVUsuModPrecio) = UsuModPrecio
    xDetalleVenta(ultimo, ColDVDescripcionCorta) = DescripcionCorta
    xDetalleVenta(ultimo, ColDVPrecioPublico) = PrecioPublico
    xDetalleVenta(ultimo, ColDVConReceta) = ConReceta
    xDetalleVenta(ultimo, ColDVFlgFracciona) = FlgFracciona
    
    Set AgregaDetalleVenta = xDetalleVenta
    
End Function

Public Function AgregaPagoVenta(CodPago As String, _
                                CodMoneda As String, _
                                MtoImporte As Double, _
                                MtoSoles As Double, _
                                NumTarjeta As String, _
                                FchVencimiento As String, _
                                NumCuotas As Integer, _
                                TipoCambio As Double, _
                                TipoCuota As Integer, _
                                NumAutorizacion As String) As XArrayDB
    Dim ultimo As Integer
    Dim aux As Integer
    Dim MtoVuelto As Double
    Dim dblPagado As Double
    Dim dblVuelto As Double
    Dim objFormaPago As New cls_OFF_FormaPago

    If xPagoVenta.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    encontro = False
    dblPagado = MtoSoles
    dblVuelto = 0
    aux = xPagoVenta.Count(1)
    While i < aux
        If xPagoVenta(i, ColDPCodPago) = CodPago And xPagoVenta(i, Me.ColDPNumTarjeta) = NumTarjeta Then
            ultimo = i
            encontro = True
            GoTo j
        Else
            encontro = False
            ultimo = xPagoVenta.Count(1)
            dblPagado = dblPagado + xPagoVenta(i, ColDPMtoSoles)
            dblVuelto = dblVuelto + xPagoVenta(i, ColDPMtoVuelto)
        End If
        i = i + 1
    Wend
    
    If encontro = False Then
        xPagoVenta.AppendRows
    End If
j:
    If xPagoVenta.Count(1) = 0 Then ultimo = 0: xPagoVenta.AppendRows
    
    
    MtoVuelto = (dblPagado - dblVuelto) - TotalVenta
    If MtoVuelto < 0 Then MtoVuelto = 0

    xPagoVenta(ultimo, ColDPCodPago) = CodPago
    xPagoVenta(ultimo, ColDPDescripcion) = objFormaPago.DescripcionFormaPago(CodPago)
    xPagoVenta(ultimo, ColDPCodMoneda) = CodMoneda
    xPagoVenta(ultimo, ColDPMtoImporte) = Trim(MtoImporte)
    xPagoVenta(ultimo, ColDPMtoSoles) = Trim(MtoSoles)
    xPagoVenta(ultimo, ColDPMtoVuelto) = Round(Trim(MtoVuelto), 5)
    xPagoVenta(ultimo, ColDPNumTarjeta) = Trim(NumTarjeta)
    xPagoVenta(ultimo, ColDPFchVencimiento) = FchVencimiento
    xPagoVenta(ultimo, ColDPNumCuotas) = NumCuotas
    xPagoVenta(ultimo, ColDPTipoCambio) = TipoCambio
    xPagoVenta(ultimo, ColDPTipoCuota) = TipoCuota
    xPagoVenta(ultimo, ColDPNumAutorizacion) = Trim(NumAutorizacion)
    
    Set objFormaPago = Nothing
    Set AgregaPagoVenta = xPagoVenta
    
End Function

Public Sub GeneraDoc(ByVal intSecuencia As Integer, _
                     ByVal xTemp As XArrayDB, _
                     ByRef rs As ADODB.Recordset)

    Dim objDocumento As cls_OFF_Documento
    Dim intTotDoc As Integer
    Dim intLineas As Integer
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer

    On Error GoTo CtrlErr

    Set objDocumento = New cls_OFF_Documento

    intLineas = Val(objDocumento.NumeroLineas(TipoDocumento))
    
    If intLineas < 1 Then
        Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "No se encontró el número de líneas para el documento " & TipoDocumento
    End If
    
    intTotDoc = Int(xTemp.Count(1) / intLineas) + IIf(xTemp.Count(1) Mod intLineas > 0, 1, 0)
    
    n = 0: i = 0
    While n < intTotDoc
        j = 0
        NumDocumento = objDocumento.UltimoCorrelativo(TipoDocumento)
        AgregaDocumentosGenerados TipoDocumento, NumDocumento
        If NumDocumento = "" Or Len(NumDocumento) <> 10 Then
            Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", "Se debe indicar el correlativo del documento"
        End If
        
        rs.Filter = "COD_CIA = '" & objOFFUsuario.Cia & "' AND COD_TIPO_DOCUMENTO = '" & TipoDocumento & "' AND NUM_DOCUMENTO = '" & NumDocumento & "'"
        If rs.RecordCount > 0 Then
            Err.Raise vbObjectError + 513, "cls_OFF_Venta.Graba", IIf(Val(TipoDocumento) < 2, "La ", "El ") & TipoDocumento & " Nº " & NumDocumento & " ya existe"
        End If
        rs.Filter = adFilterNone
        
        While j < intLineas And i < xTemp.Count(1)
            rs.AddNew
            rs!SEC_VENTA = intSecuencia
            rs!COD_CIA = objOFFUsuario.Cia
            rs!COD_LOCAL = objOFFUsuario.CodLocal
            rs!COD_TIPO_DOCUMENTO = TipoDocumento
            rs!NUM_DOCUMENTO = NumDocumento
            rs!COD_MAQUINA = objOFFUsuario.CodMaquina
            rs!USU_EMISION = objOFFUsuario.CodigoUsuario
            rs!FCH_EMISION = objOFFUsuario.sysdate
            rs!COD_ESTADO = COD_ESTADO_EMI
            rs!TIP_CLIENTE = objOFFCliente.Tipo
            rs!NUM_RUC = objOFFCliente.Ruc
            rs!DIR_CLIENTE = objOFFCliente.Direccion
            rs!NOM_CLIENTE = objOFFCliente.Nombre
            rs!NUM_ITEM = j + 1
            rs!COD_PRODUCTO = xTemp(i, 0)
            rs!DES_PRODUCTO = xTemp(i, 1)
            rs!CTD_PRODUCTO = xTemp(i, 2)
            rs!FLG_FRACCION = IIf(xTemp(i, 3) = "F", "1", "0")
            rs!PCT_DESCUENTO = xTemp(i, 4)
            rs!PRC_UNITARIO = xTemp(i, 5)
            rs!PRC_ORIGINAL = xTemp(i, 6)
            rs!MTO_IGV = xTemp(i, 7)
            rs!MTO_EXONERADO = xTemp(i, 8)
            rs!MTO_SUBTOTAL = xTemp(i, 9)
            rs!FLG_MODPRECIO = xTemp(i, 10)
            rs!FLG_REGALO = xTemp(i, 11)
            rs!PCT_IGV = xTemp(i, 12)
            rs!PARTIDA_ARANCELARIA = xTemp(i, 13)
            rs!CTD_FRACCIONA = xTemp(i, 14)
            rs!USU_MOD_PRECIO = xTemp(i, 15)
            rs!TIP_MOVIMIENTO = COD_TIP_MOV_VENTA
            rs!COD_SERIE_REL = objOFFUsuario.CodSerieEtiq
            rs.Update
            j = j + 1: i = i + 1
        Wend
        n = n + 1
        objDocumento.ActualizaCorrelativo TipoDocumento
    Wend
    Set objDocumento = Nothing
    Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Venta.GeneraDoc", Err.Description
End Sub

Private Function SecuenciaVta() As Integer
    Dim objArchivoIni As cls_ArchivoIni
    Dim strSec As String

    On Error GoTo CtrlErr
    Set objArchivoIni = New cls_ArchivoIni
    strSec = objArchivoIni.LeerIni(gstrIni, "general", "SEC_VENTA", "")
    Set objArchivoIni = Nothing
    
    If strSec = "" Or Not IsNumeric(strSec) Then
        SecuenciaVta = 0
    Else
        SecuenciaVta = Val(strSec)
    End If

    Exit Function

CtrlErr:
    SecuenciaVta = -1
End Function


Private Sub ActualizaSecuenciaVta(Optional ByVal pstrSec As String = "")
    Dim objArchivoIni As cls_ArchivoIni
    Dim strSec As String
    Dim intSec As Integer
    
    On Error GoTo CtrlErr

    Set objArchivoIni = New cls_ArchivoIni
    
    strSec = objArchivoIni.LeerIni(gstrIni, "general", "SEC_VENTA")
    
    If pstrSec = "" Or Not IsNumeric(pstrSec) Then
        If strSec = "" Or Not IsNumeric(strSec) Then
            intSec = 0
        Else
            intSec = Val(strSec) + 1
        End If
    Else
        intSec = Val(pstrSec)
    End If
    
    objArchivoIni.GuardarIni gstrIni, "general", "SEC_VENTA", Str(intSec)
    
    Set objArchivoIni = Nothing
    Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Venta.ActualizaSecuenciaVta", Err.Description
End Sub


Public Function AgregaInafectos(Codigo As String, _
                            Descripcion As String, _
                            Cantidad As Integer, _
                            FlagFraccion As String, _
                            PctDescuento As Double, _
                            PrcUnitario As Double, _
                            PrcOriginal As Double, _
                            MtoIgv As Double, _
                            MtoExonerado As Double, _
                            MtoSubTotal As Double, _
                            FlgModPrecio As String, _
                            FlgRegalo As String, _
                            PctIvg As Double, _
                            PartArancel As String, _
                            CtdFraccion As Integer, UsuModPrecio As String) As XArrayDB
    Dim ultimo As Integer
    Dim aux As Integer
    If xInaFectos.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xInaFectos.Count(1)
    While i < aux
        If xInaFectos(i, 0) = Codigo And xInaFectos(i, 11) = FlgRegalo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xInaFectos.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xInaFectos.AppendRows
    End If
j:
    If xInaFectos.Count(1) = 0 Then ultimo = 0: xInaFectos.AppendRows
    
    
    xInaFectos(ultimo, 0) = Codigo
    xInaFectos(ultimo, 1) = Descripcion
    xInaFectos(ultimo, 2) = Cantidad
    xInaFectos(ultimo, 3) = FlagFraccion
    xInaFectos(ultimo, 4) = PctDescuento
    xInaFectos(ultimo, 5) = PrcUnitario
    xInaFectos(ultimo, 6) = PrcOriginal
    xInaFectos(ultimo, 7) = MtoIgv
    xInaFectos(ultimo, 8) = MtoExonerado
    xInaFectos(ultimo, 9) = MtoSubTotal
    xInaFectos(ultimo, 10) = FlgModPrecio
    xInaFectos(ultimo, 11) = FlgRegalo
    xInaFectos(ultimo, 12) = PctIvg
    xInaFectos(ultimo, 13) = PartArancel
    xInaFectos(ultimo, 14) = CtdFraccion
    xInaFectos(ultimo, 15) = UsuModPrecio
    Set AgregaInafectos = xInaFectos
    
End Function

Public Function AgregaAfectos(Codigo As String, _
                            Descripcion As String, _
                            Cantidad As Integer, _
                            FlagFraccion As String, _
                            PctDescuento As Double, _
                            PrcUnitario As Double, _
                            PrcOriginal As Double, _
                            MtoIgv As Double, _
                            MtoExonerado As Double, _
                            MtoSubTotal As Double, _
                            FlgModPrecio As String, _
                            FlgRegalo As String, _
                            PctIgv As Double, _
                            PartArancel As String, _
                            CtdFraccion As Integer, UsuModPrecio As String) As XArrayDB
    Dim ultimo As Integer
    Dim aux As Integer
    If xAfectos.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xAfectos.Count(1)
    While i < aux
        If xAfectos(i, 0) = Codigo And xAfectos(i, 11) = FlgRegalo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xAfectos.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xAfectos.AppendRows
    End If
j:
    If xAfectos.Count(1) = 0 Then ultimo = 0: xAfectos.AppendRows
    
    
    xAfectos(ultimo, 0) = Codigo
    xAfectos(ultimo, 1) = Descripcion
    xAfectos(ultimo, 2) = Cantidad
    xAfectos(ultimo, 3) = FlagFraccion
    xAfectos(ultimo, 4) = PctDescuento
    xAfectos(ultimo, 5) = PrcUnitario
    xAfectos(ultimo, 6) = PrcOriginal
    xAfectos(ultimo, 7) = MtoIgv
    xAfectos(ultimo, 8) = MtoExonerado
    xAfectos(ultimo, 9) = MtoSubTotal
    xAfectos(ultimo, 10) = FlgModPrecio
    xAfectos(ultimo, 11) = FlgRegalo
    xAfectos(ultimo, 12) = PctIgv
    xAfectos(ultimo, 13) = PartArancel
    xAfectos(ultimo, 14) = CtdFraccion
    xAfectos(ultimo, 15) = UsuModPrecio
    Set AgregaAfectos = xAfectos
    
End Function




Public Sub Totales()
Dim dblTotalIgv As Double
Dim dblTotalExonerado As Double
Dim dblTotalVenta As Double
Dim dblTotalPago As Double
Dim dblTotalVuelto As Double
Dim xTemp As New XArrayDB

Dim i As Integer

    On Error GoTo CtrlErr

    dblTotalIgv = 0
    dblTotalExonerado = 0
    dblTotalVenta = 0

    
    For i = 0 To DetalleVenta.UpperBound(1)
        dblTotalVenta = dblTotalVenta + DetalleVenta(i, ColDVMtoSubtotal)
        dblTotalIgv = dblTotalIgv + DetalleVenta(i, ColDVMtoIgv)
        dblTotalExonerado = dblTotalExonerado + DetalleVenta(i, ColDVMtoExonerado)
    Next i
    TotalVenta = dblTotalVenta
    TotalIgv = dblTotalIgv
    TotalExonerado = dblTotalExonerado
    
    
    dblTotalPago = 0
    dblTotalVuelto = 0
    
    
    Set xTemp = PagoVenta

    For i = 0 To xTemp.UpperBound(1)
        AgregaPagoVenta xTemp(i, ColDPCodPago), _
                        xTemp(i, ColDPCodMoneda), _
                        xTemp(i, ColDPMtoImporte), _
                        xTemp(i, ColDPMtoSoles), _
                        xTemp(i, ColDPNumTarjeta), _
                        xTemp(i, ColDPFchVencimiento), _
                        xTemp(i, ColDPNumCuotas), _
                        xTemp(i, ColDPTipoCambio), _
                        xTemp(i, ColDPTipoCuota), _
                        xTemp(i, ColDPNumAutorizacion)
    Next i
    
    For i = 0 To PagoVenta.UpperBound(1)
        dblTotalPago = dblTotalPago + PagoVenta(i, ColDPMtoSoles)
        dblTotalVuelto = dblTotalVuelto + PagoVenta(i, ColDPMtoVuelto)
        
    Next i

    If dblTotalPago = 0 Then
        dblTotalVuelto = 0
    Else
        dblTotalVuelto = dblTotalPago - dblTotalVenta
    End If

    TotalPago = dblTotalPago
    TotalVuelto = dblTotalVuelto
    
    If objOFFUsuario.TipoCambio > 0 Then
        TotalDolares = Round(dblTotalVenta / Val(objOFFUsuario.TipoCambio), 2)
    End If

Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Venta.Totales", Err.Description

End Sub



Public Function Existe(sArchivo As String) As Integer
    Existe = Len(Dir$(sArchivo))
End Function

Public Property Get CodDocDefault() As String
    CodDocDefault = strCodDocDefault
End Property

Public Property Let CodDocDefault(ByVal lstrCodDocDefault As String)
    strCodDocDefault = lstrCodDocDefault
End Property

Public Property Get NumDocDefault() As String
    NumDocDefault = strNumDocDefault
End Property

Public Property Let NumDocDefault(ByVal lstrNumDocDefault As String)
    strNumDocDefault = lstrNumDocDefault
End Property

Public Property Get TotalDolares() As Double
    TotalDolares = dblTotalDolares
End Property

Public Property Let TotalDolares(ByVal ldblTotalDolares As Double)
    dblTotalDolares = ldblTotalDolares
End Property

Public Property Get DocumentosGenerados() As XArrayDB
    Set DocumentosGenerados = xDocumentosGenerados
End Property


Public Function AgregaDocumentosGenerados(TipoDocumento As String, _
                            NumDocumento As String) As XArrayDB
    Dim ultimo As Integer
    Dim aux As Integer
    If xDocumentosGenerados.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xDocumentosGenerados.Count(1)
    While i < aux
        If xDocumentosGenerados(i, 0) = TipoDocumento And xDocumentosGenerados(i, 1) = NumDocumento Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xDocumentosGenerados.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xDocumentosGenerados.AppendRows
    End If
j:
    If xDocumentosGenerados.Count(1) = 0 Then ultimo = 0: xDocumentosGenerados.AppendRows
    
    
    xDocumentosGenerados(ultimo, 0) = TipoDocumento
    xDocumentosGenerados(ultimo, 1) = NumDocumento
    Set AgregaDocumentosGenerados = xDocumentosGenerados
    
End Function




