VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim intItems As String
Dim objLocal As New clsLocal
Dim objWS As New clsWebService
Dim objProducto As New clsProducto
Dim detalleProductoError As String
'Dim strNumPedidoPadre As String
'Dim strCodLocalRef As String
'*******************************************************************************
'Este funcion lista todos los pedidos de delivery para este local de la empresa
'*******************************************************************************
Public Function ListaPedidoDLV(ByVal vstrCia As String, _
                               ByVal vstrCodLocal As String, _
                               ByVal vstrFchIni As String, _
                               ByVal vstrFchFin As String, _
                               ByVal vstrCodEstado As String) As oraDynaset

    On Error GoTo CtrlErr
    Set ListaPedidoDLV = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.LISTA_PEDIDO_DLV", 0, vstrCia, vstrCodLocal, vstrFchIni, vstrFchFin, vstrCodEstado)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDelivery", Err.Description
End Function

'*******************************************************************************
'sin documentar
'*******************************************************************************
Public Function ListaStock(ByVal Cia As String, NumeroProforma As String, ByVal LocalStock As String, Optional LocalProforma As String) As oraDynaset
    Set ListaStock = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_STOCK", 0, Cia, NumeroProforma, LocalStock, LocalProforma)
End Function

'*******************************************************************************
'sin documentar
'*******************************************************************************
Public Function ListaEstado(ByVal Estado As String) As oraDynaset
    Set ListaEstado = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_ESTADO_P", 0, Estado)
End Function

'*******************************************************************************
'sin documentar
'*******************************************************************************
Public Function ListaEstados(ByVal Cia As String, ByVal NumeroProforma As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaEstados = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_ESTADO", 0, Cia, NumeroProforma)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function

'*******************************************************************************
'sin documentar
'*******************************************************************************
Public Function ListaFormaPago(ByVal Cia As String, ByVal CodigoLocal As String, NumeroProforma As String) As oraDynaset
    Set ListaFormaPago = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_FORMA_PAGO", 0, Cia, CodigoLocal, NumeroProforma)
End Function
'I.ECASTILLO 27.10.2020
Public Function ListaFormaPago_V2(ByVal Cia As String, ByVal CodigoLocal As String, NumeroProforma As String) As oraDynaset
    Set ListaFormaPago_V2 = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_FORMA_PAGO_V2", 0, Cia, CodigoLocal, NumeroProforma)
End Function
Public Function ListaFormaPagoVale(ByVal Cia As String, ByVal CodigoLocal As String, NumeroProforma As String) As oraDynaset
    Set ListaFormaPagoVale = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_FORMA_PAGO_VALE", 0, Cia, CodigoLocal, NumeroProforma)
End Function
'F.ECASTILLO 27.10.2020
'I.ECASTILLO 17.12.2020
Public Function ListaCabConv(ByVal Cia As String, ByVal CodigoLocal As String, NumeroProforma As String) As oraDynaset
    Set ListaCabConv = gclsOracle.FN_Cursor("DELIVERY.PKG_DELIVERY_CONVENIO.F_CUR_GET_CAB_CONV", 0, Cia, CodigoLocal, NumeroProforma)
End Function
Public Function ListaDetConv(ByVal Cia As String, ByVal CodigoCliente As String, ByVal CodigoLocal As String, NumeroProforma As String) As oraDynaset
    Set ListaDetConv = gclsOracle.FN_Cursor("DELIVERY.PKG_DELIVERY_CONVENIO.F_CUR_GET_DET_CONV", 0, Cia, CodigoCliente, NumeroProforma, CodigoLocal)
End Function
'F.ECASTILLO 17.12.2020
'*******************************************************************************
'sin documentar
'*******************************************************************************
Public Function Lista_x_Estado(ByVal Cia As String, _
                               ByVal Estado As String, _
                               ByVal CodigoLocal As String, _
                               ByVal FechaInicio As String, _
                               ByVal FechaFin As String, _
                               Optional ByVal Telefono As String = "", _
                               Optional ByVal Proforma As String = "", _
                               Optional ByVal Direccion As String = "", _
                               Optional ByVal Nombre As String = "", _
                               Optional ByVal CodigoLocalDes As String = "", _
                               Optional ByVal CodigoMotorizado As String = "", _
                               Optional ByVal Documento As String = "", _
                               Optional ByVal TipoDoc As String = "", _
                               Optional ByVal usuario As String = "", Optional ByVal CodigoPais As String = "", Optional ByVal CodigoDepartamento As String = "", Optional ByVal CodigoProvincia As String = "", Optional ByVal CodigoDistrito As String = "", Optional ByVal CodigoUrbanizacion As String = "", Optional ByVal CodigoRuteador As String = "", Optional ByVal EstPendiente As String = "0", Optional ByVal ConexionSalud As String, Optional ByVal strModalidad As String = "", Optional ByVal strCiaRef As String = "") As oraDynaset
                               
                               

    Set Lista_x_Estado = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_X_ESTADO", 0, Cia, _
                                                                                           Estado, _
                                                                                         CodigoLocal, _
                                                                                           FechaInicio, _
                                                                                           FechaFin, _
                                                                                           Telefono, _
                                                                                           Proforma, _
                                                                                           Direccion, _
                                                                                           Nombre, _
                                                                                           CodigoLocalDes, _
                                                                                           CodigoMotorizado, _
                                                                                           Documento, _
                                                                                           TipoDoc, _
                                                                                           usuario, CodigoPais, CodigoDepartamento, CodigoProvincia, CodigoDistrito, CodigoUrbanizacion, CodigoRuteador, EstPendiente, ConexionSalud, strModalidad, strCiaRef)
End Function

Public Function ListaDelivery(ByVal Cia As String, ByVal Estado As String, ByVal CodigoLocal As String, ByVal CodigoZona As String, ByVal FchInicio As String, ByVal FchFin As String) As oraDynaset
    On Error GoTo CntrlError
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "0" Then
        Set ListaDelivery = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DELIVERY_NEW", 0, Cia, Estado, CodigoLocal, CodigoZona, FchInicio, FchFin)
    Else
        'I.ECASTILLO 17.12.2020
        'If Estado = "011" Then
            Set ListaDelivery = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DELIVERY_V2", 0, Cia, Estado, CodigoLocal, CodigoZona, FchInicio, FchFin)
        'Else
        '    Set ListaDelivery = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DELIVERY_NEW", 0, Cia, Estado, CodigoLocal, CodigoZona, FchInicio, FchFin)
        'End If
        'F.ECASTILLO 17.12.2020
    End If
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function ListaTransferencia(ByVal Cia As String, ByVal CodigoLocal As String, ByVal NumeroProforma As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaTransferencia = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_TRANSFERENCIA", 0, Cia, CodigoLocal, NumeroProforma)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function ListaHistoricoCliente(ByVal vstrCia As String, _
                                      ByVal vstrCod_Local As String, _
                                      ByVal vstrCod_Cliente As String, _
                                      ByVal vstrNum_Meses As Integer) As oraDynaset
    On Error GoTo CntrlError
    Set ListaHistoricoCliente = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_HIST_CLIENTE", 0, vstrCia, vstrCod_Local, vstrCod_Cliente, vstrNum_Meses)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function ListaHistoricoFPC(ByVal vstrCia As String, _
                                  ByVal vstrCod_Local As String, _
                                  ByVal vstrCod_Cliente As String, _
                                  ByVal vstrNum_Meses As Integer) As oraDynaset
    On Error GoTo CntrlError
    Set ListaHistoricoFPC = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_HIST_FPCLI", 0, vstrCia, vstrCod_Local, vstrCod_Cliente, vstrNum_Meses)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function NumeroMesesHistorico(ByVal vstrCia As String) As Integer
    On Error GoTo CntrlError
    NumeroMesesHistorico = gclsOracle.FN_Valor("BTLPROD.PKG_PROFORMA_V3.FN_NUM_MESES_HIST", vstrCia)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function

'**** Graba Proforma Creado Por -- Crueda ****'
Public Function Graba(ByVal vstrLocal As String, Optional vstrProforma As String, _
                      Optional vstrProd As String, Optional vstrCtdProd As String, _
                      Optional vstrCtdProdFrac As String, Optional vstrCtdFraccionamiento As String, _
                      Optional vstrPctDstUni As String, Optional vstrPrcUnitKairo As String, _
                      Optional vstrImpUniBase As String, Optional vstrImpUniImpuesto As String, _
                      Optional vstrImpComision As String, Optional vstrMtoSTotal As String, _
                      Optional vstrPromocion As String, Optional vstrProdBtl As String, _
                      Optional vstrProductoRM As String, Optional vstrCtdProdRM As String, _
                      Optional vstrCtdBaseRM As String, Optional vstrPctBaseRM As String, _
                      Optional vstrImpPrecioRM As String, Optional vstrUsuario As String, _
                      Optional vstrCliente As String, Optional vstrMaquina As String, _
                      Optional vstrMedico As String) As String
Dim gvarValores  As Variant
Dim gvarIO  As Variant
                On Error GoTo CntrlError
                    gvarValores = Array(vstrLocal, "", _
                                        vstrProd, vstrCtdProd, _
                                        vstrCtdProdFrac, vstrCtdFraccionamiento, _
                                        vstrPctDstUni, vstrPrcUnitKairo, _
                                        vstrImpUniBase, vstrImpUniImpuesto, _
                                        vstrImpComision, vstrMtoSTotal, _
                                        vstrPromocion, vstrProdBtl, _
                                        vstrProductoRM, vstrCtdProdRM, _
                                        vstrCtdBaseRM, vstrPctBaseRM, _
                                        vstrImpPrecioRM, vstrUsuario, _
                                        vstrCliente, vstrMaquina, _
                                        vstrMedico)
                                                    
                    gvarIO = Array(entrada, Salida, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada, entrada, _
                                   entrada)
                                   
                    Graba = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_GRABA", gvarValores, gvarIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function GrabaTransferencia(ByVal Cia As String, _
                      ByVal CodigoLocal As String, _
                      ByVal NumeroProforma As String, _
                      ByVal codigoProducto As String, _
                      ByVal CantidadProducto As String, _
                      ByVal CantidadProductoFrac As String, _
                      ByVal LocalOrigen As String, _
                      ByVal LocalDestino As String, _
                      ByVal usuario As String) As String
Dim varValores As Variant
Dim varIO As Variant
                On Error GoTo CntrlError
                    varValores = Array(Cia, _
                      CodigoLocal, _
                      NumeroProforma, _
                      codigoProducto, _
                      CantidadProducto, _
                      CantidadProductoFrac, _
                      LocalOrigen, _
                      LocalDestino, _
                      usuario)
                                                    
                    varIO = Array( _
                        entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada)
                                   
                    GrabaTransferencia = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_GRABA_TRANSFERENCIA", varValores, varIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function Anula(ByVal Cia As String, _
                      ByVal CodigoLocal As String, _
                      ByVal NumeroProforma As String, _
                      ByVal usuario As String _
) As String
Dim varValores As Variant
Dim varIO As Variant
                On Error GoTo CntrlError
                    varValores = Array(Cia, _
                      CodigoLocal, _
                      NumeroProforma, _
                      usuario)
                                                    
                    varIO = Array( _
                        entrada, _
                      entrada, _
                      entrada, _
                      entrada)
                                   
    Anula = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_ANULA", varValores, varIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function Asigna(ByVal Cia As String, _
                      ByVal CodigoLocal As String, _
                      ByVal NumeroProforma As String, _
                      ByVal usuario As String, _
                      ByVal CodigoMotorizado As String, _
                      ByVal CodigoMaquina As String, _
                      ByVal Observaciones As String, _
                      ByVal CodLocalReferencia As String, _
                      ByVal CodCliente As String, _
                      ByVal CodDireccionCli As String, _
                      ByVal Referencia As String, _
                      Optional strAsigna As String, _
                      Optional obsLocal As String, _
                      Optional AsignasinValidar As String = "0" _
                      , Optional estConfig As String = "" _
                    ) As String
Dim varValores As Variant
Dim varIO As Variant
                On Error GoTo CntrlError
                    varValores = Array(Cia, _
                      CodigoLocal, _
                      NumeroProforma, _
                      usuario, _
                      CodigoMotorizado, _
                      CodigoMaquina, _
                      Observaciones, _
                      CodLocalReferencia, _
                      CodCliente, _
                      CodDireccionCli, _
                      Referencia, _
                      strAsigna, _
                      objUsuario.CodTipoVenta, _
                      obsLocal, _
                      AsignasinValidar, _
                      estConfig)
                                                    
                    varIO = Array( _
                        entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada)
                                   
    Asigna = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_GRABA_RUTEO", varValores, varIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function AsignaLocal(ByVal Cia As String, _
                      ByVal CodigoLocal As String, _
                      ByVal NumeroPedido As String, _
                      ByVal CodigoLocalNuevo As String) As String
Dim gvarValores  As Variant
Dim gvarIO  As Variant
                On Error GoTo CntrlError
                    gvarValores = Array(Cia, CodigoLocal, NumeroPedido, CodigoLocalNuevo)
                                                    
                    gvarIO = Array(entrada, entrada, entrada, entrada)
                                   
                    AsignaLocal = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_GRABA_NVO_LOCAL_DESPACHO", gvarValores, gvarIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function CambiaEstado(ByVal Cia As String, _
                      ByVal CodigoLocal As String, _
                      ByVal NumeroPedido As String, _
                      ByVal CodigoEstado As String, _
                      ByVal CodigoUsuario As String) As String
Dim gvarValores  As Variant
Dim gvarIO  As Variant
                On Error GoTo CntrlError
                    gvarValores = Array(Cia, CodigoLocal, NumeroPedido, CodigoEstado, CodigoUsuario, "1")
                                                    
                    gvarIO = Array(entrada, entrada, entrada, entrada, entrada, entrada)
                                   
                    CambiaEstado = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_GRABA_SEGUIM_PEDIDO", gvarValores, gvarIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function


'''''''''''--============================================================================================--'
'''''''''''--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@--@@----'
'''''''''''-- I N P R E S I O N      DE      P R O F O R M A S
'''''''''''--============================================================================================--'
'''''''''''-- Fecha => 13/04/2007
''''''''''
''''''''''Public Function Imp_CabProf(ByVal vstrCia As String, _
''''''''''                             ByVal vstrCodLocal As String, _
''''''''''                             ByVal vstrNumProf As String) As OraDynaset
''''''''''    On Error GoTo CtrlErr
''''''''''    Set Imp_CabProf = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_IMP_CAB_PROF", 0, vstrCia, vstrCodLocal, vstrNumProf)
''''''''''
''''''''''    Exit Function
''''''''''CtrlErr:
''''''''''    Err.Raise Err.Number, "clsProforma.Imp_CabProf", Err.Description
''''''''''End Function
''''''''''
''''''''''Public Function Imp_DetProf(ByVal vstrCia As String, _
''''''''''                            ByVal vstrCodLocal As String, _
''''''''''                            ByVal vstrNumProf As String) As OraDynaset
''''''''''    On Error GoTo CtrlErr
''''''''''    Set Imp_DetProf = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_IMP_DET_PROF", 0, vstrCia, vstrCodLocal, vstrNumProf)
''''''''''
''''''''''    Exit Function
''''''''''CtrlErr:
''''''''''    Err.Raise Err.Number, "clsProforma.Imp_DetProf", Err.Description
''''''''''End Function

Public Function Imprimir(ByVal vstrCia As String, _
                                ByVal vstrCodLocal As String, _
                                ByVal vstrNumProf As String)
    
    Dim intRespuesta%
    Dim odyCabProf As oraDynaset
    Dim odyDetProf As oraDynaset

    On Error GoTo ErrorImpresora

    Set odyCabProf = Me.ListaCabecera(vstrCia, vstrCodLocal, vstrNumProf)
    Set odyDetProf = Me.ListaDetalle(vstrCia, vstrCodLocal, vstrNumProf)

    'Printer.Width = 220 * 56.7
    'Printer.Height = 170 * 56.7
    Printer.Print
    'Printer.FontName = "Draft 17cpi"
    Printer.Print Space(1) & "Proforma Nº" & "  " & Trim(vstrNumProf) & Space(10) & "Fecha:" & " " & Trim(odyCabProf("FCH_REGISTRA").Value) & Space(25) & "Proforma Nº" & "  " & Trim(vstrNumProf) & Space(10) & "Fecha:" & " " & Trim(odyCabProf("FCH_REGISTRA").Value)
    Printer.Print Space(1) & "Ruc:" & " " & Trim(odyCabProf("RUC").Value) & Space(65) & "E X I J A   S U   B O L E T A   D E   V E N T A"
    Printer.Print Space(1) & "Nombre : " & IIf(IsNull(odyCabProf("DES_AUX_CLI_NOMBRE").Value), "", odyCabProf("DES_AUX_CLI_NOMBRE").Value)
    Printer.Print Space(1) & "Dirección : " & IIf(IsNull(odyCabProf("DES_AUX_CLI_DIRECC").Value), "", odyCabProf("DES_AUX_CLI_DIRECC").Value)
    
    Dim dblTotal As Double
    dblTotal = 0
    Printer.Print Space(1) & left("" & Space(2), 2) & Space(1), _
                             left("" & Space(5), 5) & Space(1), _
                             left("" & Space(25), 25) & Space(1), _
                             left("Labo." & Space(4), 4) & Space(1), _
                             right(Space(5) & "Cant", 5) & Space(1), _
                             right(Space(8) & "Val.Neto", 8) & Space(1), _
                             left("Efectivo Soles" & Space(15), 15)

    odyDetProf.MoveFirst
    While Not odyDetProf.EOF
        Printer.Print Space(1) & left(Trim(odyDetProf("ITEM").Value) & Space(2), 2), _
                                 left(odyDetProf("COD_PRODUCTO").Value & Space(5), 5), _
                                 left(odyDetProf("DES_PRODUCTO").Value & Space(25), 25), _
                                 left(odyDetProf("LABORATORIO").Value & Space(4), 4), _
                                 right(Space(5) & odyDetProf("CTD_PRODUCTO").Value, 5), _
                                 right(Space(8) & odyDetProf("MTO_SUBTOTAL").Value, 8)

                      dblTotal = dblTotal + odyDetProf("MTO_SUBTOTAL").Value

        odyDetProf.MoveNext
    Wend
    Printer.Print
    Printer.Print
    Printer.Print Space(82) & "E X I J A   S U   B O L E T A   D E   V E N T A"
    Printer.Print Space(1) & "Boticas BTL Boticas" & Space(10) & Trim(odyCabProf("USUARIO").Value) & Space(25) & "Boticas BTL Boticas" & Space(10) & Trim(odyCabProf("USUARIO").Value)
    Printer.Print Space(1) & "Descuentos de Siempre" & Space(9) & "Total a Pagar: S/." & " " & Format(dblTotal, "#,###,##0.00") & Space(25) & "Descuentos de Siempre" & Space(9) & "Total a Pagar: S/." & " " & Format(dblTotal, "#,###,##0.00")
    Printer.EndDoc

    Exit Function
ErrorImpresora:
   intRespuesta = MsgBox("Existe un problema con la impresora" + Chr(13) + "Desea esperar a ser resuelto?", vbYesNo)
   If intRespuesta = vbYes Then
      Resume
   Else
      Exit Function
   End If
End Function

Public Sub ImprimirTicket_Proforma(ByVal vstrCia As String, _
                                    ByVal vstrLocal As String, _
                                    ByVal vstrDocumento As String)

   Dim intRespuesta%
   Dim rstCabecera As oraDynaset
   Dim rstDetalle As oraDynaset
   Dim objDocumento As New clsDocumento
   Dim strMensaje As String
   Dim p As Printer
   Dim vInd As Integer
   Dim i As Integer
   Dim j As Integer

   Set rstCabecera = ListaCabecera(vstrCia, vstrLocal, vstrDocumento)
   Set rstDetalle = ListaDetalle(vstrCia, vstrLocal, vstrDocumento)

    vInd = 0

    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          vInd = "1"
          Exit For
       End If
    Next p

   On Error GoTo ErrorImpresora
            
            '***********************************************************************'

            Open "lpt1" For Output As #1
            '''Open "COM3" For Output As #1
            '**** 16-DOT PITCH FONT *******'
            Print #1, Chr$(27); Chr$(58);
            '-------------------------------'
            Print #1, "    BOTICAS BTL    "
            '****** Standar ********'
            Print #1, Chr$(27); Chr$(77);
            '-----------------------------'
            Print #1, "   BOTICA TORRES DE LIMATAMBO SAC    "
            Print #1, Space(1) & objUsuario.Direccion
            Print #1, Space(8) & "RUC :" & objUsuario.Ruc
            Print #1, Chr$(27); Chr$(58);
            '-------------------------------'
            Print #1, "****** PROFORMA ******"
            '****** Standar ********'
            Print #1, Chr$(27); Chr$(77);
            '-----------------------------'

            Print #1, "BTL" & objUsuario.CodigoLocal & Space(2) & CadenaDirecc_Part1(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
            Print #1, Space(8) & CadenaDirecc_Part2(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
            Print #1, "FECHA:" & Format(rstCabecera("FECHA").Value, "DD/MM/YYYY HH:MM:SS") & Space(1) & "CAJA:" & " " & objUsuario.NombrePC
            
            Print #1, "Proforma" & " " & "Nº" & Space(1) & rstCabecera("NUM_PROFORMA").Value & " " & "Serie:"

            Print #1, ""

            Print #1, left("Codigo" & Space(5), 5) & Space(1) & _
                      left("Descripcion" & Space(20), 20) & Space(1) & _
                      right(Space(5) & "Cant.", 5) & Space(1) & _
                      right(Space(8) & "Importe", 8) & Space(1)
            Print #1, String(42, "-")

            Printer.FontBold = False
            While Not rstDetalle.EOF

               If Len(rstDetalle("DES_PRODUCTO").Value) > 20 Then ' SI LA DESCRIPCION DE PRODUCTO ES MAYOR A 20 CARACTERES'
                        Print #1, left(rstDetalle("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                  left(Mid(rstDetalle("DES_PRODUCTO").Value, 1, 20) & Space(20), 20) & Space(1) & _
                                  right(Space(5) & rstDetalle("CANT").Value, 5) & Space(2) & _
                                  right(Space(7) & rstDetalle("MTO_SUBTOTAL").Value, 7) & Space(1)
                        Print #1, Space(5) & left(Mid(rstDetalle("DES_PRODUCTO").Value, 21, 15) & Space(1) & _
                                  left(rstDetalle("LABORATORIO").Value, 4) & Space(20), 20) & Space(3)
                  Else
                        Print #1, left(rstDetalle("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                  left(Mid(rstDetalle("DES_PRODUCTO").Value, 1, 20) & Space(1) & left(rstDetalle("LABORATORIO").Value, 4), 20) & Space(1) & _
                                  Space(30) & right(Space(5) & rstDetalle("CANT").Value, 5) & Space(1) & _
                                  right(Space(7) & rstDetalle("MTO_SUBTOTAL").Value, 7) & Space(1)
                End If

                rstDetalle.MoveNext
            Wend

            Print #1, Space(15) & "Total         S/. " & " " & Format(rstCabecera("MTO_TOTAL").Value, "#,###,##0.00")
            Print #1, ""
                  
               Print #1, "Cajero:" & Space(1) & rstCabecera("CAJERA").Value
               Print #1, String(42, "-")
               Print #1, "************* www.btl.com.pe *************" & _
                         "           Gracias por su compra          " & _
                          Space(12) & objVenta.Apostrofe & objVenta.SloganBTL & objVenta.Apostrofe
               Print #1, "       DELIVERY 612-5000 LAS 24 HORAS   "
            
'            If Val(TPagTkt) > 1 Then
'               Print #1,
'               Print #1, "Ticket" & Space(1) & Val(PagTkt) & "/" & Val(TPagTkt)
'            End If
             
            ''AQUI VA EL MENSAJE
            
            strMensaje = objDocumento.DevMensaje(objUsuario.CodigoEmpresa, vstrDocumento, "PRO", objUsuario.CodigoLocal)
            If Len(strMensaje) < 40 Then
                    Print #1, strMensaje
              Else
                For i = 0 To val(Len(strMensaje)) / 40
                    Print #1, strMensaje
                Next i
            End If
            
            For j = 1 To 5
               Print #1, ""
            Next j
      
            'Copia = Copia + 1
            '** Corte Parcial **'
            Print #1, Chr(27) + "d" + "1"
            '** Corte Total **'
            'Print #1, Chr$(27); "d"; Chr$(2)

            '*********************************************************************************************************
            Close #1
                  
    On Error GoTo 0
    Exit Sub

ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
          Resume
       Else
          Exit Sub
     End If

End Sub

Public Sub ImprimirDelivery(ByVal vstrCia As String, _
                    ByVal vstrCodBtl As String, _
                    ByVal vstrProforma As String, _
                    ByVal vintNumCopia As String)
                            
   Dim dblSumaSubTotal As Double
   Dim intLinea As Integer
   Dim intCopia As Integer
   Dim i As Integer
   Dim objImpresiones As New clsProforma
   Dim odynCab As oraDynaset
   Dim odynDet As oraDynaset
   Dim rsFormaPago As oraDynaset
   Dim rsDocxProf As oraDynaset
   Dim TotItems As String
   Dim dblSumaTotalPagar As Double
   Dim dblSumaTotalVuelto As Double
   
   Set odynCab = objImpresiones.ListaCabecera(vstrCia, vstrCodBtl, vstrProforma)
   Set odynDet = objImpresiones.ListaDetalle(vstrCia, vstrCodBtl, vstrProforma)
   Set rsFormaPago = objImpresiones.ListaFormaPago(vstrCia, vstrCodBtl, vstrProforma)
   Set rsDocxProf = objImpresiones.Lista_DocsxProf(vstrCia, vstrCodBtl, vstrProforma)
                
    'Frm_SIS_Impresoras.Show vbModal
    Printer.PaperSize = vbPRPSA4
    
    TotItems = "0" & odynCab("ITEM").Value
    
    intCopia = 0: TotItems = 0: dblSumaSubTotal = 0: dblSumaTotalPagar = 0: dblSumaTotalVuelto = 0
    If vintNumCopia = 0 Then Exit Sub
    
    '-- Generar Copia --'
    Do While vintNumCopia > intCopia
        odynCab.MoveFirst
        odynDet.MoveFirst
        While Not odynDet.EOF
            'Cabecera
        
            On Error GoTo Mal_Seteo
            GoTo OK_Seteo
Mal_Seteo:
            MsgBox Err.Description, vbExclamation, "Error de Impresión"
            Exit Sub
OK_Seteo:
            On Error GoTo 0
            On Error GoTo ErrorImpresora
 
            Printer.FontName = "Draft 15cpi"
            Printer.FontSize = 9
            Printer.Print Space(43) & "Fecha :" & " " & odynCab("FECHA").Value
            Printer.Print Space(50) & "Hora  :" & " " & odynCab("HORA").Value
            Printer.Print Space(20) & "Hora Avisado :" & " " & odynCab("AVISADO").Value & Space(7) & "Hora Proformado :" & " " & odynCab("PROFORMADO").Value
            Printer.Print
            Printer.FontName = "Draft 15cpi"
            
            Printer.FontSize = 15
            If odynCab("FLG_ENTREGA_TERCERO").Value = "0" Then
                Printer.Print pfstr_Alineado("PEDIDO DELIVERY", 50, centro)
              Else
                Printer.Print pfstr_Alineado("PEDIDO DELIVERY", 50, centro) & Space(10) & pfstr_Alineado("CONEXION SALUD", 50, centro)
            End If
            Printer.FontName = "Draft 17cpi"
            Printer.FontSize = 9
            
            Printer.Print
            Printer.FontBold = True
            Printer.Print Space(2) & "Atendido        : " & pfstr_Alineado(odynCab("CAJERA").Value, 50) & " " & "Motorizado:" & " " & pfstr_Alineado(odynCab("MOTORIZADO").Value, 30)
            Printer.FontBold = False
            Printer.Print Space(2) & "Nro.Proforma    : " & pfstr_Alineado(odynCab("NUM_PROFORMA").Value, 50)
            'Printer.Print Space(2) & "Codigo Cliente  : " & pfstr_Alineado(odynCab("COD_CLIENTE").Value, 50)
            'Printer.Print Space(2) & "Telefono        : " & pfstr_Alineado("xxxxxxxx", 15) & Space(35) '& "Tarjeta" & " " & pfstr_Alineado(IIf(IsNull(odynCab("NUM_TARJETA").Value), "", odynCab("NUM_TARJETA").Value), 20) & " " & "F.V" & " " & pfstr_Alineado(IIf(IsNull(odynCab("FCH_VENCIMIENTO").Value), "", odynCab("FCH_VENCIMIENTO").Value), 15)
            Printer.Print Space(2) & "Nombre          : " & pfstr_Alineado(odynCab("NOMBRES").Value, 85) & " " & "D.N.I." & " " & pfstr_Alineado(IIf(IsNull(odynCab("NUM_DOCUMENTO_ID").Value), "", odynCab("NUM_DOCUMENTO_ID").Value), 13)
'            Printer.Print Space(2) & "Dirección       : " & pfstr_Alineado(odynCab("DIRECCION").Value, 80)
            Printer.FontBold = True
            Printer.Print Space(2) & "Dirección       : " & pfstr_Alineado(IIf(IsNull(odynCab("DES_SUFIJO").Value), "", odynCab("DES_SUFIJO").Value), 2) & " " & pfstr_Alineado(IIf(IsNull(odynCab("DIRECCION").Value), "", odynCab("DIRECCION").Value), 80)
            Printer.FontBold = False
            Printer.Print Space(2) & "Distrito        : " & pfstr_Alineado(odynCab("DISTRITO").Value, 50) & " " & "T.C" & " " & pfstr_Alineado(odynCab("IMP_TIP_CAMBIO").Value, 8) & Space(15) '& "Forma Pago" & " " & pfstr_Alineado(odynCab("DES_HIJO").Value, 15)
            Printer.Print Space(2) & "Urbanización    : " & pfstr_Alineado(odynCab("URBANIZACION").Value, 50) & " " & "Fecha Pactada" & " " & pfstr_Alineado(odynCab("FCH_HORA_PACT_ENTR").Value, 22) & Space(15)
            Printer.Print Space(2) & "Referencia      : " & pfstr_Alineado(odynCab("REFERENCIAS").Value, 80)
            'Printer.Print Space(2) & "Observación Verificación: " & pfstr_Alineado(odynCab("OBS_NOTA_VERIFICACION").Value, 91)
            'Printer.Print Space(2) & "Observación Ruteo : " & pfstr_Alineado(odynCab("OBS_NOTA_RUTEO").Value, 91)
            Printer.Print Space(2) & "Observación Local : " & pfstr_Alineado(odynCab("OBS_NOTA_LOCAL").Value, 91)
            Printer.Font.Bold = True
            Printer.Print Space(2) & "Observación Motorizado : " & pfstr_Alineado(odynCab("OBS_NOTA_MOTORIZADO").Value, 91)
            Printer.FontBold = False
            
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx '
            If rsDocxProf.RecordCount > 0 Then
                    Dim c%
                    Dim strDoc1 As String
                    Dim strDoc2 As String
                    Printer.Print Space(71) & odynCab("DES_MODALIDAD").Value & " : "
                    rsDocxProf.MoveFirst
                    c = 0
                    While Not rsDocxProf.EOF
                        c = c + 1
                        If c = 1 Then
                           strDoc1 = "" & rsDocxProf("TIPODOC").Value
                        ElseIf c = 2 Then
                           strDoc2 = "" & rsDocxProf("TIPODOC").Value
                        End If
                        rsDocxProf.MoveNext
                    Wend
                    Printer.FontName = "Roman 15cpi"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    
                    Printer.Print Space(88) & strDoc1
                    Printer.Print Space(88) & strDoc2
                    
                    Printer.FontName = "Draft 17cpi"
                    Printer.FontSize = 8
                    Printer.FontBold = False
            End If
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx '
            If IIf(IsNull(odynCab("COD_CONVENIO")), "", odynCab("COD_CONVENIO")) <> "" Then
                Dim strConv As String
                Dim intCant As Byte
                Dim rsDocumentoConv As oraDynaset
                Dim objConvenio1 As New clsConvenio
                
                Set rsDocumentoConv = objConvenio1.ListaDocVerif(odynCab("COD_CONVENIO"))
                
                Printer.Print
                
                If Not rsDocumentoConv.EOF Then
                    Printer.Print Space(53) & "Verificación de" & Space(2) & String(50, "_")
                    Printer.Print Space(2) & "CONVENIO" & Space(47) & "Documentos" & Space(3) & _
                                          pfstr_Alineado("Descripción", 30, Izquierda) & Space(9) & _
                                          pfstr_Alineado("Retención", 10, Izquierda)
                    Printer.Print Space(70) & String(50, "_")
                Else
                    Printer.Print Space(2) & "CONVENIO"
                End If
                
                intCant = 0
                rsDocumentoConv.MoveFirst
                While Not rsDocumentoConv.EOF()
                    strConv = pfstr_Alineado(rsDocumentoConv("descripcion"), 35, Izquierda) & Space(6) & _
                              pfstr_Alineado(IIf(rsDocumentoConv("flg_retencion") = 0, "Verificar", "Retener"), 10, Izquierda)
                    If intCant = 0 Then
                        Printer.Print Space(2) & "Nombre Convenio : " & pfstr_Alineado(odynCab("CONVENIO").Value, 48) & Space(2) & strConv
                    ElseIf intCant = 1 Then
                        If IIf(IsNull(odynCab("COD_CLIENTE")), "", odynCab("COD_CLIENTE")) <> "" Then
                            Printer.Print Space(2) & "Beneficiario    : " & pfstr_Alineado(odynCab("BENEFICIARIO").Value, 48) & Space(2) & strConv
                        End If
                    Else
                        Printer.Print Space(70) & strConv
                    End If
                    intCant = intCant + 1
                    rsDocumentoConv.MoveNext
                Wend
                If intCant = 0 Then
                    Printer.Print Space(2) & "Nombre Convenio : " & pfstr_Alineado(odynCab("CONVENIO").Value, 60)
                End If
                If intCant = 1 Or intCant = 0 Then
                    If IIf(IsNull(odynCab("COD_CLIENTE")), "", odynCab("COD_CLIENTE")) <> "" Then
                        Printer.Print Space(2) & "Beneficiario    : " & pfstr_Alineado(odynCab("BENEFICIARIO").Value, 60)
                    End If
                End If
                'Printer.Print String(130, "_")
            End If
            'Printer.Print Space(2) & "Observación     : " & pfstr_Alineado(odynCab_OC("DIR_ENTREGA").Value, 91)
            '-------------------------------'
            Printer.FontSize = 7
            Printer.FontName = "Draft 15cpi"
            Printer.Font.Bold = True
            Printer.FontSize = 11
            Printer.Print
         '   Printer.Print 'pfstr_Alineado(odynCab("TIPODOC").Value, 135, centro)
            Printer.FontName = "Draft 17cpi"
            Printer.FontSize = 9
            Printer.FontBold = False
            '-------------------------------'
            Printer.Print String(130, "_")
            Printer.Print pfstr_Alineado("Codigo", 8, centro) & Space(1) & _
                          pfstr_Alineado("Lab", 8, Izquierda) & Space(1) & _
                          pfstr_Alineado("Descripción", 50, Izquierda) & Space(1) & _
                          pfstr_Alineado("Cant.", 8, Derecha) & Space(1) & _
                          pfstr_Alineado("PVP", 8, Derecha) & Space(1) & _
                          pfstr_Alineado("PVPU.", 8, Derecha) & Space(1) & _
                          pfstr_Alineado("Sub Total.", 10, Derecha) & Space(1)
            Printer.Print String(130, "_")
          '---
            intItems = 0
            'Printer.FontName = "Draft 17cpi"
            While Not odynDet.EOF
                  'Detalle
                Printer.Print pfstr_Alineado(odynDet("COD_PRODUCTO").Value, 8) & Space(1) & _
                              pfstr_Alineado(odynDet("LABORATORIO").Value, 8) & Space(1) & _
                              pfstr_Alineado(odynDet("DES_PRODUCTO").Value, 50) & Space(1) & _
                              pfstr_Alineado(odynDet("CANT").Value, 8, Derecha) & _
                              pfstr_Alineado(odynDet("PRC_UNIT_KAIRO").Value, 8, Derecha) & _
                              pfstr_Alineado(odynDet("IMP_UNIT_BASE_IMP").Value, 8, Derecha) & _
                              pfstr_Alineado(odynDet("MTO_SUBTOTAL").Value, 10, Derecha)
                                                  
                dblSumaSubTotal = dblSumaSubTotal + Format(val(odynDet("MTO_SUBTOTAL").Value), "#,###,##0.00")
                intItems = intItems + 1
                odynDet.MoveNext
            Wend
            Printer.Print String(130, "_")
            Printer.Print ""
            Printer.Print "FORMAS DE PAGO"
            Printer.Print String(130, "_")
            
            Printer.FontName = "Roman 15cpi"
            Printer.FontSize = 10
            Printer.FontBold = True
            
            While Not rsFormaPago.EOF
                If rsFormaPago("COD_FORMA_PAGO") <> gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_RED") Then
                    If rsFormaPago("COD_FORMA_PAGO") = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_TJT") Then
                        'Printer.Print pfstr_Alineado(rsFormaPago("DES_FORMA_PAGO"), 10) & Space(1) & pfstr_Alineado(rsFormaPago("DES_HIJO"), 20) & Space(1) & pfstr_Alineado(Format(Mid(rsFormaPago("NUM_TARJETA"), 1, 6), "####-##") & "XX-XXXX-" & Format(Mid(rsFormaPago("NUM_TARJETA"), 13, 17), "####"), 25) & Space(1) & "D.N.I: " & pfstr_Alineado(rsFormaPago("NUM_DOCUMENTO_IDENT"), 9) & Space(1) & "Titular: " & pfstr_Alineado(rsFormaPago("DES_NOM_TITULAR"), 20) & Space(1) & " Nº Cuota: " & pfstr_Alineado(rsFormaPago("NUM_CUOTAS"), 4) & Space(1) & " Nº Aut: " & pfstr_Alineado(rsFormaPago("NUM_AUTORIZACION"), 10) & Space(1) & pfstr_Alineado(rsFormaPago("IMP_MONEDA_NAC"), 8) & Space(1) & pfstr_Alineado(Format(rsFormaPago("IMP_VUELTO"), "#,###,##0.00"), 8)
                        Printer.Print pfstr_Alineado(rsFormaPago("DES_HIJO"), 15) & Space(1) & pfstr_Alineado(Format(Mid(rsFormaPago("NUM_TARJETA"), 1, 6), "####-##") & "XX-XXXX-" & Format(Mid(rsFormaPago("NUM_TARJETA"), 13, 17), "####"), 20) & Space(1) & pfstr_Alineado(rsFormaPago("VERIF_POS"), 6) & Space(1) & "DNI " & pfstr_Alineado(rsFormaPago("NUM_DOCUMENTO_IDENT"), 8) & Space(1) & "Tit.: " & pfstr_Alineado(rsFormaPago("DES_NOM_TITULAR"), 18) & Space(1) & " Cuotas: " & pfstr_Alineado(rsFormaPago("NUM_CUOTAS"), 2) & Space(1) & " Aut: " & pfstr_Alineado(rsFormaPago("NUM_AUTORIZACION"), 10) & Space(1) & "Impor:" & pfstr_Alineado(rsFormaPago("IMP_MONEDA_NAC"), 8) '''& Space(1) & pfstr_Alineado(Format(rsFormaPago("IMP_VUELTO"), "#,###,##0.00"), 8)
                    Else
                        If rsFormaPago("COD_HIJO") = "002" Then
                            Printer.Print pfstr_Alineado(rsFormaPago("DES_FORMA_PAGO"), 15) & Space(5) & pfstr_Alineado(rsFormaPago("DES_HIJO"), 15) & Space(5) & pfstr_Alineado(rsFormaPago("IMP_SIN_REDONDEO"), 28) & Space(5) & " T/C S/. : " & pfstr_Alineado(Format(rsFormaPago("IMP_TIPO_CAMBIO"), "#,###,##0.00"), 15) & Space(5) & " IMPORTE  S/. : " & pfstr_Alineado(Format(rsFormaPago("IMP_MONEDA_NAC"), "#,###,##0.00"), 15)
                        Else
                            Printer.Print pfstr_Alineado(rsFormaPago("DES_FORMA_PAGO"), 15) & Space(5) & pfstr_Alineado(rsFormaPago("DES_HIJO"), 15) & Space(5) & pfstr_Alineado(rsFormaPago("IMP_MONEDA_NAC"), 28) '''& Space(5) & " VUELTO : " & pfstr_Alineado(Format(rsFormaPago("IMP_VUELTO"), "#,###,##0.00"), 15)
                        End If
                    End If
                    If rsFormaPago("IMP_MONEDA_NAC").Value <> rsFormaPago("IMP_VUELTO").Value Then
                        dblSumaTotalPagar = dblSumaTotalPagar + rsFormaPago("IMP_MONEDA_NAC").Value
                        dblSumaTotalVuelto = dblSumaTotalVuelto + rsFormaPago("IMP_VUELTO").Value
                    End If
                End If
                rsFormaPago.MoveNext
            Wend
            Printer.FontName = "Draft 17cpi"
            Printer.FontSize = 8
            Printer.FontBold = False
            
            Printer.Print String(130, "_")
            'Printer.FontBold = False
            For i = 1 To TotItems
                Printer.Print
            Next i
            If "" & odynCab("COD_CONVENIO") = "" Then
                Dim objConvenio As New clsConvenio
                Dim rsDatosAdicionales As oraDynaset
                Set rsDatosAdicionales = objConvenio.ListaDocVerif("" & odynCab("COD_CONVENIO"))
                While Not rsDatosAdicionales.EOF
                    Printer.Print pfstr_Alineado(rsFormaPago("Descripcion"), 30) & Space(5) & pfstr_Alineado(rsFormaPago("FLG_RETENCION"), 30)
                    rsDatosAdicionales.MoveNext
                Wend
                Set objConvenio = Nothing
            Else
                dblSumaSubTotal = dblSumaSubTotal * (val(odynCab("PCT_BENEFICIARIO")) / 100)
            End If
            
            'Printer.Print "detalle aqui"
            
            Printer.Font.name = "Draft 17cpi"
            Printer.FontSize = 8
            Printer.Print
            Printer.Print Space(2) & "Total de Registro => " & " " & intItems & Space(23) & "Total a Cancelar S/. " & " " & right(Space(15) & Format(dblSumaSubTotal, "#,###,##0.00"), 15)
            Printer.Print Space(48) & "Total Pagado     S/. " & " " & right(Space(15) & Format(dblSumaTotalPagar, "#,###,##0.00"), 15)
            Printer.Print
            Printer.FontName = "Roman 15cpi"
            Printer.FontSize = 11
            Printer.Font.Bold = True
            Printer.Print Space(42) & "Vuelto         S/. " & " " & right(Space(12) & Format(dblSumaTotalVuelto, "#,###,##0.00"), 12)
            Printer.FontName = "Draft 17cpi"
            Printer.FontSize = 8
            Printer.Font.Bold = False
            Printer.Print
            If odynCab("FLG_ENTREGA_TERCERO").Value = "1" Then
                Printer.Print Space(2) & "Datos de Entrega"
                Printer.Print Space(2) & String(18, "-")
                Printer.Print
                Printer.Print Space(2) & "Nombre:     " & Space(5) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_NOMBRE").Value, 50)
                Printer.Print Space(2) & "Dirección:  " & Space(5) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_DIRECC").Value, 50) & Space(5) & "Telefono :" & Space(5) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_TLF").Value, 15)
                Printer.Print Space(2) & "Referencia: " & Space(5) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_REF").Value, 50)
                Printer.Print
                Printer.Print
                Printer.Print
            Else
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
            End If
            Printer.Print Space(15) & String(30, "_") & Space(41) & String(30, "_")
            Printer.Print Space(22) & "Firma del Cliente" & Space(55) & "Firma del Local"
            Printer.Print Space(88) & "Nombre: "
            'Printer.Print Space(2) & "Paga con S/. " & " " & Right(Space(15) & Format(odynCab("IMP_CON_REDONDEO").Value, "#,###,##0.00"), 15)
            'Printer.Print Space(2) & "Vuelto   s/. " & " " & Right(Space(15) & Format(odynCab("IMP_VUELTO").Value, "#,###,##0.00"), 15)
            
            'Printer.NewPage
        Wend
        intCopia = intCopia + 1
        Printer.EndDoc
    Loop
    On Error GoTo 0
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
         "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
         "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
    Else
      Exit Sub
    End If
End Sub


'''''''''---- Impresión del Pedido Cabecera enviado por delivery ----'
''''''''Public Function fnImpCabPed(ByVal vstrCia As String, _
''''''''                            ByVal vstrCodBtl As String, _
''''''''                            ByVal vstrProforma As String) As OraDynaset
''''''''    On Error GoTo CtrlErr
''''''''    Set fnImpCabPed = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_IMP_CAB_PED", 0, vstrCia, vstrCodBtl, vstrProforma)
''''''''
''''''''    Exit Function
''''''''CtrlErr:
''''''''    Err.Raise Err.Number, "clsImpresiones", Err.Description
''''''''End Function
''''''''
'''''''''---- Impresión del Pedido Detalle enviado por delivery ----'
''''''''Public Function fnImpDetPed(ByVal vstrCia As String, _
''''''''                            ByVal vstrProforma As String) As OraDynaset
''''''''    On Error GoTo CtrlErr
''''''''    Set fnImpDetPed = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_IMP_DET_PED", 0, vstrCia, vstrProforma)
''''''''
''''''''    Exit Function
''''''''CtrlErr:
''''''''    Err.Raise Err.Number, "clsImpresiones", Err.Description
''''''''End Function

'''''''''Public Sub Proforma(ByVal vstrCia As String, _
'''''''''                    ByVal vstrCodBtl As String, _
'''''''''                    ByVal vstrProforma As String, _
'''''''''                    ByVal vintNumCopia As String)
'''''''''
'''''''''   Dim dblSumaSubTotal As Double
'''''''''   Dim Intlinea As Integer
'''''''''   Dim intCopia As Integer
'''''''''   Dim i As Integer
'''''''''   Dim objImpresiones As New clsProforma
'''''''''   Dim odynCab As OraDynaset
'''''''''   Dim odynDet As OraDynaset
'''''''''   Dim intItems As Integer
'''''''''   Dim TotItems As Integer
'''''''''
'''''''''   Set odynCab = objImpresiones.fnImpCabPed(vstrCia, vstrCodBtl, vstrProforma)
'''''''''   Set odynDet = objImpresiones.fnImpDetPed(vstrCia, vstrProforma)
'''''''''
'''''''''    'Frm_SIS_Impresoras.Show vbModal
'''''''''    Printer.PaperSize = vbPRPSA4
'''''''''
'''''''''    TotItems = "0" & odynCab("ITEM").Value
'''''''''
'''''''''    intCopia = 0: TotItems = 0
'''''''''    If vintNumCopia = 0 Then Exit Sub
'''''''''
'''''''''    '-- Generar Copia --'
'''''''''    Do While vintNumCopia > intCopia
'''''''''        odynCab.MoveFirst
'''''''''        odynDet.MoveFirst
'''''''''        While Not odynDet.EOF
'''''''''            'Cabecera
'''''''''
'''''''''            On Error GoTo Mal_Seteo
'''''''''            GoTo OK_Seteo
'''''''''Mal_Seteo:
'''''''''            MsgBox Err.Description, vbExclamation, "Error de Impresión"
'''''''''            Exit Sub
'''''''''OK_Seteo:
'''''''''            On Error GoTo 0
'''''''''            On Error GoTo ErrorImpresora
'''''''''
'''''''''            Printer.FontName = "Draft 15cpi"
'''''''''            Printer.FontSize = 9
'''''''''            Printer.Print Space(155) & "Fecha :" & " " & odynCab("FECHA").Value
'''''''''            Printer.Print Space(155) & "Hora  :" & " " & odynCab("HORA").Value
'''''''''            Printer.Print
'''''''''            Printer.FontName = "Draft 15cpi"
'''''''''            Printer.Font.Bold = True
'''''''''            Printer.FontSize = 15
'''''''''            Printer.Print pfstr_Alineado("PROFORMA", 125, centro)
'''''''''            Printer.FontName = "Draft 17cpi"
'''''''''            Printer.FontSize = 9
'''''''''            Printer.FontBold = False
'''''''''            Printer.Print
'''''''''            Printer.Print Space(2) & "Atendido        : " & pfstr_Alineado(odynCab("USUARIO").Value, 50) & " " & "Motorizado:" & " " & pfstr_Alineado(odynCab("MOTORIZADO").Value, 30)
'''''''''            Printer.Print Space(2) & "Nro.Proforma    : " & pfstr_Alineado(odynCab("NUM_PROFORMA").Value, 50)
'''''''''            Printer.Print Space(2) & "Codigo Cliente  : " & pfstr_Alineado(odynCab("COD_CLIENTE").Value, 50)
'''''''''            Printer.Print Space(2) & "Telefono        : " & pfstr_Alineado("xxxxxxxx", 15) & Space(35) & "Tarjeta" & " " & pfstr_Alineado(IIf(IsNull(odynCab("NUM_TARJETA").Value), "", odynCab("NUM_TARJETA").Value), 20) & " " & "F.V" & " " & pfstr_Alineado(IIf(IsNull(odynCab("FCH_VENCIMIENTO").Value), "", odynCab("FCH_VENCIMIENTO").Value), 15)
'''''''''            Printer.Print Space(2) & "Nombre          : " & pfstr_Alineado(odynCab("NOMBRE").Value, 85) & " " & "D.N.I." & " " & pfstr_Alineado(IIf(IsNull(odynCab("NUM_DOCUMENTO_ID").Value), "", odynCab("NUM_DOCUMENTO_ID").Value), 13)
'''''''''            Printer.Print Space(2) & "Dirección       : " & pfstr_Alineado(odynCab("DIRECCION").Value, 80)
'''''''''            Printer.Print Space(2) & "Distrito        : " & pfstr_Alineado(odynCab("DISTRITO").Value, 50) & " " & "T.C" & " " & pfstr_Alineado(odynCab("IMP_TIP_CAMBIO").Value, 8) & Space(15) & "Forma Pago" & " " & pfstr_Alineado(odynCab("DES_HIJO").Value, 15)
'''''''''            'Printer.Print Space(2) & "Observación     : " & pfstr_Alineado(odynCab_OC("DIR_ENTREGA").Value, 91)
'''''''''            '-------------------------------'
'''''''''            Printer.FontSize = 7
'''''''''            Printer.FontName = "Draft 15cpi"
'''''''''            Printer.Font.Bold = True
'''''''''            Printer.FontSize = 11
'''''''''            Printer.Print
'''''''''            Printer.Print pfstr_Alineado(odynCab("TIPODOC").Value, 135, centro)
'''''''''            Printer.FontName = "Draft 17cpi"
'''''''''            Printer.FontSize = 9
'''''''''            Printer.FontBold = False
'''''''''            '-------------------------------'
'''''''''            Printer.Print String(210, "_")
'''''''''            Printer.Print pfstr_Alineado("Codigo", 8, centro) & Space(1) & _
'''''''''                          pfstr_Alineado("Lab", 8, Izquierda) & Space(1) & _
'''''''''                          pfstr_Alineado("Descripción", 50, Izquierda) & Space(1) & _
'''''''''                          pfstr_Alineado("Cant.", 8, Derecha) & Space(1) & _
'''''''''                          pfstr_Alineado("PVP", 8, Derecha) & Space(1) & _
'''''''''                          pfstr_Alineado("PVPU.", 8, Derecha) & Space(1) & _
'''''''''                          pfstr_Alineado("Sub Total.", 10, Derecha) & Space(1)
'''''''''            Printer.Print ; String(210, "_")
'''''''''          '---
'''''''''            intItems = 0
'''''''''            'Printer.FontName = "Draft 17cpi"
'''''''''            While Not odynDet.EOF
'''''''''                  'Detalle
'''''''''                Printer.Print pfstr_Alineado(odynDet("COD_PRODUCTO").Value, 8) & Space(1) & _
'''''''''                              pfstr_Alineado(odynDet("DES_LABORATORIO").Value, 8) & Space(1) & _
'''''''''                              pfstr_Alineado(odynDet("DES_PRODUCTO").Value, 50) & Space(1) & _
'''''''''                              pfstr_Alineado(odynDet("CTD_PRODUCTO").Value, 8, Derecha) & _
'''''''''                              pfstr_Alineado(odynDet("PRC_UNIT_KAIRO").Value, 8, Derecha) & _
'''''''''                              pfstr_Alineado(odynDet("IMP_UNIT_BASE_IMP").Value, 8, Derecha) & _
'''''''''                              pfstr_Alineado(odynDet("MTO_SUBTOTAL").Value, 10, Derecha)
'''''''''
'''''''''                dblSumaSubTotal = dblSumaSubTotal + Format(Val(odynDet("MTO_SUBTOTAL").Value), "#,###,##0.00")
'''''''''                intItems = intItems + 1
'''''''''                odynDet.MoveNext
'''''''''            Wend
'''''''''            'Printer.FontBold = False
'''''''''            For i = 1 To TotItems
'''''''''                Printer.Print
'''''''''            Next i
'''''''''            'Printer.Print "detalle aqui"
'''''''''            Printer.Print String(210, "_")
'''''''''            Printer.Font.name = "Draft 15cpi"
'''''''''            Printer.FontSize = 8
'''''''''            Printer.Print
'''''''''            Printer.Print Space(2) & "Total de Registro => " & " " & intItems & Space(90) & "Total a Cancelar S/. " & " " & Right(Space(15) & Format(dblSumaSubTotal, "#,###,##0.00"), 15)
'''''''''            Printer.Print
'''''''''            Printer.Print Space(2) & "Paga con S/. " & " " & Right(Space(15) & Format(odynCab("IMP_CON_REDONDEO").Value, "#,###,##0.00"), 15)
'''''''''            Printer.Print Space(2) & "Vuelto   s/. " & " " & Right(Space(15) & Format(odynCab("IMP_VUELTO").Value, "#,###,##0.00"), 15)
'''''''''
'''''''''            'Printer.NewPage
'''''''''        Wend
'''''''''        intCopia = intCopia + 1
'''''''''        Printer.EndDoc
'''''''''    Loop
'''''''''    On Error GoTo 0
'''''''''    Exit Sub
'''''''''ErrorImpresora:
'''''''''    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
'''''''''         "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
'''''''''         "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'''''''''        Resume
'''''''''    Else
'''''''''      Exit Sub
'''''''''    End If
'''''''''End Sub

Public Sub ImprimirTicket_Delivery(ByVal vstrCia As String, _
                                   ByVal vstrCodBtl As String, _
                                   ByVal vstrProforma As String, _
                                   ByVal vintNumCopia As String, _
                                   ByVal vflgDetalle As Boolean)

   'Autor: mlaguna
   'Motivo: impresion de proforma en ticketera
   'fecha: 31/03/2010

   Dim dblSumaSubTotal As Double
   Dim intLinea As Integer
   Dim intCopia As Integer
   Dim i As Integer
   Dim objImpresiones As New clsProforma
   Dim odynCab As oraDynaset
   Dim odynDet As oraDynaset
   Dim rsFormaPago As oraDynaset
   Dim rsDocxProf As oraDynaset
   Dim TotItems As String
   Dim dblSumaTotalPagar As Double
   Dim dblSumaTotalVuelto As Double
   Dim p As Printer
   Dim j As Integer
   Dim vInd As Integer
   
   Set odynCab = objImpresiones.ListaCabecera(vstrCia, vstrCodBtl, vstrProforma)
   Set odynDet = objImpresiones.ListaDetalle(vstrCia, vstrCodBtl, vstrProforma)
   Set rsFormaPago = objImpresiones.ListaFormaPago(vstrCia, vstrCodBtl, vstrProforma)
   Set rsDocxProf = objImpresiones.Lista_DocsxProf(vstrCia, vstrCodBtl, vstrProforma)
                
    'Frm_SIS_Impresoras.Show vbModal
    'Printer.PaperSize = vbPRPSA4
    
    TotItems = "0" & odynCab("ITEM").Value
    
    intCopia = 0: TotItems = 0: dblSumaSubTotal = 0: dblSumaTotalPagar = 0: dblSumaTotalVuelto = 0
    If vintNumCopia = 0 Then Exit Sub

    vInd = 0

    '***Comenta espacio vacío

    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          vInd = "1"
          Exit For
       End If
    Next p

    On Error GoTo Mal_Seteo
    GoTo OK_Seteo
Mal_Seteo:
    MsgBox Err.Description, vbExclamation, "Error de Impresión"
    Exit Sub
OK_Seteo:
    On Error GoTo 0
    On Error GoTo ErrorImpresora

    Open "lpt1" For Output As #1
    
    '***comenta espacion vacio
    
    '-- Generar Copia --'
    Do While vintNumCopia > intCopia
        odynCab.MoveFirst
        odynDet.MoveFirst
        'While Not odynDet.EOF
            'Cabecera
            
            '''Open "COM3" For Output As #1
            '**** 16-DOT PITCH FONT *******'
                        
            '-------------------------------'
            
            Print #1, Chr$(27); Chr$(58);
            
            If odynCab("FLG_ENTREGA_TERCERO").Value = "0" Then
                'Print #1, ""
                Print #1, "*** PEDIDO DELIVERY ***"
                Print #1, Chr$(27); Chr$(77);
                Print #1, pfstr_Alineado("*** Este documento no es un comprobante de", 42, centro)
                Print #1, pfstr_Alineado("pago. Solicite su comprobante de pago  ***", 42, centro)
                
                Print #1, ""
              Else
                'Print #1, ""
                Print #1, "*** PEDIDO DELIVERY ***"
                Print #1, Chr$(27); Chr$(77);
                Print #1, pfstr_Alineado("*** Este documento no es un comprobante de", 42, centro)
                Print #1, pfstr_Alineado("pago. Solicite su comprobante de pago  ***", 42, centro)
                Print #1, Chr$(27); Chr$(58);
                Print #1, "*** CONEXION SALUD  ***"
                Print #1, ""
            End If
            
            Print #1, Chr$(27); Chr$(77);
            Printer.Font = 20
            
            Print #1, "Fecha:" & Space(1) & odynCab("FECHA").Value  '" " &
            'Print #1, "Hora:" & Space(1) & odynCab("HORA").Value   '" " &
            'Print #1, "Hora Avisado:" & Space(1) & odynCab("AVISADO").Value  '& " "
            'Print #1, "Hora Proformado:" & Space(1) & odynCab("PROFORMADO").Value '& " "
            
            If IsNull(odynCab("FCH_HORA_PACT_ENTR").Value) = False And Trim(odynCab("FCH_HORA_PACT_ENTR").Value) <> "" Then
            Print #1, "Fecha Pactada:" & odynCab("FCH_HORA_PACT_ENTR").Value '& " "
            End If
            
            Print #1, Chr$(27); Chr$(58);

            Print #1, Chr$(27); Chr$(77);
            'Print #1, "Atendido:" & Space(1) & odynCab("CAJERA").Value
            Print #1, "Motorizado:" & Space(1) & odynCab("MOTORIZADO").Value
            Print #1, "Nro.Proforma:" & Space(1) & odynCab("NUM_PROFORMA").Value
            Print #1, "Nombre:" & Space(1) & odynCab("NOMBRES").Value
            
            If IIf(IsNull(odynCab("NUM_DOCUMENTO_ID").Value), "", odynCab("NUM_DOCUMENTO_ID").Value) <> "" Then
                Print #1, "D.N.I.:" & Space(1) & IIf(IsNull(odynCab("NUM_DOCUMENTO_ID").Value), "", odynCab("NUM_DOCUMENTO_ID").Value)
            End If
            
            Print #1, "Direccion:" & Space(1) & IIf(IsNull(odynCab("DES_SUFIJO").Value), "", odynCab("DES_SUFIJO").Value) & Space(1) & _
                                   IIf(IsNull(odynCab("DIRECCION").Value), "", odynCab("DIRECCION").Value)
            Print #1, "Distrito:" & " " & odynCab("DISTRITO").Value
            'Print #1, "T.C" & " " & odynCab("IMP_TIP_CAMBIO").Value
            
            If IsNull(odynCab("URBANIZACION").Value) = False And Trim(odynCab("URBANIZACION").Value) <> "" Then
                Print #1, "Urbanizacion:" & " " & odynCab("URBANIZACION").Value
            End If
            
            
            If IsNull(odynCab("REFERENCIAS").Value) = False And Trim(odynCab("REFERENCIAS").Value) <> "" Then
                 Print #1, "Referencia:" & " " & odynCab("REFERENCIAS").Value
            End If
            
            If IsNull(odynCab("OBS_NOTA_LOCAL").Value) = False And Trim(odynCab("OBS_NOTA_LOCAL").Value) <> "" Then
                Print #1, "Observacion Local:" & " " & odynCab("OBS_NOTA_LOCAL").Value
            End If
            
             If IsNull(odynCab("OBS_NOTA_MOTORIZADO").Value) = False And Trim(odynCab("OBS_NOTA_MOTORIZADO").Value) <> "" Then
                Print #1, "Observacion Motorizado:" & " " & odynCab("OBS_NOTA_MOTORIZADO").Value
             End If
            
            
            Print #1, ""

            'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx '
            'Documentos a emitir
            If rsDocxProf.RecordCount > 0 Then
                    Dim c%
                    Dim strDoc1 As String
                    Dim strDoc2 As String
                    
                    If IIf(IsNull(odynCab("COD_CONVENIO")), "", odynCab("COD_CONVENIO")) <> "" Then
                    Print #1, Space(5) & odynCab("DES_MODALIDAD").Value & " : "
                    End If
                    
                    rsDocxProf.MoveFirst
                    c = 0
                    While Not rsDocxProf.EOF
                        c = c + 1
                        If c = 1 Then
                           strDoc1 = "" & rsDocxProf("TIPODOC").Value & " : " & rsDocxProf("NUM_DOCUMENTO").Value
                            Print #1, strDoc1
                        ElseIf c = 2 Then
                           strDoc2 = "" & rsDocxProf("TIPODOC").Value & " : " & rsDocxProf("NUM_DOCUMENTO").Value
                           Print #1, strDoc2
                        End If
                        rsDocxProf.MoveNext
                    Wend
            End If

            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx '
            'Documentos a verificar
            If IIf(IsNull(odynCab("COD_CONVENIO")), "", odynCab("COD_CONVENIO")) <> "" Then
                
                Print #1, "CONVENIO" & Space(3)
                Print #1, "Nombre Convenio : " & left(odynCab("CONVENIO").Value, 40)
                If IIf(IsNull(odynCab("COD_CLIENTE")), "", odynCab("COD_CLIENTE")) <> "" Then
                    Print #1, "Beneficiario : " & left(odynCab("BENEFICIARIO").Value, 30)
                End If
                
                'Dim strConv As String
                Dim intCant As Byte
                Dim rsDocumentoConv As oraDynaset
                Dim objConvenio1 As New clsConvenio
                Set rsDocumentoConv = objConvenio1.ListaDocVerif(odynCab("COD_CONVENIO"))
                'Print #1, ""

                If Not rsDocumentoConv.EOF Then
                    Print #1, String(42, "-")
                    Print #1, "Verificacion de Documentos" & Space(2) '& String(10, "-")
                    Print #1, String(42, "-")
                    'Print #1, pfstr_Alineado("Descripcion", 10, Izquierda) & Space(9) & _
                    '          pfstr_Alineado("Retencion", 10, Izquierda)
                    'Print #1, String(42, "-")
                Else
                    'Print #1, "CONVENIO"
                End If
                
                'Documentos
                intCant = 0
                rsDocumentoConv.MoveFirst
                While Not rsDocumentoConv.EOF()
                    
                    'If intCant = 0 Then
                        
                    'ElseIf intCant = 1 Then
                        
                    'Else
'                        Print #1, strConv
                    'End If
                    
                    'Print #1, left(rsDocumentoConv("descripcion"), 30) & Space(1) & IIf(rsDocumentoConv("flg_retencion") = 0, "Verificar", "Retener")
                    
                    Print #1, pfstr_Alineado(Mid(rsDocumentoConv("descripcion"), 1, 30), 30, Izquierda) & Space(1) & _
                              pfstr_Alineado(IIf(rsDocumentoConv("flg_retencion") = 0, "Verificar", "Retener"), 10, Izquierda)
                        
                    
                    intCant = intCant + 1
                    rsDocumentoConv.MoveNext
                Wend

'                If intCant = 0 Then
'                    Print #1, "Nombre Convenio: " & pfstr_Alineado(odynCab("CONVENIO").Value, 30)
'                End If
'
'                If intCant = 1 Or intCant = 0 Then
'                    If IIf(IsNull(odynCab("COD_CLIENTE")), "", odynCab("COD_CLIENTE")) <> "" Then
'                        Print #1, "Beneficiario: " & pfstr_Alineado(odynCab("BENEFICIARIO").Value, 30)
'                    End If
'                End If
            End If
            
            'Print #1, String(42, "-") 'Separacion
            
            'NO DEBE SALIR EN TICKET DELIVERY, SOLO EN EL CASO DE REIMPRESIONES SI SALE.
            '***
            If vflgDetalle = True Then
                                    
                    'Cabecera de articulos
                    Print #1, String(42, "-")
                    Print #1, pfstr_Alineado("Codigo", 5, centro) & Space(1) & _
                              pfstr_Alineado("Descripcion", 20, Izquierda) & Space(1) & _
                              pfstr_Alineado("Cant.", 6, Derecha) & Space(1) & _
                              pfstr_Alineado("SubTotal", 8, Derecha)
                    Print #1, String(42, "-")
        
                    'Detalle articulos
                    intItems = 0
                    While Not odynDet.EOF
                          'Detalle
                        If Len(odynDet("DES_PRODUCTO").Value) > 20 Then ' SI LA DESCRIPCION DE PRODUCTO ES MAYOR A 20 CARACTERES'
                            Print #1, pfstr_Alineado(odynDet("COD_PRODUCTO").Value, 5) & Space(1) & _
                                      left(Mid(odynDet("DES_PRODUCTO").Value, 1, 20) & Space(20), 20) & Space(1) & _
                                      pfstr_Alineado(odynDet("CANT").Value, 5, Derecha) & _
                                      pfstr_Alineado(odynDet("MTO_SUBTOTAL").Value, 8, Derecha)
                            Print #1, Space(5) & left(Mid(odynDet("DES_PRODUCTO").Value, 21, 15), 20)
                        Else
                            Print #1, pfstr_Alineado(odynDet("COD_PRODUCTO").Value, 5) & Space(1) & _
                                      left(Mid(odynDet("DES_PRODUCTO").Value, 1, 20), 20) & Space(1) & _
                                      pfstr_Alineado(odynDet("CANT").Value, 5, Derecha) & _
                                      pfstr_Alineado(odynDet("MTO_SUBTOTAL").Value, 8, Derecha)
                        End If
                        dblSumaSubTotal = dblSumaSubTotal + Format(val(odynDet("MTO_SUBTOTAL").Value), "#,###,##0.00")
                        intItems = intItems + 1
                        odynDet.MoveNext
                    Wend
        
                    'Print #1, ""
                    
            End If '***
            
            'FIN DE ARTICULOS
            
            If rsFormaPago.RecordCount > 0 Then
                 Print #1, String(42, "-")
                Print #1, pfstr_Alineado("Forma(s) de Pago:", 16, centro) & Space(1)
                Print #1, String(42, "-")
            End If
            
            If rsFormaPago("COD_FORMA_PAGO") = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_TJT") Then
                Print #1, "Pago con tarjeta:"
            Else
                'Print #1, "FORMAS DE PAGO"
            End If
            
            'Print #1, ""
               
            
            While Not rsFormaPago.EOF
            
               
                
                If rsFormaPago("COD_FORMA_PAGO") <> gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_RED") Then
                    If rsFormaPago("COD_FORMA_PAGO") = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_TJT") Then 'Pago con tarjeta
                        Print #1, rsFormaPago("DES_HIJO") & " => " & pfstr_Alineado(Format(Mid(rsFormaPago("NUM_TARJETA"), 1, 6), "####-##") & "XX-XXXX-" & Format(Mid(rsFormaPago("NUM_TARJETA"), 13, 17), "####"), 20)  'pfstr_Alineado(, 10)
                        Print #1, rsFormaPago("VERIF_POS")                          'pfstr_Alineado(, 6)
                        Print #1, "DNI " & rsFormaPago("NUM_DOCUMENTO_IDENT")       'pfstr_Alineado(, 8)
                        Print #1, "Tit.: " & rsFormaPago("DES_NOM_TITULAR")         'pfstr_Alineado(, 18)
                        Print #1, "Cuotas: " & rsFormaPago("NUM_CUOTAS")            'pfstr_Alineado(, 2)
                        Print #1, "Nro. Aut: " & rsFormaPago("NUM_AUTORIZACION")    'pfstr_Alineado(, 10)
                        Print #1, "Importe:" & rsFormaPago("IMP_MONEDA_NAC")        'pfstr_Alineado(, 8)
                    
                    Else 'Diferente de pago con tarjeta
                        If rsFormaPago("COD_HIJO") = "002" Then 'Dolares
                            Print #1, pfstr_Alineado(rsFormaPago("DES_FORMA_PAGO"), 10) & Space(5) & _
                                      pfstr_Alineado(rsFormaPago("DES_HIJO"), 15) & Space(5) & _
                                      pfstr_Alineado(Format(rsFormaPago("IMP_SIN_REDONDEO"), "#0.00"), 28) & Space(5) & _
                                      "T/C S/. : " & pfstr_Alineado(Format(rsFormaPago("IMP_TIPO_CAMBIO"), "#,###,##0.00"), 15) & Space(5) & _
                                      "IMPORTE  S/. : " & pfstr_Alineado(Format(rsFormaPago("IMP_MONEDA_NAC"), "#,###,##0.00"), 15)
                        Else 'Otras formas de pago
                            Print #1, pfstr_Alineado(rsFormaPago("DES_FORMA_PAGO"), 10) & Space(5) & _
                                      pfstr_Alineado(rsFormaPago("DES_HIJO"), 15) & Space(5) & pfstr_Alineado(Format(rsFormaPago("IMP_MONEDA_NAC"), "#0.00"), 5)
                        End If
                     End If
                    
                    'If rsFormaPago("IMP_MONEDA_NAC").Value <> rsFormaPago("IMP_VUELTO").Value Then
                        dblSumaTotalPagar = dblSumaTotalPagar + rsFormaPago("IMP_MONEDA_NAC").Value
                        dblSumaTotalVuelto = dblSumaTotalVuelto + rsFormaPago("IMP_VUELTO").Value
                    'End If
                End If
                rsFormaPago.MoveNext
            Wend
            
            'Print #1, String(42, "_") 'Separacion
            
'            For i = 1 To TotItems
'                Print #1, ""
'            Next i
            
            If "" & odynCab("COD_CONVENIO") = "" Then
                Dim objConvenio As New clsConvenio
                Dim rsDatosAdicionales As oraDynaset
                Set rsDatosAdicionales = objConvenio.ListaDocVerif("" & odynCab("COD_CONVENIO"))
                While Not rsDatosAdicionales.EOF
                    Print #1, pfstr_Alineado(rsFormaPago("Descripcion"), 30) & Space(5) & pfstr_Alineado(rsFormaPago("FLG_RETENCION"), 30)
                    rsDatosAdicionales.MoveNext
                Wend
                Set objConvenio = Nothing
            Else
                dblSumaSubTotal = dblSumaSubTotal * (val(odynCab("PCT_BENEFICIARIO")) / 100)
            End If

            'Printer.Print "detalle aqui"
              
            'Total documento
            If vflgDetalle = True Then
                Print #1, pfstr_Alineado(String(22, "-"), 42, Derecha)
                Print #1, Space(1) & "Total de Registro => " & " " & intItems
                Print #1, Space(1) & "Total a Cancelar S/. " & " " & pfstr_Alineado(right(Space(5) & Format(dblSumaSubTotal, "#,###,##0.00"), 15), 15, Derecha)
                Print #1, Space(1) & "    Total Pagado S/. " & " " & pfstr_Alineado(right(Space(5) & Format(dblSumaTotalPagar, "#,###,##0.00"), 15), 15, Derecha)
                Print #1, Space(1) & "          Vuelto S/. " & " " & pfstr_Alineado(right(Space(5) & Format(dblSumaTotalVuelto, "#,###,##0.00"), 15), 15, Derecha)
                'Print #1, Space(1) & "T.C. :" & " " & pfstr_Alineado(odynCab("IMP_TIP_CAMBIO").Value, 8)
                Print #1, ""
            End If

            If odynCab("FLG_ENTREGA_TERCERO").Value = "1" Then
                Print #1, Space(2) & "Datos de Entrega"
                Print #1, Space(2) & String(18, "-")
                Print #1, ""
                Print #1, "Nombre:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_NOMBRE").Value, 30)
                Print #1, "Direccion:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_DIRECC").Value, 50)
                Print #1, "Telefono:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_TLF").Value, 15)
                Print #1, "Referencia:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_REF").Value, 50)
                Print #1, ""
            Else
                Print #1, ""
            End If
        'Wend
        
        Print #1, pfstr_Alineado("*** Este documento no es un comprobante de", 42, centro)
        Print #1, pfstr_Alineado("pago. Solicite su comprobante de pago  ***", 42, centro)
        
        
        intCopia = intCopia + 1

        For j = 1 To 5
          Print #1, ""
        Next j

        'Copia = Copia + 1
        '** Corte Parcial **'
        
        Print #1, Chr(27) + "d" + "1"
        
        '** Corte Total **'
        'Print #1, Chr$(27); "d"; Chr$(2)

        '*********************************************************************************************************
        Close #1
    Loop
    On Error GoTo 0
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
         "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
         "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
    Else
      Exit Sub
    End If
End Sub



Public Sub ImprimirTicket_DeliveryNew(ByVal vstrCia As String, _
                                   ByVal vstrCodBtl As String, _
                                   ByVal vstrProforma As String, _
                                   ByVal vintNumCopia As String, _
                                   ByVal vflgDetalle As Boolean)

   'Autor: mlaguna
   'Motivo: impresion de proforma en ticketera
   'fecha: 31/03/2010

   Dim dblSumaSubTotal As Double
   Dim intLinea As Integer
   Dim intCopia As Integer
   Dim i As Integer
   Dim objImpresiones As New clsProforma
   Dim odynCab As oraDynaset
   Dim odynDet As oraDynaset
   Dim rsFormaPago As oraDynaset
   Dim rsDocxProf As oraDynaset
   Dim TotItems As String
   Dim dblSumaTotalPagar As Double
   Dim dblSumaTotalVuelto As Double
   Dim p As Printer
   Dim j As Integer
   Dim vInd As Integer
   
   Set odynCab = objImpresiones.ListaCabecera(vstrCia, vstrCodBtl, vstrProforma)
   Set odynDet = objImpresiones.ListaDetalle(vstrCia, vstrCodBtl, vstrProforma)
   Set rsFormaPago = objImpresiones.ListaFormaPago(vstrCia, vstrCodBtl, vstrProforma)
   Set rsDocxProf = objImpresiones.Lista_DocsxProf(vstrCia, vstrCodBtl, vstrProforma)
                
    'Frm_SIS_Impresoras.Show vbModal
    'Printer.PaperSize = vbPRPSA4
    
    TotItems = "0" & odynCab("ITEM").Value
    
    intCopia = 0: TotItems = 0: dblSumaSubTotal = 0: dblSumaTotalPagar = 0: dblSumaTotalVuelto = 0
    If vintNumCopia = 0 Then Exit Sub

    vInd = 0

    '***Comenta espacio vacío

'    For Each p In Printers
'       If UCase(p.Port) = "LPT1:" Then
'          Set Printer = p
'          vInd = "1"
'          Exit For
'       End If
'    Next p

    frm_VTA_ElegirImpresoras.Show vbModal

    On Error GoTo Mal_Seteo
    GoTo OK_Seteo
Mal_Seteo:
    MsgBox Err.Description, vbExclamation, "Error de Impresión"
    Exit Sub
OK_Seteo:
    On Error GoTo 0
    On Error GoTo ErrorImpresora

    'Open "lpt1" For Output As #1
    
    '***comenta espacion vacio
    
    '-- Generar Copia --'
    Do While vintNumCopia > intCopia
        odynCab.MoveFirst
        odynDet.MoveFirst
        'While Not odynDet.EOF
            'Cabecera
            
            '''Open "COM3" For Output As #1
            '**** 16-DOT PITCH FONT *******'
                        
            '-------------------------------'
            
            Printer.Print Chr$(27); Chr$(58);
            
            If odynCab("FLG_ENTREGA_TERCERO").Value = "0" Then
                'Print #1, ""
                Printer.Print "*** PEDIDO DELIVERY ***"
                Printer.Print Chr$(27); Chr$(77);
                Printer.Print pfstr_Alineado("*** Este documento no es un comprobante de", 42, centro)
                Printer.Print pfstr_Alineado("pago. Solicite su comprobante de pago  ***", 42, centro)
                
                Printer.Print ""
              Else
                'Print #1, ""
                Printer.Print "*** PEDIDO DELIVERY ***"
                Printer.Print Chr$(27); Chr$(77);
                Printer.Print pfstr_Alineado("*** Este documento no es un comprobante de", 42, centro)
                Printer.Print pfstr_Alineado("pago. Solicite su comprobante de pago  ***", 42, centro)
                Printer.Print Chr$(27); Chr$(58);
                Printer.Print "*** CONEXION SALUD  ***"
                Printer.Print ""
            End If
            
            Printer.Print Chr$(27); Chr$(77);
            Printer.Font = 20
            
            Printer.Print "Fecha:" & Space(1) & odynCab("FECHA").Value  '" " &
            'Print #1, "Hora:" & Space(1) & odynCab("HORA").Value   '" " &
            'Print #1, "Hora Avisado:" & Space(1) & odynCab("AVISADO").Value  '& " "
            'Print #1, "Hora Proformado:" & Space(1) & odynCab("PROFORMADO").Value '& " "
            
            If IsNull(odynCab("FCH_HORA_PACT_ENTR").Value) = False And Trim(odynCab("FCH_HORA_PACT_ENTR").Value) <> "" Then
            Printer.Print "Fecha Pactada:" & odynCab("FCH_HORA_PACT_ENTR").Value '& " "
            End If
            
            Printer.Print Chr$(27); Chr$(58);

            Printer.Print Chr$(27); Chr$(77);
            'Print #1, "Atendido:" & Space(1) & odynCab("CAJERA").Value
            Printer.Print "Motorizado:" & Space(1) & odynCab("MOTORIZADO").Value
            Printer.Print "Nro.Proforma:" & Space(1) & odynCab("NUM_PROFORMA").Value
            Printer.Print "Nombre:" & Space(1) & odynCab("NOMBRES").Value
            
            If IIf(IsNull(odynCab("NUM_DOCUMENTO_ID").Value), "", odynCab("NUM_DOCUMENTO_ID").Value) <> "" Then
                Printer.Print "D.N.I.:" & Space(1) & IIf(IsNull(odynCab("NUM_DOCUMENTO_ID").Value), "", odynCab("NUM_DOCUMENTO_ID").Value)
            End If
            
            Printer.Print "Direccion:" & Space(1) & IIf(IsNull(odynCab("DES_SUFIJO").Value), "", odynCab("DES_SUFIJO").Value) & Space(1) & _
                                   IIf(IsNull(odynCab("DIRECCION").Value), "", odynCab("DIRECCION").Value)
            Printer.Print "Distrito:" & " " & odynCab("DISTRITO").Value
            'Print #1, "T.C" & " " & odynCab("IMP_TIP_CAMBIO").Value
            
            If IsNull(odynCab("URBANIZACION").Value) = False And Trim(odynCab("URBANIZACION").Value) <> "" Then
                Printer.Print "Urbanizacion:" & " " & odynCab("URBANIZACION").Value
            End If
            
            
            If IsNull(odynCab("REFERENCIAS").Value) = False And Trim(odynCab("REFERENCIAS").Value) <> "" Then
                 Printer.Print "Referencia:" & " " & odynCab("REFERENCIAS").Value
            End If
            
            If IsNull(odynCab("OBS_NOTA_LOCAL").Value) = False And Trim(odynCab("OBS_NOTA_LOCAL").Value) <> "" Then
                Printer.Print "Observacion Local:" & " " & odynCab("OBS_NOTA_LOCAL").Value
            End If
            
             If IsNull(odynCab("OBS_NOTA_MOTORIZADO").Value) = False And Trim(odynCab("OBS_NOTA_MOTORIZADO").Value) <> "" Then
                Printer.Print "Observacion Motorizado:" & " " & odynCab("OBS_NOTA_MOTORIZADO").Value
             End If
            
            
            Printer.Print ""

            'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx '
            'Documentos a emitir
            If rsDocxProf.RecordCount > 0 Then
                    Dim c%
                    Dim strDoc1 As String
                    Dim strDoc2 As String
                    
                    If IIf(IsNull(odynCab("COD_CONVENIO")), "", odynCab("COD_CONVENIO")) <> "" Then
                    Printer.Print Space(5) & odynCab("DES_MODALIDAD").Value & " : "
                    End If
                    
                    rsDocxProf.MoveFirst
                    c = 0
                    While Not rsDocxProf.EOF
                        c = c + 1
                        If c = 1 Then
                           strDoc1 = "" & rsDocxProf("TIPODOC").Value & " : " & rsDocxProf("NUM_DOCUMENTO").Value
                            Printer.Print strDoc1
                        ElseIf c = 2 Then
                           strDoc2 = "" & rsDocxProf("TIPODOC").Value & " : " & rsDocxProf("NUM_DOCUMENTO").Value
                           Printer.Print strDoc2
                        End If
                        rsDocxProf.MoveNext
                    Wend
            End If

            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx '
            'Documentos a verificar
            If IIf(IsNull(odynCab("COD_CONVENIO")), "", odynCab("COD_CONVENIO")) <> "" Then
                
                Printer.Print "CONVENIO" & Space(3)
                Printer.Print "Nombre Convenio : " & left(odynCab("CONVENIO").Value, 40)
                If IIf(IsNull(odynCab("COD_CLIENTE")), "", odynCab("COD_CLIENTE")) <> "" Then
                    Printer.Print "Beneficiario : " & left(odynCab("BENEFICIARIO").Value, 30)
                End If
                
                'Dim strConv As String
                Dim intCant As Byte
                Dim rsDocumentoConv As oraDynaset
                Dim objConvenio1 As New clsConvenio
                Set rsDocumentoConv = objConvenio1.ListaDocVerif(odynCab("COD_CONVENIO"))
                'Print #1, ""

                If Not rsDocumentoConv.EOF Then
                    Printer.Print String(42, "-")
                    Printer.Print "Verificacion de Documentos" & Space(2) '& String(10, "-")
                    Printer.Print String(42, "-")
                    'Print #1, pfstr_Alineado("Descripcion", 10, Izquierda) & Space(9) & _
                    '          pfstr_Alineado("Retencion", 10, Izquierda)
                    'Print #1, String(42, "-")
                Else
                    'Print #1, "CONVENIO"
                End If
                
                'Documentos
                intCant = 0
                rsDocumentoConv.MoveFirst
                While Not rsDocumentoConv.EOF()
                    
                    'If intCant = 0 Then
                        
                    'ElseIf intCant = 1 Then
                        
                    'Else
'                        Print #1, strConv
                    'End If
                    
                    'Print #1, left(rsDocumentoConv("descripcion"), 30) & Space(1) & IIf(rsDocumentoConv("flg_retencion") = 0, "Verificar", "Retener")
                    
                    Printer.Print pfstr_Alineado(Mid(rsDocumentoConv("descripcion"), 1, 30), 30, Izquierda) & Space(1) & _
                              pfstr_Alineado(IIf(rsDocumentoConv("flg_retencion") = 0, "Verificar", "Retener"), 10, Izquierda)
                        
                    
                    intCant = intCant + 1
                    rsDocumentoConv.MoveNext
                Wend

'                If intCant = 0 Then
'                    Print #1, "Nombre Convenio: " & pfstr_Alineado(odynCab("CONVENIO").Value, 30)
'                End If
'
'                If intCant = 1 Or intCant = 0 Then
'                    If IIf(IsNull(odynCab("COD_CLIENTE")), "", odynCab("COD_CLIENTE")) <> "" Then
'                        Print #1, "Beneficiario: " & pfstr_Alineado(odynCab("BENEFICIARIO").Value, 30)
'                    End If
'                End If
            End If
            
            'Print #1, String(42, "-") 'Separacion
            
            'NO DEBE SALIR EN TICKET DELIVERY, SOLO EN EL CASO DE REIMPRESIONES SI SALE.
            '***
            If vflgDetalle = True Then
                                    
                    'Cabecera de articulos
                    Printer.Print String(42, "-")
                    Printer.Print pfstr_Alineado("Codigo", 5, centro) & Space(1) & _
                              pfstr_Alineado("Descripcion", 20, Izquierda) & Space(1) & _
                              pfstr_Alineado("Cant.", 6, Derecha) & Space(1) & _
                              pfstr_Alineado("SubTotal", 8, Derecha)
                    Printer.Print String(42, "-")
        
                    'Detalle articulos
                    intItems = 0
                    While Not odynDet.EOF
                          'Detalle
                        If Len(odynDet("DES_PRODUCTO").Value) > 20 Then ' SI LA DESCRIPCION DE PRODUCTO ES MAYOR A 20 CARACTERES'
                            Printer.Print pfstr_Alineado(odynDet("COD_PRODUCTO").Value, 5) & Space(1) & _
                                      left(Mid(odynDet("DES_PRODUCTO").Value, 1, 20) & Space(20), 20) & Space(1) & _
                                      pfstr_Alineado(odynDet("CANT").Value, 5, Derecha) & _
                                      pfstr_Alineado(odynDet("MTO_SUBTOTAL").Value, 8, Derecha)
                            Printer.Print Space(5) & left(Mid(odynDet("DES_PRODUCTO").Value, 21, 15), 20)
                        Else
                            Printer.Print pfstr_Alineado(odynDet("COD_PRODUCTO").Value, 5) & Space(1) & _
                                      left(Mid(odynDet("DES_PRODUCTO").Value, 1, 20), 20) & Space(1) & _
                                      pfstr_Alineado(odynDet("CANT").Value, 5, Derecha) & _
                                      pfstr_Alineado(odynDet("MTO_SUBTOTAL").Value, 8, Derecha)
                        End If
                        dblSumaSubTotal = dblSumaSubTotal + Format(val(odynDet("MTO_SUBTOTAL").Value), "#,###,##0.00")
                        intItems = intItems + 1
                        odynDet.MoveNext
                    Wend
        
                    'Print #1, ""
                    
            End If '***
            
            'FIN DE ARTICULOS
            
            If rsFormaPago.RecordCount > 0 Then
                 Printer.Print String(42, "-")
                Printer.Print pfstr_Alineado("Forma(s) de Pago:", 16, centro) & Space(1)
                Printer.Print String(42, "-")
            End If
            
            If rsFormaPago("COD_FORMA_PAGO") = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_TJT") Then
                Printer.Print "Pago con tarjeta:"
            Else
                'Print #1, "FORMAS DE PAGO"
            End If
            
            'Print #1, ""
               
            
            While Not rsFormaPago.EOF
            
               
                
                If rsFormaPago("COD_FORMA_PAGO") <> gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_RED") Then
                    If rsFormaPago("COD_FORMA_PAGO") = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_TJT") Then 'Pago con tarjeta
                        Printer.Print rsFormaPago("DES_HIJO") & " => " & pfstr_Alineado(Format(Mid(rsFormaPago("NUM_TARJETA"), 1, 6), "####-##") & "XX-XXXX-" & Format(Mid(rsFormaPago("NUM_TARJETA"), 13, 17), "####"), 20)  'pfstr_Alineado(, 10)
                        Printer.Print rsFormaPago("VERIF_POS")                          'pfstr_Alineado(, 6)
                        Printer.Print "DNI " & rsFormaPago("NUM_DOCUMENTO_IDENT")       'pfstr_Alineado(, 8)
                        Printer.Print "Tit.: " & rsFormaPago("DES_NOM_TITULAR")         'pfstr_Alineado(, 18)
                        Printer.Print "Cuotas: " & rsFormaPago("NUM_CUOTAS")            'pfstr_Alineado(, 2)
                        Printer.Print "Nro. Aut: " & rsFormaPago("NUM_AUTORIZACION")    'pfstr_Alineado(, 10)
                        Printer.Print "Importe:" & rsFormaPago("IMP_MONEDA_NAC")        'pfstr_Alineado(, 8)
                    
                    Else 'Diferente de pago con tarjeta
                        If rsFormaPago("COD_HIJO") = "002" Then 'Dolares
                            Printer.Print pfstr_Alineado(rsFormaPago("DES_FORMA_PAGO"), 10) & Space(5) & _
                                      pfstr_Alineado(rsFormaPago("DES_HIJO"), 15) & Space(5) & _
                                      pfstr_Alineado(Format(rsFormaPago("IMP_SIN_REDONDEO"), "#0.00"), 28) & Space(5) & _
                                      "T/C S/. : " & pfstr_Alineado(Format(rsFormaPago("IMP_TIPO_CAMBIO"), "#,###,##0.00"), 15) & Space(5) & _
                                      "IMPORTE  S/. : " & pfstr_Alineado(Format(rsFormaPago("IMP_MONEDA_NAC"), "#,###,##0.00"), 15)
                        Else 'Otras formas de pago
                            Printer.Print pfstr_Alineado(rsFormaPago("DES_FORMA_PAGO"), 10) & Space(5) & _
                                      pfstr_Alineado(rsFormaPago("DES_HIJO"), 15) & Space(5) & pfstr_Alineado(Format(rsFormaPago("IMP_MONEDA_NAC"), "#0.00"), 5)
                        End If
                     End If
                    
                    'If rsFormaPago("IMP_MONEDA_NAC").Value <> rsFormaPago("IMP_VUELTO").Value Then
                        dblSumaTotalPagar = dblSumaTotalPagar + rsFormaPago("IMP_MONEDA_NAC").Value
                        dblSumaTotalVuelto = dblSumaTotalVuelto + rsFormaPago("IMP_VUELTO").Value
                    'End If
                End If
                rsFormaPago.MoveNext
            Wend
            
            'Print #1, String(42, "_") 'Separacion
            
'            For i = 1 To TotItems
'                Print #1, ""
'            Next i
            
            If "" & odynCab("COD_CONVENIO") = "" Then
                Dim objConvenio As New clsConvenio
                Dim rsDatosAdicionales As oraDynaset
                Set rsDatosAdicionales = objConvenio.ListaDocVerif("" & odynCab("COD_CONVENIO"))
                While Not rsDatosAdicionales.EOF
                    Printer.Print pfstr_Alineado(rsFormaPago("Descripcion"), 30) & Space(5) & pfstr_Alineado(rsFormaPago("FLG_RETENCION"), 30)
                    rsDatosAdicionales.MoveNext
                Wend
                Set objConvenio = Nothing
            Else
                dblSumaSubTotal = dblSumaSubTotal * (val(odynCab("PCT_BENEFICIARIO")) / 100)
            End If

            'Printer.Print "detalle aqui"
              
            'Total documento
            If vflgDetalle = True Then
                Printer.Print pfstr_Alineado(String(22, "-"), 42, Derecha)
                Printer.Print Space(1) & "Total de Registro => " & " " & intItems
                Printer.Print , Space(1) & "Total a Cancelar S/. " & " " & pfstr_Alineado(right(Space(5) & Format(dblSumaSubTotal, "#,###,##0.00"), 15), 15, Derecha)
                Printer.Print , Space(1) & "    Total Pagado S/. " & " " & pfstr_Alineado(right(Space(5) & Format(dblSumaTotalPagar, "#,###,##0.00"), 15), 15, Derecha)
                Printer.Print Space(1) & "          Vuelto S/. " & " " & pfstr_Alineado(right(Space(5) & Format(dblSumaTotalVuelto, "#,###,##0.00"), 15), 15, Derecha)
                'Print #1, Space(1) & "T.C. :" & " " & pfstr_Alineado(odynCab("IMP_TIP_CAMBIO").Value, 8)
                Printer.Print ""
            End If

            If odynCab("FLG_ENTREGA_TERCERO").Value = "1" Then
                Printer.Print Space(2) & "Datos de Entrega"
                Printer.Print Space(2) & String(18, "-")
                Printer.Print ""
                Printer.Print "Nombre:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_NOMBRE").Value, 30)
                Printer.Print "Direccion:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_DIRECC").Value, 50)
                Printer.Print "Telefono:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_TLF").Value, 15)
                Printer.Print "Referencia:" & Space(1) & pfstr_Alineado(odynCab("DES_AUX_RECOGE_REF").Value, 50)
                Printer.Print ""
            Else
                Printer.Print ""
            End If
        'Wend
        
        Printer.Print pfstr_Alineado("*** Este documento no es un comprobante de", 42, centro)
        Printer.Print pfstr_Alineado("pago. Solicite su comprobante de pago  ***", 42, centro)
        
        
        intCopia = intCopia + 1

        For j = 1 To 5
          Printer.Print ""
        Next j

        'Copia = Copia + 1
        '** Corte Parcial **'
        
        'Print #1, Chr(27) + "d" + "1"
        
        '** Corte Total **'
        'Print #1, Chr$(27); "d"; Chr$(2)

        '*********************************************************************************************************
        'Close #1
        Printer.EndDoc
        
    Loop
    On Error GoTo 0
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
         "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
         "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
    Else
      Exit Sub
    End If
End Sub



'*******************************************************************************
'Lista la cabecera de una proforma
'*******************************************************************************
Public Function ListaCabecera(ByVal Cia As String, ByVal vstrCodBtl As String, ByVal vstrNumProf As String) As oraDynaset
    On Error GoTo CntrlError
        
    Set ListaCabecera = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_CAB", 0, Cia, vstrCodBtl, vstrNumProf)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function
Public Function ListaDetalle(ByVal Cia As String, ByVal vstrCodBtl As String, ByVal vstrNumProf As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaDetalle = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DET", 0, Cia, vstrCodBtl, vstrNumProf)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function
'I.ECASTILLO 27.10.2020
Public Function ListaDetalle_V2(ByVal Cia As String, ByVal vstrCodBtl As String, ByVal vstrNumProf As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaDetalle_V2 = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DET_V2", 0, Cia, vstrCodBtl, vstrNumProf)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma.ListaDetalle_V2", Err.Description
End Function
'F.ECASTILLO 27.10.2020

Public Function ListaDetalleInsumosRM(ByVal Cia As String, ByVal vstrCodBtl As String, ByVal vstrNumProf As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaDetalleInsumosRM = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DET_INS_RM", 0, Cia, vstrCodBtl, vstrNumProf)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function


''Public Function ListaFormaPago(ByVal Cia As String, ByVal vstrCodBtl As String, ByVal vstrNumProf As String) As OraDynaset
''    On Error GoTo CntrlError
''    Set ListaFormaPago = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_FPAGO", 0, Cia, vstrCodBtl, vstrNumProf)
''    Exit Function
''CntrlError:
''    Err.Raise Err.Number, "clsProforma", Err.Description
''End Function
''

Public Function GrabarReclamo(ByVal Cia As String, _
                      ByVal CodigoLocal As String, _
                      ByVal NumeroProforma As String, _
                      ByVal Observacion As String, _
                      ByVal usuario As String _
            ) As String
Dim varValores As Variant
Dim varIO As Variant
On Error GoTo CntrlError
    If gstrFlagReclamo = "1" Then
        varValores = Array(Cia, _
          CodigoLocal, _
          NumeroProforma, _
          Observacion, _
          usuario)
                                        
        varIO = Array( _
            entrada, _
          entrada, _
          entrada, _
          entrada, _
          entrada)
    Else
        varValores = Array(Cia, _
          CodigoLocal, _
          NumeroProforma, _
          Observacion)
                                        
        varIO = Array( _
            entrada, _
          entrada, _
          entrada, _
          entrada)
    End If
                                   
    GrabarReclamo = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_GRABA_RECLAMO", varValores, varIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Property Get PedidoGenerado() As String
    PedidoGenerado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_GENERADO")
End Property

Public Property Get PedidoVerificado() As String
    PedidoVerificado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_VERIFICADO")
End Property

Public Property Get PedidoAsignado() As String
    PedidoAsignado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_ASIGNADO")
End Property

Public Property Get PedidoAvisado() As String
    PedidoAvisado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_AVISADO")
End Property

Public Property Get PedidoProforma() As String
    PedidoProforma = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_PROFORMA")
End Property

Public Property Get PedidoLlevando() As String
    PedidoLlevando = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_LLEVANDO")
End Property

Public Property Get PedidoLlegada() As String
    PedidoLlegada = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_LLEGADA")
End Property

Public Property Get PedidoEntregado() As String
    PedidoEntregado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_ENTREGADO")
End Property

Public Property Get PedidoLlegadaLocal() As String
    PedidoLlegadaLocal = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_LLEGADA_LOCAL")
End Property

Public Property Get PedidoLiberado() As String
    PedidoLiberado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_LIBERADO")
End Property

Public Property Get PedidoAnulado() As String
    PedidoAnulado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_ANULADO")
End Property

Public Function intFnUF(ByVal intUnidades As Integer, _
                    ByVal intFraciones As Integer, _
                    ByVal intCtdFraccionamiento As Integer, _
                    ByVal strUF As String) As Integer
    On Error GoTo CntrlError
    intFnUF = gclsOracle.FN_Valor("CMR.PKG_UTIL.FN_UF", intUnidades, intFraciones, intCtdFraccionamiento, strUF)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma.intFnUF", Err.Description
End Function

Public Function ListaUltimoPedidos(ByVal vstrCia As String, _
                                  ByVal vstrCod_Local As String, _
                                  ByVal vstrCod_Cliente As String, _
                                  ByVal vstrCod_Producto As String, _
                                  ByVal vstrNum_Meses As Integer) As oraDynaset
    On Error GoTo CntrlError
    Set ListaUltimoPedidos = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_ULT_PEDIDOS", 0, vstrCia, vstrCod_Local, vstrCod_Cliente, vstrCod_Producto, vstrNum_Meses)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Function ActualizaMotorizado_Proforma(ByVal Cia As String, _
                      ByVal CodigoLocal As String, _
                      ByVal NumeroPedido As String, _
                      ByVal CodigoMotorizado As String _
                      ) As String
Dim gvarValores  As Variant
Dim gvarIO  As Variant
On Error GoTo CntrlError
    
    gvarValores = Array(Cia, CodigoLocal, NumeroPedido, CodigoMotorizado)
    gvarIO = Array(entrada, entrada, entrada, entrada)
    ActualizaMotorizado_Proforma = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_ACTUALIZA_MOTORIZADO", gvarValores, gvarIO)
Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma.ActualizaMotorizado_Proforma", Err.Description
End Function

Public Function Documento_x_Anular(ByVal vstrCia As String, _
                                   ByVal vstrCodLocal As String) As oraDynaset
    On Error GoTo CntrlError
    Set Documento_x_Anular = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_PED_LIB_PEN_ANU", 0, vstrCia, _
                                                                                                vstrCodLocal)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma.Documento_x_Anular", Err.Description
End Function

Public Function Dev_RucProv_RecMag_x_Btl(ByVal vstrCodLocal As String) As String
    On Error GoTo CntrlError
    Dev_RucProv_RecMag_x_Btl = gclsOracle.FN_Valor("BTLPROD.PKG_PROFORMA_V3.DEV_RUCPROV_RECMAG_X_BTL", vstrCodLocal)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma.Dev_RucProv_RecMag_x_Btl", Err.Description
End Function

Public Function ListaPedidoDLVReimp(ByVal vstrCia As String, _
                               ByVal vstrCodLocal As String, _
                               ByVal vstrFchIni As String, _
                               ByVal vstrFchFin As String, _
                               Optional ByVal vstrNumProforma As String = "", _
                               Optional ByVal vstrDesDireccion As String = "") As oraDynaset

    On Error GoTo CtrlErr
    Set ListaPedidoDLVReimp = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.LISTA_PEDIDO_DLV_REIMP", 0, vstrCia, vstrCodLocal, vstrFchIni, vstrFchFin, vstrNumProforma, vstrDesDireccion)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDelivery", Err.Description
End Function

Public Function fndevNomRuteador(ByVal vstrCia As String, _
                                 ByVal vstrProforma As String) As String
    On Error GoTo CtrlErr
    fndevNomRuteador = gclsOracle.FN_Valor("BTLPROD.PKG_PROFORMA_V3.FN_DEV_NOMB_RUTEADOR", vstrCia, vstrProforma)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsProforma.fndevNomRuteador", Err.Description
End Function

Public Function GrabarLocalDespacho(ByVal Cia As String, _
                                    ByVal NumeroProforma As String, _
                                    ByVal CodigoLocalRef As String) As String
    Dim varValores As Variant
    Dim varIO As Variant
    On Error GoTo CntrlError
        varValores = Array(Cia, _
                           NumeroProforma, _
                           CodigoLocalRef)
                    
        varIO = Array(entrada, _
                      entrada, _
                      entrada)
                                   
    GrabarLocalDespacho = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_ACT_LOCAL_DESPACHO", varValores, varIO)

        Exit Function
CntrlError:
        Err.Raise Err.Number, "clsProforma", Err.Description
End Function

Public Sub EliminaTra(ByVal Cia As String, _
                        ByVal CodLocal As String, _
                        ByVal NumProforma As String)

    Dim varValores As Variant
    Dim varIO As Variant
    Dim Resultado As Variant
    
    On Error GoTo CntrlError
        varValores = Array(Cia, _
                           CodLocal, _
                           NumProforma)
                    
        varIO = Array(entrada, _
                      entrada, _
                      entrada)
                                   
    Resultado = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_ELIMINA_TRA", varValores, varIO)

        Exit Sub
CntrlError:
        Err.Raise Err.Number, "clsProforma.EliminaTra", Err.Description
End Sub

Public Sub ActObsRuta(ByVal Cia As String, _
                    ByVal CodLocal As String, _
                    ByVal NumProforma As String, _
                    ByVal obsLocal As String, _
                    ByVal obsRuta As String, _
                    ByVal obsMotorizado As String)
                    
Dim varValores As Variant
Dim varIO As Variant
Dim Resultado As Variant
                    
On Error GoTo handle
                    
        varValores = Array(Cia, _
                           CodLocal, _
                           NumProforma, _
                           obsLocal, _
                           obsRuta, _
                           obsMotorizado)
                    
        varIO = Array(entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada)
                    
    Resultado = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_ACT_OBS_RUTA", varValores, varIO)

Exit Sub

handle:
    Err.Raise Err.Number, "clsProforma.ActObsRuta", Err.Description

End Sub

Public Function ListaTransfPedido(ByVal vstrCia As String, _
                                  ByVal vstrProforma As String, _
                                  ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaTransfPedido = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.LISTA_TRANSFERENCIA_PEDIDO", 0, vstrCia, vstrProforma, CodigoLocal)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsProforma.ListaTransfPedido", Err.Description
End Function

Public Function ListaDeliveryProvincia(ByVal Cia As String, _
                                    ByVal Estado As String, _
                                    ByVal CodLocalRef As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaDeliveryProvincia = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DELIVERY_PROVINCIA", 0, Cia, Estado, CodLocalRef)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProforma.ListaDeliveryProvincia", Err.Description
End Function

Public Function ActualizaCliente(ByVal Cia As String, _
                            ByVal CodLocal As String, _
                            ByVal NumProforma As String) As String
                    
Dim varValores As Variant
Dim varIO As Variant
Dim Resultado As Variant
                    
On Error GoTo handle
                    
        varValores = Array(Cia, _
                           CodLocal, _
                           NumProforma)
                    
        varIO = Array(entrada, _
                      entrada, _
                      entrada)
                    
    ActualizaCliente = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_ACT_CLIENTE", varValores, varIO)

Exit Function

handle:
    Err.Raise Err.Number, "clsProforma.ActualizaCliente", Err.Description

End Function

'Public Function Cliente_Sufijo(ByVal vstrCodSufijo As String) As String
'    On Error GoTo CtrlErr
'    Set Cliente_Sufijo = gclsOracle.FN_Valor("BTLPROD.PKG_PROFORMA_V3.FN_CLIENTE_SUFIJO", vstrCodSufijo)
'
'    Exit Function
'CtrlErr:
'    Err.Raise Err.Number, "clsProforma.Cliente_Sufijo", Err.Description
'End Function

Property Get ClienteSufijo(ByVal vstrCodSuFijo As String)
    ClienteSufijo = gclsOracle.FN_Valor("BTLPROD.PKG_PROFORMA_V3.FN_CLIENTE_SUFIJO", vstrCodSuFijo)
End Property

Public Function Lista_DocsxProf(ByVal vstrCia As String, _
                                ByVal vstrCodLocal As String, _
                                ByVal vstrNumProforma As String) As oraDynaset
    On Error GoTo CtrlErr
    Set Lista_DocsxProf = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DOCS_X_PROF", 0, vstrCia, vstrCodLocal, vstrNumProforma)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsProforma.Lista_DocsxProf", vbCritical, Err.Description
End Function

Public Function ListaUlt_Diez_Pedidos(ByVal vstrCodCliente As String, _
                                      ByVal vstrNroPedido As String) As oraDynaset

    On Error GoTo CtrlErr
    Set ListaUlt_Diez_Pedidos = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_ULTIMOS_DIEZ_PEDIDOS", 0, vstrCodCliente, vstrNroPedido)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsProforma.ListaUlt_Diez_Pedidos", Err.Description

End Function

'Devuelve la cantidad de impresoras por máquina
'mlaguna 06/04/2010

Public Function Cuenta_Impresoras_x_Maquina(ByVal vstrCia As String, _
                                      ByVal vstrCodLocal As String) As Integer

    On Error GoTo CtrlErr
    Cuenta_Impresoras_x_Maquina = gclsOracle.FN_Valor("BTLPROD.PKG_PROFORMA_V3.fn_impresoras_x_maquina", vstrCia, vstrCodLocal)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsProforma.Cuenta_Impresoras_x_Maquina", Err.Description

End Function

Public Function DevuelveIPServer(ByVal CodigoLocal As String) As String
    On Error GoTo CtrlError
Dim StrCodigoZona As String
StrCodigoZona = "" & gclsOracle.FN_Valor("BTLPROD.FN_DEV_IP_FV", CodigoLocal)
DevuelveIPServer = StrCodigoZona
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsZona.DevuelveIPServer", Err.Description
End Function
'I.ECASTILLO 27.07.2020
Public Function getFlgReserva(ByVal Cia As String, ByVal CodLocal As String, ByVal NumProforma As String) As String
On Error GoTo CtrlError
    Dim strFlg As String
    strFlg = "" & gclsOracle.FN_Valor("BTLPROD.PKG_PROFORMA_V3.FN_GET_FLG_RESERVA", Cia, CodLocal, NumProforma)
    getFlgReserva = Trim(strFlg)
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsProforma.getFlgReserva", Err.Description
End Function
'F.ECASTILLO 27.07.2020
'I.ECASTILLO 27.07.2020
Public Function AnulaReservaStock(ByVal CodLocal As String, ByVal NumProforma As String, ByVal CodUsuario As String, Optional objProforma As Object) As Boolean
On Error GoTo Err
    Dim odynCab, odynTransf As oraDynaset
    Dim odynDet As oraDynaset
    Dim rsCia As oraDynaset
    Dim sCia As String
'    Dim arrInfo As New XArrayDB
    Dim arrProductos As New XArrayDB
    Dim obj As New Dictionary
    Dim x, i, j As Integer
    Dim Glosa As String
    Dim localSAP As String
    Dim sCodLoc As String
    Dim v_local, v_localSap, v_numProforma, v_cia As String
    Dim jsonProformas, jsonMaestro, jsonMaestroDetalle, message, Status As String
    Dim dicProformas As New Dictionary, dicMaestro As New Dictionary, dicErrores As New Dictionary
    Dim flgReservo As String
    Glosa = "LIBERACIÓN PEDIDO CALL"
    If objProforma Is Nothing Then
        'SOLO DEBE PROCESAR SI NO ES UNA PROFORMA TRANSFERENCIA.
        Set odynCab = ListaCabecera(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, NumProforma)
        Set odynDet = ListaDetalle(objUsuario.CodigoEmpresa, CodLocal, NumProforma)
        Set odynTransf = ListaTransfPedido(objUsuario.CodigoEmpresa, NumProforma, objUsuario.CodigoLocal)
        
        'OBTIENE CIA
        sCodLoc = odynCab("COD_LOCAL_REF").Value
        flgReservo = odynCab("FLG_RESERVA").Value
        If flgReservo <> "1" Then GoTo ok
        Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, sCodLoc)
        If (rsCia.RecordCount > 0) Then
          sCia = CStr(rsCia(1))
        End If
        Set rsCia = Nothing
        localSAP = objLocal.GetCodPosu(sCodLoc)
'        estConfig = objLocal.GetEstConfig(sCia, sCodLoc, "RESERVA_STOCK")
        '---------------------------------------
        Dim jsonProformasDetalle As String
        jsonProformasDetalle = procesaOdnyDet(odynDet, sCia, localSAP)
        If detalleProductoError = "1" Then MsgBox "El servicio no retorno algunos productos.", vbOKOnly + vbCritical, App.ProductName & " - Error al liberar reserva de stock": GoTo FAIL
        '*******************ARMA JSON MAESTRO*******************
        
        jsonProformas = "{'cia':'" & sCia & "','numProforma':'" & NumProforma & "','codLocal':'" & localSAP & "'," & _
                        "'codLocalBtl':'" & sCodLoc & "','posu':'" & objLocal.EvaluaLocalInkMig(sCodLoc) & "','detalle': [" & jsonProformasDetalle & "]},"
        Debug.Print jsonProformas
        odynTransf.MoveFirst
        For x = 1 To odynTransf.RecordCount
        
            v_local = "": v_numProforma = ""
            v_local = odynTransf("COD_LOCAL").Value
            Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, v_local)
            If (rsCia.RecordCount > 0) Then
              v_cia = CStr(rsCia(1))
            End If
            Set rsCia = Nothing
            v_localSap = objLocal.GetCodPosu(v_local)
            v_numProforma = odynTransf("NUM_PROFORMA").Value
            Set odynDet = Nothing
            Set odynDet = ListaDetalle(objUsuario.CodigoEmpresa, v_local, v_numProforma)
            jsonProformasDetalle = ""
            jsonProformasDetalle = procesaOdnyDet(odynDet, v_cia, v_localSap)
            If detalleProductoError = "1" Then MsgBox "El servicio no retorno algunos productos.", vbOKOnly + vbCritical, App.ProductName & " - Error al liberar reserva de stock": GoTo FAIL
            jsonProformas = jsonProformas & "{'cia':'" & v_cia & "','numProforma':'" & NumProforma & "'," & _
                                            "'codLocal':'" & v_localSap & "','codLocalBtl':'" & v_local & "'," & _
                                            "'posu':'" & objLocal.EvaluaLocalInkMig(v_local) & "','detalle': [" & jsonProformasDetalle & "]},"
            odynTransf.MoveNext
        Next x
        jsonProformas = left(jsonProformas, Len(jsonProformas) - 1)
        jsonMaestro = "{'data':[" & jsonProformas & "]}"
        Debug.Print jsonMaestro
        Set dicMaestro = JsonConverter.ParseJson(jsonMaestro)
        jsonProformas = ""
    Else
        Debug.Print objProforma("data").Count()
        Set dicMaestro = objProforma
    End If
    
    '---------------------------------------
    Debug.Print dicMaestro("data").Count
    If dicMaestro("data").Count > 0 Then
        x = 1
        For x = 1 To dicMaestro("data").Count
            'si local es inkaventa consume LiberaciondeStock
            'caso contrario AnulaReservadeStock
            'N si es inka y no migro a posu
            Dim example As String
            Set obj = Nothing
            arrProductos.ReDim 0, -1, 0, 3
            i = 1
            For i = 1 To dicMaestro("data")(x)("detalle").Count()
                arrProductos.AppendRows
                arrProductos(i - 1, 0) = dicMaestro("data")(x)("detalle")(i)("codProducto")
                arrProductos(i - 1, 1) = dicMaestro("data")(x)("detalle")(i)("quantity")
                arrProductos(i - 1, 2) = dicMaestro("data")(x)("detalle")(i)("isFractional")
                arrProductos(i - 1, 3) = dicMaestro("data")(x)("detalle")(i)("unitQuantity")
            Next i
'            sCia = objProforma("data")(i)("cia"): localSAP = objProforma("data")(i)("codLocal"): strProforma = objProforma("data")(i)("numProforma")
'            example = "": message = "": status = "": jsonProformasDetalle = ""
            v_numProforma = dicMaestro("data")(x)("numProforma")
            v_localSap = dicMaestro("data")(x)("codLocal")
            'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
            Dim strCia As String
            strCia = "" & objLocal.GetMarcaLocalPosu(v_localSap, 1)
            strCia = Trim(strCia)
            'F.ECASTILLO 17.09.2021
            If dicMaestro("data")(x)("posu") = "N" Then
                'Si es InkaVenta enviar codigo inka
                v_localSap = ""
                v_localSap = objLocal.GetCodInka(dicMaestro("data")(x)("codLocalBtl"))
                Set obj = objWS.LiberaciondeStock("INKV", sCia, v_localSap, Glosa, v_numProforma, gvarUSUARIO, arrProductos, strCia)
'                example = ""
'                example = "{'data':{'code':'SUCCESS','reserveList':[{'productId':'515335','quantity':3,'status':'0'}]}}"
'''                example = "{'data':{'code':'ERROR_INSINK_NOT_FOUND_LOCAL_ID','message':'LocalIDnotFoundinExternalDatabaseParametersTable','reserveList':null}}"
''                example = "{'data':{'code':'SUCCESS','reserveList':[{'productId':'515335','quantity':2,'status':'1403','description':'ORA-01403:nodatafound\nORA-06512:at\'PTOVENTA.DIGITAL_CONSUMO\',line44\nORA-06512:atline1\n'}]}}"
'                Set obj = JsonConverter.ParseJson(example)
            Else
                Set obj = objWS.AnulaReservaStock(sCia, v_numProforma, Glosa, v_localSap, arrProductos, gvarUSUARIO, strCia)
'                example = ""
'                example = "{'data':{'code':'SUCCESS','reserveList':[{'productId':'515335','quantity':3,'status':'0'}]}}"
'''                example = "{'data':{'code':'ERROR_INSINK_NOT_FOUND_LOCAL_ID','message':'LocalIDnotFoundinExternalDatabaseParametersTable','reserveList':null}}"
'''                example = "{'data':{'code':'SUCCESS','reserveList':[{'productId':'515335','quantity':2,'status':'1403','description':'ORA-01403:nodatafound\nORA-06512:at\'PTOVENTA.DIGITAL_CONSUMO\',line44\nORA-06512:atline1\n'}]}}"
'                Set obj = JsonConverter.ParseJson(example)
            End If
            Dim json_error As String
            Dim json_succes As New Dictionary
            
            Debug.Print obj.Count
            If obj.Count > 0 Then 'hubo respuesta
                If obj("data").Count() > 0 Then
                    'validar si logro anular la reserva de todos los productos sino error
                    'si code es SUCCESS y en reserveList existe status con valor 1403 = ERROR
                    'si code es ERROR_INSINK_NOT_FOUND_LOCAL_ID y reserveList es null = ERROR
                    'caso contrario OK
                    '{'data':{'code':'SUCCESS','reserveList':[{'productId':'515335','quantity':2,'status':'0'},{'productId':'515335','quantity':2,'status':'0'}]}}
                    If obj("data")("code") = "SUCCESS" Then
                        Debug.Print obj("data")("reserveList").Count
                        If obj("data")("reserveList").Count() > 0 Then
                            Dim Detalle As String
                            Dim bool_error As Boolean
                            'considerando que el servicio anula todo o nada, entonces [POR LOCAL]
                            'basta con que 1 de X producto no se anule para que interrumpa la anulación en rac
                            'y se vuelvan a reservar los productos que si se anularon
                            i = 1
                            While i <= obj("data")("reserveList").Count() And bool_error = False
                                Debug.Print obj("data")("reserveList")(i)("status")
                                If obj("data")("reserveList")(i)("status") <> 0 Then bool_error = True
                                i = i + 1
                            Wend
                            
                            Dim pd As Object
                            Set pd = obj("data")("reserveList")
                            Debug.Print pd.Count()
                            If bool_error Then
                                Detalle = procesaDetalleError(obj("data")("reserveList"), dicMaestro("data")(x)("detalle"))
                                message = "Existen productos con errores"
                                Status = "500"
                            Else
                                message = "Anulación de reserva exitosa": Status = "200"
                            End If
                            Debug.Print Detalle
                            json_error = Detalle
'                        ElseIf obj("data")("reserveList").Count() = 0 Then
'                            GoTo FAIL
                        End If
                    ElseIf obj("data")("code") = "ERROR_INSINK_NOT_FOUND_LOCAL_ID" Then
                        message = "El local tiene problemas de homologación/configuración y ningun producto fue procesado": Status = "500"
                    ElseIf IsNull(obj("data")("code")) Then
                        message = obj("data")("message"): Status = 500
                    Else
'                        'SI ES EXITOSO SE ALMACENA EN UN ARRAY LA PROFORMA, LOCALES Y LOS PRODUCTOS PROCESADOS
'                        'PARA PODER REALIZAR UN "ROLLBACK" EN CASO DE QUE ALGUNO OTRO FALLE
                        message = obj("data")("message"): Status = 500
                    End If
                    Debug.Print obj("data")("message")
                    Debug.Print vbNewLine
                End If
                
            Else
                message = "Servicio no retorno respuesta": Status = 500
            End If
            Dim COMMA As String
            COMMA = ""
            If x <> dicMaestro("data").Count Then COMMA = ","
            jsonProformas = jsonProformas & "{" & _
                                            "'cia': '" & dicMaestro("data")(x)("cia") & "'," & _
                                            "'numProforma': '" & dicMaestro("data")(x)("numProforma") & "'," & _
                                            "'codLocal': '" & v_localSap & "'," & _
                                            "'message': '" & message & "'," & _
                                            "'status': '" & Status & "'," & _
                                            "'detalle': [" & json_error & "]}" & COMMA
        Next x
    End If
    jsonMaestro = "{'data':[" & jsonProformas & "]}"
    Debug.Print jsonMaestro
    Set dicErrores = JsonConverter.ParseJson(jsonMaestro)
    Debug.Print "jsonErrores: " & vbNewLine & dicErrores.Count & vbNewLine & jsonMaestro
    'si dicErrores("data")(x)("detalle").Count() > 0 y dicErrores("data")(x)("status") <> 200
    'entonces debe volver a reservar lo anulado
    x = 1
    Dim boolErr As Boolean
    While x <= dicErrores("data").Count() And boolErr = False
        If dicErrores("data")(x)("detalle").Count() > 0 Or dicErrores("data")(x)("status") <> 200 Then
            boolErr = True
        End If
        x = x + 1
    Wend
    Dim msgReserva As String
    Dim msgError As String
    Dim msgComma As String
    If objProforma Is Nothing Then
        If boolErr Then
            msgReserva = ReservaStock(2, dicMaestro, dicErrores)
            x = 1: msgError = ""
            msgError = "Ocurrieron inconvenientes al intentar liberar la reserva de stock: " & vbNewLine & "===========================================" & vbNewLine
            For x = 1 To dicErrores("data").Count()
                msgError = msgError & vbNewLine & dicErrores("data")(x)("numProforma") & "|" & dicErrores("data")(x)("codLocal") & "|" & dicErrores("data")(x)("message") & ":" & vbNewLine
                msgComma = ""
                If x <> dicErrores("data").Count() Then msgComma = ","
                i = 1
                For i = 1 To dicErrores("data")(x)("detalle").Count()
                    msgError = msgError & _
                               dicErrores("data")(x)("detalle")(i)("codProducto") & " " & _
                               dicErrores("data")(x)("detalle")(i)("desProducto") & vbNewLine
                Next i
            Next x
            MsgBox msgError & msgReserva, vbOKOnly + vbCritical, App.ProductName & " - Error al liberar reserva de stock"
            GoTo FAIL
        Else
            x = 1: msgError = ""
            For x = 1 To dicErrores("data").Count()
                msgError = msgError & vbNewLine & dicErrores("data")(x)("numProforma") & "|" & dicErrores("data")(x)("codLocal") & "|" & dicErrores("data")(x)("message")
            Next x
            MsgBox msgError, vbOKOnly + vbInformation, App.ProductName & " - Éxito al liberar reserva de stock"
            GoTo ok
        End If
    End If
    If boolErr Then GoTo FAIL
ok:
    AnulaReservaStock = True
    Exit Function
FAIL:
    
    AnulaReservaStock = False
    Exit Function
Err:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
End Function
'F.ECASTILLO 27.07.2020
'I.ECASTILLO 27.07.2020
Public Function procesaDetalle(ByVal reserveList As Object, ByVal Detalle As Object) As String
    Debug.Print Detalle.Count()
    Debug.Print reserveList.Count()
    Dim x, i, j As Integer
    Dim str As String
    Dim boolExiste As Boolean
    If reserveList.Count() > 0 Then
        x = 1
        For x = 1 To reserveList.Count()
            Debug.Print reserveList(1)("status")
            Dim COMMA, desProducto As String
            Dim flg_frac, ctd_frac As Integer
            flg_frac = 0: ctd_frac = 1: desProducto = "": COMMA = ""
            'If x <> reserveList.Count() Then COMMA = ","
            i = 1
            While i <= Detalle.Count() And boolExiste = False
                If reserveList(x)("productId") = Detalle(i)("codProducto") Then flg_frac = Detalle(i)("isFractional"): ctd_frac = Detalle(i)("unitQuantity"): boolExiste = True: desProducto = Replace(Detalle(i)("desProducto"), "'", "")
                i = i + 1
            Wend
            If boolExiste Then
                str = str & "{" & _
                                "'codProducto':'" & reserveList(x)("productId") & "'," & _
                                "'quantity':'" & reserveList(x)("quantity") & "'," & _
                                "'unitQuantity':'" & ctd_frac & "'," & _
                                "'isFractional':'" & flg_frac & "'," & _
                                "'status':'" & reserveList(x)("status") & "'" & _
                                "'desProducto':'" & Replace(desProducto, "'", "") & "'" & _
                            "}," '& COMMA
            End If
        Next x
    End If
    str = left(str, Len(str) - 1)
    procesaDetalle = str
End Function
'F.ECASTILLO 27.07.2020
'I.ECASTILLO 17.12.2020
Public Function procesaDetalleError(ByVal reserveList As Object, ByVal Detalle As Object) As String
    Debug.Print Detalle.Count()
    Debug.Print reserveList.Count()
    Dim x, i, j As Integer
    Dim str As String
    Dim boolExiste As Boolean
    If reserveList.Count() > 0 Then
        x = 1
        For x = 1 To reserveList.Count()
            Debug.Print reserveList(1)("status")
            Dim COMMA, desProducto As String
            Dim flg_frac, ctd_frac As Integer
            flg_frac = 0: ctd_frac = 1: desProducto = "": COMMA = "": boolExiste = False
            'If x <> reserveList.Count() Then COMMA = ","
            i = 1
            While i <= Detalle.Count() And boolExiste = False
                If reserveList(x)("productId") = Detalle(i)("codProducto") And reserveList(x)("status") <> 0 Then
                   flg_frac = Detalle(i)("isFractional")
                   ctd_frac = Detalle(i)("unitQuantity")
                   boolExiste = True
                   desProducto = Detalle(i)("desProducto")
                End If
                i = i + 1
            Wend
            If boolExiste Then
                str = str & "{" & _
                                "'codProducto':'" & reserveList(x)("productId") & "'," & _
                                "'quantity':'" & reserveList(x)("quantity") & "'," & _
                                "'unitQuantity':'" & ctd_frac & "'," & _
                                "'isFractional':'" & flg_frac & "'," & _
                                "'status':'" & reserveList(x)("status") & "'" & _
                                "'desProducto':'" & Replace(desProducto, "'", "") & "'" & _
                            "}," '& COMMA
            End If
        Next x
    End If
    str = left(str, Len(str) - 1)
    procesaDetalleError = str
End Function
'F.ECASTILLO 17.12.2020
'I.ECASTILLO 27.07.2020
Public Function procesaOdnyDet(ByVal ODyn As oraDynaset, ByVal Cia As String, ByVal CodLocal As String) As String
    Dim x, i, j As Integer
    Dim str As String
    '**************OBTIENE STOCK WEB SERVICE****************
    Dim obj As New Dictionary
    Dim ArrCode As Variant
    Dim codProductoPosu, codLocalPosu As String
    detalleProductoError = ""
    ReDim ArrCode(ODyn.RecordCount - 1)
    ODyn.MoveFirst
    For x = 0 To ODyn.RecordCount - 1
        ArrCode(x) = objProducto.GetCodPosu(ODyn("COD_PRODUCTO").Value)
        ODyn.MoveNext
    Next x
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    Dim sCia As String
    sCia = "" & objLocal.GetMarcaLocalPosu(CodLocal, 1)
    sCia = Trim(sCia)
    'F.ECASTILLO 17.09.2021
    Set obj = Nothing
    Set obj = objWS.GetStockRealWS(Cia, CodLocal, ArrCode, sCia)
    '******************ARMA JSON DETALE*********************
    ODyn.MoveFirst
    i = 1
    While Not ODyn.EOF
        Dim flg_frac, ctd_frac As Integer
        Dim codPosu, COMMA As String
        Dim boolExiste As Boolean
        Dim ora_ctd_ent, ora_ctd_frac
        Dim ws_ctd_quantity
        Dim isVtaFracc
        flg_frac = 0
        ctd_frac = 1
        COMMA = ""
        If i <> ODyn.RecordCount Then COMMA = ","
        codPosu = objProducto.GetCodPosu(ODyn("COD_PRODUCTO").Value)
        boolExiste = False
        If obj.Count > 0 Then
            Debug.Print obj("data").Count()
            For x = 1 To obj("data").Count()
                If obj("data")(x)("productId") = codPosu And boolExiste = False Then
                    If obj("data")(x)("isFractional") = 1 Then flg_frac = 1
                    If flg_frac = 1 Then
                        For j = 1 To obj("data")(x)("fractionType").Count()
                            If obj("data")(x)("fractionType")(j)("fractionatedText") = "PACK_MODE" Then
                                ctd_frac = obj("data")(x)("fractionType")(j)("unitQuantity") 'cnt_fraccionamiento
                            End If
                        Next j
                        'ctd_frac = obj("data")(X)("fractionType")(2)("unitQuantity")
                    Else
                        For j = 1 To obj("data")(x)("fractionType").Count()
                            If obj("data")(x)("fractionType")(j)("fractionatedText") = "PART_MODE" Then
                                ctd_frac = obj("data")(x)("fractionType")(j)("unitQuantity") 'cnt_fraccionamiento
                            End If
                        Next j
                        'ctd_frac = obj("data")(X)("fractionType")(1)("unitQuantity")
                    End If
                    boolExiste = True
                End If
            Next x
        Else
            detalleProductoError = "1"
            Exit Function
            'ctd_frac = "10": detalleProductoError = ""
        End If
        isVtaFracc = InStr(1, ODyn("CANT").Value, "F")
        ora_ctd_ent = IIf(isVtaFracc = 0, ODyn("CTD_PRODUCTO").Value, Int(ODyn("CTD_PRODUCTO").Value / ODyn("CTD_FRACCIONAMIENTO").Value))
        ora_ctd_frac = IIf(isVtaFracc = 0, 0, ODyn("CTD_PRODUCTO").Value - (ora_ctd_ent * ODyn("CTD_FRACCIONAMIENTO").Value))
        ws_ctd_quantity = ora_ctd_ent * ctd_frac + ora_ctd_frac
        str = str & "{"
        str = str & "'codProducto':'" & codPosu & "',"
        'str = str & "'quantity':'" & ODyn("CTD_PRODUCTO").Value & "',"
        str = str & "'quantity':'" & ws_ctd_quantity & "',"
        str = str & "'isFractional':'" & flg_frac & "',"
        str = str & "'unitQuantity':'" & ctd_frac & "',"
        Debug.Print ODyn("DES_PRODUCTO").Value
        str = str & "'desProducto':'" & Replace(ODyn("DES_PRODUCTO").Value, "'", "") & "'"
        str = str & "}"
        str = str & COMMA
        i = i + 1
        ODyn.MoveNext
    Wend
    Debug.Print str
    procesaOdnyDet = str
End Function
'F.ECASTILLO 27.07.2020
'I.ECASTILLO 27.07.2020
Public Function ReservaStock(ByVal Tipo As String, ByVal objProforma As Object, Optional objErrores As Object) As String
'ByVal strProforma As String, ByVal localId As String, ByVal Productos As XArrayDB
On Error GoTo handle
    Dim obj As New Dictionary
    Dim lk, i, x, j As Integer
    Dim Glosa, sCia, localSAP, strProforma As String
    Dim boolAnulo As Boolean
    Dim boolError As Boolean
    Dim arrProductos As New XArrayDB
    Dim msgReserva, COMMA, Detalle As String
    Dim jsonProformas, jsonProformasDetalle As String
    Dim dicErrores As New Dictionary
    Dim sinErrores As New Dictionary
    Dim example, message, Status As String
    Glosa = "RESERVA PEDIDO CALL"
    'buscar si proforma/local de objProforma se encuentra en objErrores
    'de ser verdadero entonces se envia a reservar
    '**SE MANDA A RESERVAR SOLO LAS PROFORMAS QUE NO TUVIERON ERROR (SI SE ANULARON)
    i = 1
    Debug.Print objProforma.Count()
    For i = 1 To objProforma("data").Count()
        COMMA = ""
        If i <> objProforma("data").Count() Then COMMA = ","
        Debug.Print objProforma("data")(i)("cia")
        boolAnulo = False
        j = 1
        If Not objErrores Is Nothing Then
            Debug.Print objErrores.Count()
            While j <= objErrores.Count() And boolAnulo = False
                If (objProforma("data")(i)("cia") = objErrores("data")(j)("cia")) _
                    And (objProforma("data")(i)("numProforma") = objErrores("data")(j)("numProforma")) _
                    And (objProforma("data")(i)("codLocal") = objErrores("data")(j)("codLocal")) _
                    And (objErrores("data")(j)("status") = 200) Then
                    boolAnulo = True
                End If
                j = j + 1
            Wend
        End If
        'tipo: 1 [reservaNormal], 2 [reservaAnulado]
        If (boolAnulo And Tipo = 2) Or Tipo = 1 Then
            arrProductos.ReDim 0, -1, 0, 3
            x = 1
            For x = 1 To objProforma("data")(x)("reserveList").Count()
                arrProductos.AppendRows
                arrProductos(x - 1, 0) = objProforma("data")(i)("reserveList")(x)("codProducto")
                arrProductos(x - 1, 1) = objProforma("data")(i)("reserveList")(x)("quantity")
                arrProductos(x - 1, 2) = objProforma("data")(i)("reserveList")(x)("isFractional")
                arrProductos(x - 1, 3) = objProforma("data")(i)("reserveList")(x)("unitQuantity")
            Next x
            sCia = objProforma("data")(i)("cia"): localSAP = objProforma("data")(i)("codLocal"): strProforma = objProforma("data")(i)("numProforma")
            'I.ECASTILLO 28.01.2021 | se reversa el cambio realizado por gnibin
'            'INI GNIBIN 20210127 Proyecto MultiMarca
'            Select Case sCia
'                Case "94", "93", "92", "1DLV"
'                    sCia = "1DLV"
'                Case Else
'                    sCia = "0DLV"
'            End Select
'            'FIN GNIBIN 20210127 Proyecto MultiMarca
            'F.ECASTILLO 28.01.2021
            'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
            Dim strCia As String
            strCia = "" & objLocal.GetMarcaLocalPosu(localSAP, 1)
            strCia = Trim(strCia)
            'F.ECASTILLO 17.09.2021
            example = "": message = "": Status = "": jsonProformasDetalle = ""
            If objProforma("data")(i)("posu") = "N" Then
                Set obj = objWS.ReservadeStock("INKV", sCia, localSAP, Glosa, strProforma, gvarUSUARIO, arrProductos, strCia)
            Else
                Set obj = objWS.ReservaStock(sCia, strProforma, Glosa, localSAP, arrProductos, gvarUSUARIO, strCia)
            End If
            If obj("data").Count() > 0 Then '26.01.2021
                j = 1: Detalle = "": boolError = False: jsonProformasDetalle = "": message = "": Status = ""
                If Not IsNull(obj("data")("reserveList")) Then
                    For j = 1 To obj("data")("reserveList").Count()
                        'si existe algun producto con error, se arma json con la proforma
                        Detalle = Detalle & "{'codProducto':'" & obj("data")("reserveList")(j)("productId") & "','status':'" & obj("data")("reserveList")(j)("status") & "'},"
                        If obj("data")("reserveList")(j)("status") <> 0 Then boolError = True: Exit For
                    Next j
                    If boolError Then
                        jsonProformasDetalle = procesaDetalleError(obj("data")("reserveList"), objProforma("data")(i)("reserveList"))
                        message = "Existen productos con errores"
                        Status = "500"
                    End If
                Else
                    message = obj("data")("message")
                End If
                If Len(jsonProformasDetalle) > 0 Or Len(message) > 0 Then
                    jsonProformas = jsonProformas & "{" & _
                                                "'cia': '" & sCia & "'," & _
                                                "'numProforma': '" & strProforma & "'," & _
                                                "'codLocal': '" & localSAP & "'," & _
                                                "'codLocalBtl': '" & objProforma("data")(i)("codLocalBtl") & "'," & _
                                                "'posu': '" & objProforma("data")(i)("posu") & "'," & _
                                                "'message': '" & Replace(message, "'", "") & "'," & _
                                                "'status': '" & Replace(Status, ",", "") & "'," & _
                                                "'detalle': [" & jsonProformasDetalle & "]},"
'                    jsonProformas = "{'cia':'" & sCia & "','numProforma':'" & NumProforma & "','codLocal':'" & localSAP & "'," & _
'                        "'codLocalBtl':'" & sCodLoc & "','posu':'" & objLocal.EvaluaLocalInkMig(sCodLoc) & "','detalle': [" & jsonProformasDetalle & "]}"
        
                End If
                If Len(Detalle) > 0 Then Detalle = left(Detalle, Len(Detalle) - 1)
                If Len(Detalle) > 0 Then msgReserva = msgReserva & "{'code':'" & obj("data")("code") & "','reserveList':[" & Detalle & "]},"
            End If
        End If
    Next i
    If Len(msgReserva) > 0 Then msgReserva = left(msgReserva, Len(msgReserva) - 1)
    If Len(jsonProformas) > 0 Then
        jsonProformas = left(jsonProformas, Len(jsonProformas) - 1)
        jsonProformas = "{'data':[" & jsonProformas & "]}"
        Debug.Print jsonProformas
        Set dicErrores = JsonConverter.ParseJson(jsonProformas) 'proformas con errores
        jsonProformas = ""
    End If
    x = 1
    Dim boolErr As Boolean
    Dim msgError As String
    Dim msgComma As String
    Dim jsonSinError As String
    Debug.Print dicErrores.Count()
    If Tipo = 1 Then
        If dicErrores.Count() > 0 Then
            x = 1: msgError = "": Debug.Print dicErrores("data").Count()
            msgError = "Ocurrieron inconvenientes al intentar reservar stock: " & vbNewLine & "==========================================="
            
            For x = 1 To dicErrores("data").Count()
                msgError = msgError & vbNewLine & dicErrores("data")(x)("numProforma") & "|" & dicErrores("data")(x)("codLocal") & "|" & dicErrores("data")(x)("message") & ":" & vbNewLine
                i = 1
                For i = 1 To dicErrores("data")(x)("detalle").Count()
                    msgError = msgError & _
                               dicErrores("data")(x)("detalle")(i)("desProducto") & vbNewLine
                Next i
            Next x
            MsgBox msgError & msgReserva, vbOKOnly + vbCritical, App.ProductName & " - Error al reservar stock"
            jsonSinError = filtraProformasError(objProforma, dicErrores)
            Debug.Print jsonSinError
            Set sinErrores = JsonConverter.ParseJson(jsonSinError)
            GoTo ErrorReserva
        Else
            x = 1: msgError = ""
            For x = 1 To objProforma("data").Count()
                msgError = msgError & vbNewLine & objProforma("data")(x)("numProforma") & "|" & objProforma("data")(x)("codLocal") & "| Reserva exitosa"
            Next x
            MsgBox msgError, vbOKOnly + vbInformation, App.ProductName & " - Éxito al reservar stock"
            GoTo ok
        End If
    End If
    If Len(msgReserva) > 0 Then msgReserva = left(msgReserva, Len(msgReserva) - 1)
    ReservaStock = IIf(Len(msgReserva) > 0, "{'data':[" & msgReserva & "]}", "")
    Exit Function
ErrorReserva:
    'Enviar anular solo proformas sin error
    If AnulaReservaStock("", "", "", sinErrores) = False Then
        ReservaStock = "Error"
        Exit Function
    End If
    ReservaStock = "Error"
    Exit Function
ok:
    Exit Function
handle:
    ReservaStock = "Error"
    If Not IsNull(Err) Then MsgBox Err.Description, vbCritical, App.ProductName
End Function
'F.ECASTILLO 27.07.2020
'I.ECASTILLO 27.07.2020
Public Function filtraProformasError(ByVal proformas As Object, ByVal errores As Object) As String
    Dim newJson As String
    Dim Detalle As String
    Dim boolError As Boolean
    Dim i, j, x As Integer
    Debug.Print proformas.Count()
    Debug.Print errores.Count()
    newJson = "": boolError = False: i = 0: j = 0: x = 0
    If proformas.Count() > 0 Then
        If proformas("data").Count() > 0 Then
            Debug.Print (proformas("data").Count())
            For x = 1 To proformas("data").Count()
                Debug.Print (proformas("data")(x)("codLocal"))
                If errores.Count() > 0 Then
                    If errores("data").Count() > 0 Then
                        i = 1
                        For i = 1 To errores("data").Count()
                            If (proformas("data")(x)("numProforma") = errores("data")(i)("numProforma")) _
                                And (proformas("data")(x)("codLocalBtl") = errores("data")(i)("codLocalBtl")) Then boolError = True: Exit For
                        Next i
                    End If
                End If
                j = 1: Detalle = ""
                For j = 1 To proformas("data")(x)("reserveList").Count()
                    Detalle = "{'codProducto':'" & proformas("data")(x)("reserveList")(j)("codProducto") & "'," & _
                              "'quantity':'" & proformas("data")(x)("reserveList")(j)("quantity") & "'," & _
                              "'isFractional':'" & proformas("data")(x)("reserveList")(j)("isFractional") & "'," & _
                              "'unitQuantity':'" & proformas("data")(x)("reserveList")(j)("unitQuantity") & "'," & _
                              "'desProducto':'" & Replace(proformas("data")(x)("reserveList")(j)("desProducto"), "'", "") & "'},"
                Next j
                If Len(Detalle) > 0 Then Detalle = left(Detalle, Len(Detalle) - 1)
                If boolError = False Then
                    newJson = "{" & _
                              "'cia':'" & proformas("data")(x)("cia") & "'," & _
                              "'numProforma':'" & proformas("data")(x)("numProforma") & "'," & _
                              "'codLocal':'" & proformas("data")(x)("codLocal") & "'," & _
                              "'codLocalBtl':'" & proformas("data")(x)("codLocalBtl") & "'," & _
                              "'posu':'" & proformas("data")(x)("posu") & "'," & _
                              "'detalle':[" & Detalle & "]},"
                End If
            Next x
        End If
    End If
    If Len(newJson) > 0 Then newJson = left(newJson, Len(newJson) - 1)
    filtraProformasError = "{'data':[" & newJson & "]}"
End Function
'F.ECASTILLO 27.07.2020
Public Function ReviertePedido(ByVal A_CIA As String, _
                               ByVal A_COD_LOCAL As String, _
                               ByVal A_NUM_PROFORMA As String, _
                               ByVal A_COD_USUARIO As String, _
                               ByVal A_FLG_COMMIT As String) As String
Dim varValores As Variant
Dim varIO As Variant
                    
On Error GoTo handle
                    
        varValores = Array(A_CIA, _
                           A_COD_LOCAL, _
                           A_NUM_PROFORMA, _
                           A_COD_USUARIO, _
                           A_FLG_COMMIT)
                    
        varIO = Array(entrada, _
                      entrada, _
                      entrada, _
                      entrada, _
                      entrada)
                    
    ReviertePedido = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_REVIERTE_PEDIDO", varValores, varIO)

Exit Function

handle:
    Err.Raise Err.Number, "clsProforma.ActualizaCliente", Err.Description

End Function
Public Function revierteReservaStock(ByVal A_CIA As String, _
                                     ByVal A_COD_LOCAL As String, _
                                     ByVal A_NUM_PROFORMA As String, _
                                     ByVal A_FLG_RESERVA As String, _
                                     ByVal A_FLG_COMMIT As String) As String
Dim varValores As Variant
Dim varIO As Variant
On Error GoTo handle
    varValores = Array(A_CIA, A_COD_LOCAL, A_NUM_PROFORMA, A_FLG_RESERVA, A_FLG_COMMIT)
    varIO = Array(entrada, entrada, entrada, entrada, entrada)
    revierteReservaStock = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_REVIERTE_RESERVA", varValores, varIO)
    Exit Function
handle:
    Err.Raise Err.Number, "clsProforma.revierteReservaStock", Err.Description
End Function
'I.ECASTILLO 27.10.2020
Function reservaCapacidad(ByVal CodLocal As String, ByVal Fecha As String, ByVal hora As String, ByVal Tipo As String, ByVal strProforma As String) As String
    Dim resp As Dictionary
    reservaCapacidad = "0"
    If objVenta.bk_chkRET = "1" Then reservaCapacidad = "0": Exit Function
    If objVenta.flgReservoCap = "1" Then reservaCapacidad = "1": Exit Function
    Set resp = objWS.reservaCapacidades(CodLocal, _
                                        Fecha, _
                                        hora, _
                                        Tipo, _
                                        strProforma)
    If Not resp Is Nothing Then
        If IsObject(resp("data")) = False Then
            Exit Function
        End If
        If resp("data")("processStatus") = True Then
            reservaCapacidad = "1"
            objVenta.flgReservoCap = "1"
        Else
            'msg error
            objVenta.flgReservoCap = "0"
        End If
    End If
End Function
Function retornaCapacidad(ByVal CodLocal As String, ByVal Fecha As String, ByVal hora As String) As String
    Dim resp As Dictionary
    retornaCapacidad = "0"
    Set resp = objWS.retornaCapacidades(CodLocal, _
                                        Fecha, _
                                        hora, _
                                        "EXP")
    If Not resp Is Nothing Then
        If IsObject(resp("data")) = False Then
            Exit Function
        End If
        If resp("data")("processStatus") = True Then
            retornaCapacidad = "1"
        Else
            'msg error
        End If
    End If
End Function
Function asignaDigital(ByVal Cia As String, ByVal CodLocal As String, ByVal NumProforma As String) As String
    Dim resp As Dictionary
    Dim odynCab As oraDynaset
    Dim odynFP As oraDynaset
    Dim odynFPV As oraDynaset
    Dim odynDet As oraDynaset
    Dim odynCliente As oraDynaset
    Dim productoArr As New XArrayDB
    Dim valeArr As New XArrayDB
    Dim objCliente As New clsCliente
    Dim x, i, j As Integer
    Dim proformaNumber As String
    Dim localCode As String
    Dim ciaCode As String
    Dim localCodeRef As String
    Dim dateCreated As String
    Dim dateConfirmada As String
    Dim processStartConfirm As String
    Dim Source As String
    Dim companyCode As String
    Dim creditCardProviderId As String
    Dim currencyId As String
    Dim discountApplied As String
    Dim deliveryType As String
    Dim amount As String
    Dim localCodeReference As String
    Dim user_dni As String
    Dim user_isInkaClub As String
    Dim user_phone As String
    Dim user_email As String
    Dim user_name As String
    Dim user_lastName As String
    Dim user_isAnonymous As String
    Dim paymentMethod_id As String
    Dim paymentMethod_name As String
    Dim receipt_ruc As String
    Dim receipt_companyName As String
    Dim receipt_companyAddress As String
    Dim receipt_receiptType_name As String
    Dim deliveryAddress_street As String
    Dim deliveryAddress_number As String
    Dim deliveryAddress_apartment As String
    Dim deliveryAddress_country As String
    Dim deliveryAddress_city As String
    Dim deliveryAddress_district As String
    Dim deliveryAddress_latitude As String
    Dim deliveryAddress_longitude As String
    Dim deliveryAddress_receiverName As String
    Dim paymentAmountDto_productsTotalCostNoDiscount As String
    Dim paymentAmountDto_grossPrice As String
    Dim paymentAmountDto_discountApplied As String
    Dim paymentAmountDto_productsTotalCost As String
    Dim paymentAmountDto_deliveryCost As String
    Dim paymentAmountDto_paidAmountByCoin As String
    'I.ECASTILLO 17.12.2020
    Dim varTemp_CurCabConv As oraDynaset
    Dim varTemp_CurDetConv As oraDynaset
    Dim varTemp_IndConvenio As String
    Dim varTemp_CodConvenio As String
    Dim varTemp_CodModalidadVenta As String
    Dim varTemp_CodCliConvenio As String
    Dim varTemp_PctBeneficiario As String
    Dim varTemp_CodCampo As String
    Dim varTemp_NombreCampo As String
    Dim varTemp_CodValorIn As String
    Dim varTemp_DescripcionCampo As String
    Dim agreementFlag As String
    Dim agreementCode As String
    Dim agreementCodCli As String
    Dim agreementBeneficiaryPct As String
    Dim agreementFields As New XArrayDB
    Dim argAdicionales As New XArrayDB
    Dim argDelivery As New XArrayDB
    Dim dataCli As String
    Dim nameOf As String
    Dim pctIGV As String
    Dim totalTax As String
    Dim flgCappa As String
    Dim observation As String
    Dim observationMotorized As String
    Dim flagMultiBrand As String
    Dim company As String
    Dim deliveryTime As String
    Dim dateConfirmadaEnd As String
    Dim deliveryAddress_notes As String
    Dim deliveryAddress_urbanization As String
    'F.ECASTILLO 17.12.2020
    asignaDigital = "0"
    Set odynCab = ListaCabecera(Cia, CodLocal, NumProforma)
    Set odynDet = ListaDetalle_V2(Cia, CodLocal, NumProforma)
    Set odynFP = ListaFormaPago_V2(Cia, CodLocal, NumProforma)
    Set odynFPV = ListaFormaPagoVale(Cia, CodLocal, NumProforma)
    Set odynCliente = objCliente.Lista_v2(odynCab("COD_CLIENTE_DLV").Value)
    
    Dim sCia As String
    Dim rsCia As oraDynaset
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, CodLocal)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, CodLocal, "RESERVA_STOCK_2DA")
    Dim fchPactada As String
    fchPactada = "" & odynCab("FCH_HORA_PACT_ENTR").Value
    proformaNumber = "" & odynCab("NUM_PROFORMA").Value
    localCode = "" & odynCab("COD_LOCAL").Value
    ciaCode = "" & odynCab("CIA").Value
    localCodeRef = "" & odynCab("COD_LOCAL_SAP_REF").Value
    dateCreated = Format(odynCab("FCH_REGISTRA").Value, "yyyy-mm-dd") & "T" & odynCab("HORA").Value
    'si al grabar venta no se envia fch_pactada se setea el valor sysdate que es igual a fch_registra
    If Len(Trim(fchPactada)) = 0 Then
        dateConfirmada = "" & Format(odynCab("FCH_REGISTRA").Value, "yyyy-mm-ddTHH:mm:ss")
    Else
        dateConfirmada = "" & Format(odynCab("FCH_HORA_PACT_ENTR").Value, "yyyy-mm-ddTHH:mm:ss")
    End If
    deliveryTime = "" & odynCab("DELIVERY_TIME").Value
    deliveryTime = IIf(Len(Trim(deliveryTime)) = 0, "null", deliveryTime)
    dateConfirmadaEnd = "" & Format(odynCab("FCH_HORA_PACT_ENTR_FIN").Value, "yyyy-mm-ddTHH:mm:ss")
    processStartConfirm = "" & Format(odynCab("FCH_REGISTRA").Value, "yyyy-mm-dd") & "T" & odynCab("HORA").Value
    Dim isCappa As String
    isCappa = "" & objLocal.GetIndCDCAP(odynCab("COD_LOCAL_REF").Value)
    Source = "CALL" 'IIf(isCappa = "1", "CALL", "CALL_LOCALES")
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, odynCab("COD_LOCAL_REF").Value)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    company = "" & objLocal.GetMarcaLocal(odynCab("COD_LOCAL_REF").Value, 2)
    company = Trim(company)
    If Len(Trim(company)) = 0 Then
        Select Case sCia
            Case "94", "93", "92", "1DLV"
                company = "IKF"
            Case Else
                company = "MF"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    companyCode = IIf(odynCab("COD_LOCAL_CALL_CENTER").Value = "1DLV", "IKF", "MF")
    flagMultiBrand = IIf(company = companyCode, "false", "true")
    creditCardProviderId = "" & odynFP("COD_FORMA_PAGO_DIG").Value 'objVenta.dc_cod_forma_pago
    currencyId = "" & odynFP("CURRENCY_ID").Value
    creditCardProviderId = IIf(Len(Trim(creditCardProviderId)) = 0, "1", creditCardProviderId)
    currencyId = IIf(Len(Trim(currencyId)) = 0, "1", currencyId)
    
    discountApplied = "0.000"
    Dim FlgEntregaLocal As String
    FlgEntregaLocal = "" & odynCab("FLG_ENTREGA_LOCAL").Value
    If FlgEntregaLocal = "0" Then
        'deliveryType = "RAD"
        deliveryType = "" & odynCab("DELIVERY_TYPE").Value
    Else
        deliveryType = "RET"
    End If
    
    pctIGV = "" & objUsuario.IGV
    totalTax = "" & pctIGV
    amount = "" & odynFP("IMP_MONEDA_NAC").Value 'soles
    amount = IIf(Len(Trim(amount)) = 0, "0.000", amount)
    flgCappa = "" & IIf(isCappa = "1", "true", "false")
    'localCodeReference = odynCab("COD_LOCAL_SAP_REF").Value
    'I.ECASTILLO 17.12.2020
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    Dim flg_ruteoA_cnv
    flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto
    End If
    Dim flgFunLocal As String
    flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "1" And flgFunLocal = "1" Then
        GoTo CallServiceReserva
    End If
    If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
    Else
CallServiceReserva:
        If objVenta.ptmModalidad = Venta_Convenio Then
            Set varTemp_CurCabConv = ListaCabConv(Cia, CodLocal, NumProforma)
            Set varTemp_CurDetConv = ListaDetConv(Cia, "" & odynCab("COD_CLIENTE").Value, CodLocal, NumProforma)
            
            agreementFlag = "" & varTemp_CurCabConv("ES_DLV_CONV").Value
            agreementCode = "" & varTemp_CurCabConv("CodConvenio").Value
            'varTemp_CodModalidadVenta = "" & varTemp_CurCabConv("cod_modalidad_venta").Value
            agreementCodCli = "" & varTemp_CurCabConv("CodCliConvenio").Value
            agreementBeneficiaryPct = "" & varTemp_CurCabConv("PctBeneficiario").Value
            
            agreementFields.ReDim 0, -1, 0, 12
            i = 0
            odynDet.MoveFirst
            While Not varTemp_CurDetConv.EOF
                agreementFields.AppendRows
                agreementFields(i, 0) = "" & "" & odynCab("COD_GRUPO_CIA_POSU").Value
                agreementFields(i, 1) = "" & localCodeRef
                agreementFields(i, 2) = "" & NumProforma
                agreementFields(i, 3) = "" & varTemp_CurDetConv("COD_CAMPO").Value
                agreementFields(i, 3) = "" & Trim(Replace(agreementFields(i, 3), "Ã", "_"))
                agreementFields(i, 3) = "" & JsonConverter.toUnicode(procesaCadena(agreementFields(i, 3)))
                agreementFields(i, 4) = "" & varTemp_CurCabConv("CodConvenio").Value
                agreementFields(i, 5) = "" & objVenta.bk_codBeneficiario 'odynCab("COD_CLIENTE").Value 'esto o beneficario guardado en memoria
                agreementFields(i, 6) = ""
                agreementFields(i, 7) = "" & Format(odynCab("FCH_REGISTRA").Value, "yyyy-mm-dd") & "T" & odynCab("HORA").Value
                agreementFields(i, 8) = ""
                'agreementFields(i, 9) = Replace("" & varTemp_CurDetConv("DESCRIPCION_CAMPO").Value, "'", "|")
                agreementFields(i, 9) = "" & varTemp_CurDetConv("DESCRIPCION_CAMPO").Value
                agreementFields(i, 9) = Replace(Trim(agreementFields(i, 9)), "Ã", "_")
                agreementFields(i, 9) = "" & JsonConverter.toUnicode(procesaCadena(agreementFields(i, 9)))
                'agreementFields(i, 10) = Replace("" & varTemp_CurDetConv("NOMBRE_CAMPO").Value, "'", "|")
                agreementFields(i, 10) = "" & varTemp_CurDetConv("NOMBRE_CAMPO").Value
                agreementFields(i, 10) = Replace(Trim(agreementFields(i, 10)), "Ã", "_")
                agreementFields(i, 10) = "" & JsonConverter.toUnicode(procesaCadena(agreementFields(i, 10)))
                agreementFields(i, 11) = "" & varTemp_CurDetConv("FLG_IMPRIME").Value
                agreementFields(i, 12) = "" & Trim(varTemp_CurDetConv("COD_VALOR_IN").Value)
                i = i + 1
                varTemp_CurDetConv.MoveNext
            Wend
            Dim rsCliente As oraDynaset
            Set rsCliente = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA", 0, objVenta.bk_codBeneficiario)
            'dataCli = Replace("" & Trim(rsCliente("DES_NOM_CLIENTE").Value) & " " & Trim(rsCliente("DES_APE_CLIENTE").Value) & " " & Trim(rsCliente("DES_APE2_CLIENTE").Value), "'", "|")
            dataCli = "" & Trim(rsCliente("DES_NOM_CLIENTE").Value) & " " & Trim(rsCliente("DES_APE_CLIENTE").Value) & " " & Trim(rsCliente("DES_APE2_CLIENTE").Value)
            dataCli = Replace(dataCli, "Ã", "_")
            dataCli = "" & JsonConverter.toUnicode(dataCli)
            'nameOf = Replace("" & objUsuario.Codigo & "-" & Trim(objUsuario.Nombre), "'", "|")
            nameOf = "" & objUsuario.Codigo & "-" & Trim(objUsuario.Nombre)
            nameOf = Replace(nameOf, "Ã", "_")
            nameOf = "" & JsonConverter.toUnicode(procesaCadena(nameOf))
        End If
    End If
    'F.ECASTILLO 17.12.2020
cnvNoRuteaAuto:
    user_dni = "" & odynCab("NUM_DOCUMENTO").Value
    user_dni = Me.quitaSaltos(user_dni, "")
    user_dni = "" & JsonConverter.toUnicode(user_dni)
    user_isInkaClub = IIf(Len(Trim(user_dni)) > 0, "Y", "N")
    'user_phone = Replace("" & Trim(odynCliente("DES_VALOR").Value), "'", "|")
    user_phone = "" & Trim(odynCliente("DES_VALOR").Value)
    user_phone = Me.quitaSaltos(user_phone, "")
    user_phone = "" & JsonConverter.toUnicode(procesaCadena(user_phone))
    'user_email = Replace("" & Trim(odynCliente("DES_EMAIL").Value), "'", "|")
    user_email = "" & Trim(odynCliente("DES_EMAIL").Value)
    user_email = Me.quitaSaltos(user_email, "")
    user_email = Replace(user_email, "Ã", "_")
    user_email = "" & JsonConverter.toUnicode(procesaCadena(user_email))
    'user_name = Replace("" & Trim(odynCliente("DES_NOM_CLIENTE").Value), "'", "|")
    user_name = "" & Trim(odynCliente("DES_NOM_CLIENTE").Value)
    user_name = Me.quitaSaltos(user_name, "")
    user_name = Replace(user_name, "Ã", "_")
    user_name = "" & JsonConverter.toUnicode(procesaCadena(user_name))
    'user_lastName = Replace("" & Trim(odynCliente("DES_APE_CLIENTE").Value), "'", "|")
    user_lastName = "" & Trim(odynCliente("DES_APE_CLIENTE").Value)
    user_lastName = Me.quitaSaltos(user_lastName, "")
    user_lastName = Replace(user_lastName, "Ã", "_")
    user_lastName = "" & JsonConverter.toUnicode(procesaCadena(user_lastName))
    user_isAnonymous = "Y"
    
    paymentMethod_id = "" & odynFP("IND_TIPO_PAGO").Value
    paymentMethod_name = "" & odynFP("NOM_TIPO_PAGO").Value
    
    paymentMethod_id = IIf(Len(Trim(paymentMethod_id)) = 0, "1", paymentMethod_id)
    paymentMethod_name = IIf(Len(Trim(paymentMethod_name)) = 0, "Pago Efectivo", paymentMethod_name)
    
    receipt_ruc = "" & odynCab("NUM_RUC_EMPRESA").Value
    receipt_ruc = Me.quitaSaltos(receipt_ruc, "")
    'receipt_companyName = Replace("" & Trim(odynCab("DES_RAZON_SOCIAL").Value), "'", "|")
    receipt_companyName = "" & Trim(odynCab("DES_RAZON_SOCIAL").Value)
    receipt_companyName = Me.quitaSaltos(receipt_companyName, "")
    receipt_companyName = Replace(receipt_companyName, "Ã", "_")
    receipt_companyName = "" & JsonConverter.toUnicode(procesaCadena(receipt_companyName))
    'receipt_companyAddress = Replace("" & Trim(objVenta.dc_street), "'", "|") 'odynCab("DES_DIREC_SOCIAL").Value
    receipt_companyAddress = "" & Trim(objVenta.dc_street)
    receipt_companyAddress = Me.quitaSaltos(receipt_companyAddress, "")
    receipt_companyAddress = Replace(receipt_companyAddress, "Ã", "_")
    receipt_companyAddress = "" & JsonConverter.toUnicode(procesaCadena(receipt_companyAddress))
    receipt_receiptType_name = "" & odynCab("TIPO_DOCUMENTO_DIGITAL").Value
    
    'deliveryAddress_street = Replace("" & Trim(objVenta.dc_street), "'", "|")
    deliveryAddress_street = "" & procesaCadena(Trim(objVenta.dc_street))
    deliveryAddress_street = Me.quitaSaltos(deliveryAddress_street, "")
    deliveryAddress_street = Replace(deliveryAddress_street, "Ã", "_")
    deliveryAddress_street = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_street))
    deliveryAddress_number = "" & objVenta.dc_number
    deliveryAddress_number = Replace(deliveryAddress_number, "Ã", "_")
    deliveryAddress_number = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_number))
    'deliveryAddress_apartment = Replace("" & objVenta.dc_apartment, "'", "|")
    deliveryAddress_apartment = "" & objVenta.dc_apartment
    deliveryAddress_apartment = Replace(deliveryAddress_apartment, "Ã", "_")
    deliveryAddress_apartment = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_apartment))
    'deliveryAddress_country = Replace("" & objVenta.dc_country, "'", "|")
    deliveryAddress_country = "" & objVenta.dc_country
    deliveryAddress_country = Replace(deliveryAddress_country, "Ã", "_")
    deliveryAddress_country = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_country))
    'deliveryAddress_city = Replace("" & objVenta.dc_city, "'", "|")
    deliveryAddress_city = "" & objVenta.dc_city
    deliveryAddress_city = Replace(deliveryAddress_city, "Ã", "_")
    deliveryAddress_city = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_city))
    'deliveryAddress_district = Replace("" & objVenta.dc_district, "'", "|")
    deliveryAddress_district = "" & objVenta.dc_district
    deliveryAddress_district = Replace(deliveryAddress_district, "Ã", "_")
    deliveryAddress_district = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_district))
    'deliveryAddress_district = "" & procesaCadena(deliveryAddress_district)
    'deliveryAddress_latitude = Replace("" & objVenta.dc_latitude, "'", "|")
    deliveryAddress_latitude = "" & objVenta.dc_latitude
    deliveryAddress_latitude = "" & JsonConverter.toUnicode(deliveryAddress_latitude)
    'deliveryAddress_longitude = Replace("" & objVenta.dc_longitude, "'", "|")
    deliveryAddress_longitude = "" & objVenta.dc_longitude
    deliveryAddress_longitude = "" & JsonConverter.toUnicode(deliveryAddress_longitude)
    'deliveryAddress_receiverName = Replace("" & Trim(odynCab("DES_AUX_RECOGE_NOMBRE").Value) & "--" & odynCab("DES_AUX_RECOGE_TLF").Value, "'", "|") 'objVenta.dc_receiverName
    deliveryAddress_receiverName = "" & Trim(odynCab("DES_AUX_RECOGE_NOMBRE").Value) & "--" & odynCab("DES_AUX_RECOGE_TLF").Value
    deliveryAddress_receiverName = Me.quitaSaltos(deliveryAddress_receiverName, "")
    deliveryAddress_receiverName = Replace(deliveryAddress_receiverName, "Ã", "_")
    deliveryAddress_receiverName = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_receiverName))
    deliveryAddress_notes = "" & objVenta.dc_urbanizacion & " " & objVenta.dc_referencia
    deliveryAddress_notes = Replace(deliveryAddress_notes, "Ã", "_")
    deliveryAddress_notes = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_notes))
    deliveryAddress_urbanization = "" & objVenta.dc_urbanizacion
    deliveryAddress_urbanization = Replace(deliveryAddress_urbanization, "Ã", "_")
    deliveryAddress_urbanization = "" & JsonConverter.toUnicode(procesaCadena(deliveryAddress_urbanization))
    argDelivery.ReDim 0, -1, 0, 12
    argDelivery.AppendRows
    argDelivery(0, 0) = "" & deliveryAddress_street
    argDelivery(0, 1) = "" & deliveryAddress_number
    argDelivery(0, 2) = "" & deliveryAddress_apartment
    argDelivery(0, 3) = "" & deliveryAddress_country
    argDelivery(0, 4) = "" & deliveryAddress_city
    argDelivery(0, 5) = "" & deliveryAddress_district
    argDelivery(0, 6) = "" & deliveryAddress_latitude
    argDelivery(0, 7) = "" & deliveryAddress_longitude
    argDelivery(0, 8) = "" & deliveryAddress_receiverName
    argDelivery(0, 9) = "" & deliveryAddress_notes
    argDelivery(0, 10) = "" & deliveryAddress_urbanization
    
    paymentAmountDto_productsTotalCostNoDiscount = "" & Format(CDbl(odynCab("MTO_TOTAL").Value), "0.000")
    paymentAmountDto_grossPrice = "" & Format(CDbl(odynCab("MTO_TOTAL").Value), "0.000")
    paymentAmountDto_discountApplied = "0.000"
    paymentAmountDto_productsTotalCost = "" & Format(CDbl(odynCab("MTO_TOTAL").Value), "0.000")
    paymentAmountDto_deliveryCost = "0.000"
    paymentAmountDto_paidAmountByCoin = "" & IIf(odynFP("COD_MONEDA").Value = 2, odynFP("IMP_SIN_REDONDEO").Value, "") 'dolares
    valeArr.ReDim 0, -1, 0, 2
    i = 0
    odynFPV.MoveFirst
    While Not odynFPV.EOF
        valeArr.AppendRows
        valeArr(i, 0) = "" & odynFPV("COD_FORMA_PAGO_POSU").Value
        valeArr(i, 1) = "" & odynFPV("NUM_VALE").Value
        i = i + 1
        odynFPV.MoveNext
    Wend
    productoArr.ReDim 0, -1, 0, 19
    i = 0
    odynDet.MoveFirst
    Dim fractionatedPrice
    While Not odynDet.EOF
        productoArr.AppendRows
        productoArr(i, 0) = "" & odynDet("COD_POSU_INKA").Value 'CODIGO POSU-INKA
        productoArr(i, 1) = "" & odynDet("COD_POSU").Value 'CODIGO POSU
        productoArr(i, 2) = "" & odynDet("COD_SAP_INKA").Value 'CODIGO SAP
        productoArr(i, 3) = "" & odynDet("DES_PRODUCTO").Value 'NOMBRE
        'productoArr(i, 3) = Replace("" & odynDet("DES_PRODUCTO").Value, "'", "|") 'NOMBRE
        'productoArr(i, 3) = Replace("" & productoArr(i, 3), """", "|") 'NOMBRE
        productoArr(i, 3) = Me.quitaSaltos(productoArr(i, 3), "")
        productoArr(i, 3) = Replace(productoArr(i, 3), "Ã", "_")
        productoArr(i, 3) = "" & JsonConverter.toUnicode(procesaCadena(productoArr(i, 3)))
        'productoArr(i, 4) = "" 'DESCRIPCIÓN
        'productoArr(i, 5) = Replace("" & odynDet("DES_MARCA").Value, "'", "|") 'MARCA
        productoArr(i, 5) = "" & odynDet("DES_MARCA").Value 'MARCA
        productoArr(i, 5) = Me.quitaSaltos(productoArr(i, 5), "")
        productoArr(i, 5) = Replace(productoArr(i, 5), "Ã", "_")
        productoArr(i, 5) = "" & JsonConverter.toUnicode(procesaCadena(productoArr(i, 5)))
        productoArr(i, 6) = "" & odynDet("QUANTITY").Value 'CANTIDAD
        productoArr(i, 7) = "0" 'CANTIDAD UNIDAD
        productoArr(i, 8) = "" & IIf(odynDet("FLG_FRACCIONAMIENTO").Value = 1, "Y", "N") 'FRACCIONADO
        fractionatedPrice = "" & odynDet("FRACTIONATEDPRICE").Value
        fractionatedPrice = IIf(Len(Trim(fractionatedPrice)) = 0, "0", fractionatedPrice)
        If fractionatedPrice = "0" Then
            productoArr(i, 9) = "" & Format(CDbl(fractionatedPrice), "0.000") 'PRECIO UNITARIO
        Else
            productoArr(i, 9) = "" & Format(CDbl(fractionatedPrice) - 0.0005, "0.000") 'PRECIO UNITARIO
        End If
        'productoArr(i, 9) = Format(CDbl(odynDet("FRACTIONATEDPRICE").Value) - 0.0005, "0.000") 'PRECIO UNITARIO
        productoArr(i, 10) = "" & Format(CDbl(odynDet("MTO_SUBTOTAL").Value), "0.000") 'PRECIO TOTAL
        productoArr(i, 11) = "" & odynDet("UNID_VTA").Value 'PRESENTACIÓN
        'productoArr(i, 12) = objVenta.Producto(i, 4) 'TIPO FAMILIA
        'productoArr(i, 13) = objVenta.Producto(i, 4) 'CODIGO EAN}
        If fractionatedPrice = "0" Then
            productoArr(i, 14) = "" & Format(CDbl(fractionatedPrice), "0.000") 'PRECIO UNITARIO
        Else
            productoArr(i, 14) = "" & Format(CDbl(fractionatedPrice) - 0.0005, "0.000") 'PRECIO UNITARIO
        End If
        'productoArr(i, 14) = Format(CDbl(odynDet("FRACTIONATEDPRICE").Value) - 0.0005, "0.000") 'PRECIO FRACCIONADO
        productoArr(i, 15) = "" & IIf(odynDet("FLG_FRACCIONAMIENTO").Value = 1, "3", "1") 'TIPO PRESENTACIÓN
        productoArr(i, 16) = "" & odynDet("VAL_FRAC_LOCAL").Value 'VALOR FRACCION
        productoArr(i, 17) = "" & odynDet("VALOR_UMV").Value 'VALOR UMV
        productoArr(i, 18) = "" & odynDet("QUANTITYTOTAL").Value 'CANTIDAD TOTAL
        i = i + 1
        odynDet.MoveNext
    Wend
    'observation = Replace("" & Trim(odynCab("OBS_NOTA_LOCAL").Value), "'", "|")
    observation = "" & Trim(odynCab("OBS_NOTA_LOCAL").Value)
    observation = quitaSaltos(observation, " ")
    observation = Replace(observation, "Ã", "_")
    observation = "" & JsonConverter.toUnicode(procesaCadena(observation))
    'observationMotorized = Replace("" & Trim(odynCab("OBS_NOTA_MOTORIZADO").Value), "'", "|")
    observationMotorized = "" & Trim(odynCab("OBS_NOTA_MOTORIZADO").Value)
    observationMotorized = quitaSaltos(observationMotorized, " ")
    observationMotorized = Replace(observationMotorized, "Ã", "_")
    observationMotorized = "" & JsonConverter.toUnicode(procesaCadena(observationMotorized))
    
    
    If gstrIndRAv3 = "2" Then
    argAdicionales.ReDim 0, -1, 0, 12
    argAdicionales.AppendRows
    argAdicionales(0, 0) = "" & IIf(odynCab("IND_RESERVE_CAPACIDAD").Value = "1", "false", "true")
    argAdicionales(0, 1) = "" & odynCab("START_HOUR").Value
    argAdicionales(0, 1) = IIf(Len(Trim(argAdicionales(0, 1))) = 0, "null", argAdicionales(0, 1))
    argAdicionales(0, 2) = "" & odynCab("END_HOUR").Value
    argAdicionales(0, 2) = IIf(Len(Trim(argAdicionales(0, 2))) = 0, "null", argAdicionales(0, 2))
    End If
    
    If gstrIndDCSAP = "1" Then
        If objVenta.isDCSAP = "1" Then
            Set resp = objWS.asignaDigitalSAP(proformaNumber, localCode, ciaCode, localCodeRef, dateCreated, dateConfirmada, processStartConfirm, Source, companyCode, creditCardProviderId, currencyId, discountApplied, _
                                    deliveryType, totalTax, amount, flgCappa, user_dni, user_isInkaClub, user_phone, user_email, user_name, user_lastName, user_isAnonymous, paymentMethod_id, _
                                    paymentMethod_name, receipt_ruc, receipt_companyName, receipt_companyAddress, receipt_receiptType_name, argDelivery, _
                                    paymentAmountDto_productsTotalCostNoDiscount, paymentAmountDto_grossPrice, paymentAmountDto_discountApplied, paymentAmountDto_productsTotalCost, _
                                    paymentAmountDto_deliveryCost, paymentAmountDto_paidAmountByCoin, productoArr, valeArr, _
                                    agreementFlag, agreementCode, agreementCodCli, agreementBeneficiaryPct, agreementFields, dataCli, nameOf, observation, observationMotorized, flagMultiBrand, deliveryTime, dateConfirmadaEnd, argAdicionales)
        Else
            GoTo asignaRegular
        End If
    Else
asignaRegular:
        Set resp = objWS.asignaDigital(proformaNumber, localCode, ciaCode, localCodeRef, dateCreated, dateConfirmada, processStartConfirm, Source, companyCode, creditCardProviderId, currencyId, discountApplied, _
                                    deliveryType, totalTax, amount, flgCappa, user_dni, user_isInkaClub, user_phone, user_email, user_name, user_lastName, user_isAnonymous, paymentMethod_id, _
                                    paymentMethod_name, receipt_ruc, receipt_companyName, receipt_companyAddress, receipt_receiptType_name, argDelivery, _
                                    paymentAmountDto_productsTotalCostNoDiscount, paymentAmountDto_grossPrice, paymentAmountDto_discountApplied, paymentAmountDto_productsTotalCost, _
                                    paymentAmountDto_deliveryCost, paymentAmountDto_paidAmountByCoin, productoArr, valeArr, _
                                    agreementFlag, agreementCode, agreementCodCli, agreementBeneficiaryPct, agreementFields, dataCli, nameOf, observation, observationMotorized, flagMultiBrand, deliveryTime, dateConfirmadaEnd, argAdicionales)
    End If
    If Not resp Is Nothing Then
        If IsObject(resp("data")) = False Then
            Exit Function
        End If
        If resp("data")("description") = "OK" Or resp("data")("description") = "successful process" Then
            asignaDigital = "1"
            objVenta.dc_orderDigital = resp("data")("code")
        Else
            'msg error
            objVenta.dc_orderDigital = ""
        End If
    End If
End Function
Public Function cancelaDigital(ByVal Cia As String, ByVal CodLocal As String, ByVal NumProforma As String) As String
    'ByVal order As String, Optional orderCode As String, Optional orderObs As String
    Dim resp As Dictionary
    Dim odynCab As oraDynaset
    Dim order As String
    Dim flgReservo As String
    cancelaDigital = "0"
    Set odynCab = ListaCabecera(Cia, CodLocal, NumProforma)
    'flgReservo = "" & odynCab("FLG_RESERVA_CAPACIDAD").Value
    'If flgReservo <> "1" Then GoTo ok
    order = "" & odynCab("ORDER_DIGITAL").Value
    If Len(Trim(order)) = 0 Then GoTo ok
    Set resp = objWS.cancelaDigital(order, "DU")
    cancelaDigital = "Sistema inicio anulación, pero no se completo el proceso por inconveniente con el servicio."
    If Not resp Is Nothing Then
        If IsObject(resp("data")) = False Then
            Exit Function
        End If
        If resp("data")("orderStatus")("name") = "CANCELLED_ORDER" Then
            cancelaDigital = "1"
        Else
            cancelaDigital = "code: " & resp("data")("orderStatus")("code") & vbNewLine & _
                             "detail: " & resp("data")("orderStatus")("detail")
        End If
    End If
ok:
    'cancelaDigital = "0"
End Function
Public Function asignaOrderDigitalProf(ByVal A_CIA As String, _
                                     ByVal A_COD_LOCAL As String, _
                                     ByVal A_NUM_PROFORMA As String, _
                                     ByVal A_ORDER_DIGITAL As String, _
                                     ByVal A_FLG_COMMIT As String) As String
Dim varValores As Variant
Dim varIO As Variant
On Error GoTo handle
    varValores = Array(A_CIA, A_COD_LOCAL, A_NUM_PROFORMA, A_ORDER_DIGITAL, A_FLG_COMMIT)
    varIO = Array(entrada, entrada, entrada, entrada, entrada)
    asignaOrderDigitalProf = gclsOracle.SP("BTLPROD.PKG_PROFORMA_V3.SP_ASIGNA_ORDER_DIGITAL", varValores, varIO)
    Exit Function
handle:
    Err.Raise Err.Number, "clsProforma.asignaOrderDigitalProf", Err.Description
End Function
'F.ECASTILLO 27.10.2020

Function quitaSaltos(ByVal str As String, ByVal str2 As String) As String
    Dim newStr As String
    newStr = Replace(str, Chr(13), str2)
    newStr = Replace(newStr, vbNewLine, str2, , , vbTextCompare)
    newStr = Replace(newStr, vbLf, str2)
    quitaSaltos = newStr
End Function

Function procesaCadena(ByVal str As String) As String
    Dim newStr As Variant
    Dim response As String
    Dim arrToFind() As String
    Dim charsToRemove As String
    Dim i As Integer
    If gstrIndDCSAP = "1" Then
        If objVenta.isDCSAP = "1" Then
            charsToRemove = gstrSpecialChrsToWS1
        Else
            GoTo regular
        End If
    Else
regular:
        charsToRemove = gstrSpecialChrsToWS0
    End If
    newStr = str
    arrToFind = Split(charsToRemove, ",")
    For i = 0 To UBound(arrToFind)
        If InStr(newStr, arrToFind(i)) > 0 Then
            newStr = Replace(newStr, arrToFind(i), " ")
        End If
    Next
    response = Trim(newStr)
    procesaCadena = response
End Function

