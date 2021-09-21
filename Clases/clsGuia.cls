VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaGuia(ByVal vstrNum_Guia As String, _
                          ByVal vstrFch_Ini As String, _
                          ByVal vstrFch_Fin As String, _
                          ByVal vstrEst_Guia As String, _
                          ByVal vstrOrigen As String, _
                          ByVal vstrDestino As String, _
                          ByVal vstrCod_Motivo_Guia As String, _
                          ByVal vstrCod_Producto As String) As oraDynaset

    On Error GoTo CtrlErr
    Set ListaGuia = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_AL_DETALLE", 0, _
                            vstrNum_Guia, _
                            vstrFch_Ini, _
                            vstrFch_Fin, _
                            vstrEst_Guia, _
                            vstrOrigen, _
                            vstrDestino, _
                            vstrCod_Motivo_Guia, _
                            vstrCod_Producto)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia", Err.Description
End Function

Public Function ListaCab(Optional ByVal vstrNumguia As String, _
                        Optional ByVal vstrCodOrigen As String, _
                        Optional ByVal vstrCodDestino As String, _
                        Optional ByVal vstrEstGuia As String, _
                        Optional ByVal vstrFchEmi_Ini As String, _
                        Optional ByVal vstrFchEmi_Fin As String, _
                        Optional ByVal vstrFchRec_Ini As String, _
                        Optional ByVal vstrFchRec_Fin As String, _
                        Optional ByVal vstrFlgGuia_Prov As String = "", _
                        Optional ByVal vstrCodProducto As String = "") As oraDynaset

    On Error GoTo CtrlErr
    Set ListaCab = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_LISTA_CAB", 0, _
                                        vstrNumguia, _
                                        vstrCodOrigen, _
                                        vstrCodDestino, _
                                        vstrEstGuia, _
                                        vstrFchEmi_Ini, _
                                        vstrFchEmi_Fin, _
                                        vstrFchRec_Ini, _
                                        vstrFchRec_Fin, _
                                        vstrFlgGuia_Prov, _
                                        vstrCodProducto, _
                                        objUsuario.CodigoEmpresa, _
                                        objUsuario.CodigoLocal)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia::Lista_Cab", Err.Description
End Function

Public Function ListaDet(ByVal vstrNumguia As String) As oraDynaset

    On Error GoTo CtrlErr
    Set ListaDet = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_LISTA_DET", 0, _
                                        vstrNumguia, _
                                        objUsuario.CodigoEmpresa, _
                                        objUsuario.CodigoLocal)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia::Lista_Det", Err.Description
End Function

Public Function ListaDetLote(ByVal vstrNumguia As String) As oraDynaset

    On Error GoTo CtrlErr
    Set ListaDetLote = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_LISTA_DET_LOTE", 0, _
                                            vstrNumguia, _
                                            objUsuario.CodigoEmpresa, _
                                            objUsuario.CodigoLocal)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia::Lista_Det", Err.Description
End Function

Public Function Anula(ByVal vstrCia As String, _
                      ByVal vstrCod_Local As String, _
                      ByVal vstrNum_Guia As String, _
                      ByVal vstrCod_Usuario As String) As String

    On Error GoTo CtrlErr
    
    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(vstrCia, vstrCod_Local, vstrNum_Guia, vstrCod_Usuario)
    varIO = Array(entrada, entrada, entrada, entrada)
    
    Anula = gclsOracle.SP("NUEVO.PKG_GUIA.SP_ANULA_GUIA_LOCAL", varValores, varIO)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia", Err.Description
End Function

Public Function Recepciona(ByVal vstrCia As String, _
                      ByVal vstrCod_Local As String, _
                      ByVal vstrNum_Guia As String, _
                      ByVal vstrCod_Usuario As String) As String

    On Error GoTo CtrlErr
    
    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(vstrCia, vstrCod_Local, vstrNum_Guia, vstrCod_Usuario)
    varIO = Array(entrada, entrada, entrada, entrada)
    
    Recepciona = gclsOracle.SP("NUEVO.PKG_GUIA.SP_RECEPCIONA_GUIA_LOCAL", varValores, varIO)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia", Err.Description
End Function

'========================================================
Public Function RecepcionaNew(ByVal vstrCia As String, _
                              ByVal vstrCod_Local As String, _
                              ByVal vstrNum_Guia As String, _
                              ByVal vstrCod_Usuario As String, _
                              ByVal varrDet_Recepcion As XArrayDB) As String

    Dim NumItem As OracleInProcServer.OraParamArray
    Dim NumLote As OracleInProcServer.OraParamArray
    Dim CodProd As OracleInProcServer.OraParamArray
    Dim CtdUnid As OracleInProcServer.OraParamArray
    Dim ctdFrac As OracleInProcServer.OraParamArray
    Dim NroLote As OracleInProcServer.OraParamArray
    Dim FchVmto As OracleInProcServer.OraParamArray
    Dim byteNumElementos As Byte
    Dim oradb As OraDatabase
    Dim i As Integer
    
    On Error GoTo CtrlErr
    
    byteNumElementos = IIf(varrDet_Recepcion.UpperBound(1) < 0, 0, varrDet_Recepcion.UpperBound(1)) + 1
    
    Set oradb = gclsOracle.ODataBase
    
    For i = oradb.Parameters.Count - 1 To 0 Step -1
        oradb.Parameters.Remove i
    Next
    
    oradb.Parameters.Add "A_COD_CIA", vstrCia, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_COD_LOCAL", vstrCod_Local, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_NUM_GUIA", vstrNum_Guia, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_COD_USUARIO", vstrCod_Usuario, ORAPARM_INPUT, ORATYPE_VARCHAR2

    oradb.Parameters.AddTable "A_ARR_NUM_ITEM", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementos, 200
    oradb.Parameters.AddTable "A_ARR_NUM_ITEM_LOTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementos, 200
    oradb.Parameters.AddTable "A_ARR_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementos, 200
    oradb.Parameters.AddTable "A_ARR_NUM_LOTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementos, 200
    oradb.Parameters.AddTable "A_ARR_FCH_VENCIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementos, 200
    oradb.Parameters.AddTable "A_ARR_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementos, 200
    oradb.Parameters.AddTable "A_ARR_CTD_PRODUCTO_FRAC", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementos, 200

    Set NumItem = oradb.Parameters("A_ARR_NUM_ITEM")
    Set NumLote = oradb.Parameters("A_ARR_NUM_ITEM_LOTE")
    Set CodProd = oradb.Parameters("A_ARR_COD_PRODUCTO")
    Set NroLote = oradb.Parameters("A_ARR_NUM_LOTE")
    Set FchVmto = oradb.Parameters("A_ARR_FCH_VENCIMIENTO")
    Set CtdUnid = oradb.Parameters("A_ARR_CTD_PRODUCTO")
    Set ctdFrac = oradb.Parameters("A_ARR_CTD_PRODUCTO_FRAC")
    
    For i = 0 To varrDet_Recepcion.UpperBound(1)
        NumItem(i) = varrDet_Recepcion(i, 2)    'NUM_ITEM
        NumLote(i) = varrDet_Recepcion(i, 3)    'NUM_ITEM_LOTE
        CodProd(i) = varrDet_Recepcion(i, 4)    'COD_PRODUCTO
        NroLote(i) = varrDet_Recepcion(i, 7)    'NUM_LOTE
        FchVmto(i) = varrDet_Recepcion(i, 8)    'FCH_VENCIMIENTO
        CtdUnid(i) = varrDet_Recepcion(i, 13)   'CTDU
        ctdFrac(i) = varrDet_Recepcion(i, 14)   'CTDF
    Next
    
    oradb.ExecuteSQL "BEGIN nuevo.pkg_guia.sp_recepciona_guia_local_new(" & _
                                          ":a_cod_cia," & _
                                          ":a_cod_local," & _
                                          ":a_num_guia," & _
                                          ":a_cod_usuario," & _
                                          ":a_arr_num_item," & _
                                          ":a_arr_num_item_lote," & _
                                          ":a_arr_cod_producto," & _
                                          ":a_arr_num_lote," & _
                                          ":a_arr_fch_vencimiento," & _
                                          ":a_arr_ctd_producto," & _
                                          ":a_arr_ctd_producto_frac);  END;"
    RecepcionaNew = True
    
    Exit Function
CtrlErr:
    RecepcionaNew = False
    Err.Raise Err.Number, "clsGuia", Err.Description
End Function
'========================================================

Public Function ImprimeDetalle(ByVal vstrNumguia As String) As oraDynaset

    On Error GoTo CntrlErr
        
        Set ImprimeDetalle = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_IMP_DET", 0, vstrNumguia, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
        
        Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsGuia.ImprimeDetalle", Err.Description
End Function


Public Function ImprimeCabecera(ByVal vstrNumguia As String) As oraDynaset

    On Error GoTo CntrlErr
        
        Set ImprimeCabecera = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_IMP_CAB", 0, vstrNumguia, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
        
        Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsGuia.ImprimeCabecera", Err.Description
End Function

Public Function Imp_Cab_Dev(ByVal vstrRuc_Proveedor As String, _
                            ByVal vstrTip_Documento As String, _
                            ByVal vstrNum_Documento As String) As oraDynaset

    On Error GoTo CtrlErr
    Set Imp_Cab_Dev = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_IMP_CAB_DEV", 0, vstrRuc_Proveedor, _
                                                                           vstrTip_Documento, _
                                                                           vstrNum_Documento)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia", Err.Description
End Function

Public Function Imp_Det_Dev(ByVal vstrRuc_Proveedor As String, _
                            ByVal vstrTip_Documento As String, _
                            ByVal vstrNum_Documento As String, _
                            ByVal vstrNum_Guia As String) As oraDynaset

    On Error GoTo CtrlErr
    Set Imp_Det_Dev = gclsOracle.FN_Cursor("NUEVO.PKG_GUIA.FN_IMP_DET_DEV", 0, vstrRuc_Proveedor, _
                                                                           vstrTip_Documento, _
                                                                           vstrNum_Documento, _
                                                                           vstrNum_Guia)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsGuia", Err.Description
End Function

Public Sub spImprime_Guia_Dev(ByVal vstrRuc_Proveedor As String, _
                              ByVal vstrTip_Documento As String, _
                              ByVal vstrNum_Documento As String)

Dim rstCabecera As oraDynaset
Dim rstDetalle As oraDynaset
Dim dblCtdCodigo As Double
Dim dblCtdUni As Double
Dim p As Printer


    For Each p In Printers
       If UCase(p.Port) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
    

    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
        Dim objImpresion As New clsImpresion
     
        If Not objImpresion.cargaParametros("002") Then
            Exit Sub
        End If
    Else
        Printer.Width = 1440 * 8.5
        Printer.Height = 1440 * 8
    End If
    
 '-------------------------------------------------------

    Set rstCabecera = Imp_Cab_Dev(vstrRuc_Proveedor, vstrTip_Documento, vstrNum_Documento)

    If rstCabecera.RecordCount = 0 Then
        MsgBox "Documento no Tiene Devoluciones", vbInformation, "Impresión"
        Exit Sub
    End If

    If rstCabecera.RecordCount > 1 Then
        MsgBox "Error al Leer Base de Datos", vbExclamation, "Impresión"
        Exit Sub
    End If

    Set rstDetalle = Imp_Det_Dev(rstCabecera("RUC_DESTINO").Value, rstCabecera("TIP_DOCUMENTO").Value, rstCabecera("NUM_DOCUMENTO").Value, rstCabecera("NUM_GUIA").Value)
    If rstDetalle.EOF Then
        MsgBox "Documento no Tiene Deatalle", vbCritical, "Impresión"
        Exit Sub
    End If
Copia:
    rstCabecera.MoveFirst
    rstDetalle.MoveFirst

    On Error GoTo ErrorImpresora

    Printer.FontName = "Draft 12cpi"
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Space(15) & Format(rstCabecera("FCH_EMISION").Value, "DD") & Space(4) & left(Format(rstCabecera("FCH_EMISION").Value, "MMMM") & Space(10), 10) & Space(6) & Format(rstCabecera("FCH_EMISION").Value, "YYYY")
    Printer.FontName = "Draft 17cpi"
    Printer.Print Space(20) & gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", rstCabecera("COD_ORIGEN").Value)
    Printer.FontName = "Draft 12cpi"
    Printer.Print Space(20) & "PROVEEDOR"
    Printer.Print Space(70) & Format(rstCabecera("NUM_GUIA").Value, "000-0000000")
    Printer.Print Space(20) & rstCabecera("DES_PROVEEDOR").Value
    Printer.FontName = "Draft 17cpi"
    Printer.Print Space(20) & rstCabecera("DIR_PROVEEDOR").Value
    Printer.FontName = "Draft 12cpi"
    Printer.Print Space(20) & rstCabecera("RUC_DESTINO").Value
    Printer.Print
    Printer.FontName = "Draft 17cpi"
    Printer.Print Space(3) & String(130, "-")
    Printer.Print Space(5) & " #" & Space(2) & _
                  left("CODIGO" & Space(5), 5) & Space(2) & _
                  left("DESCRIPCIÓN" & Space(60), 60) & Space(2) & _
                  right(Space(8) & "UND.", 8) & Space(2) & _
                  left("MOTIVO" & Space(50), 50)
    Printer.Print Space(3) & String(130, "-")
    While Not rstDetalle.EOF
        Printer.Print Space(5) & Format(rstDetalle("NUM_ITEM").Value, "00") & Space(2) & _
                      left(rstDetalle("COD_PRODUCTO").Value & Space(5), 5) & Space(2) & _
                      left(rstDetalle("DES_PRODUCTO").Value & Space(60), 60) & Space(2) & _
                      right(Space(8) & rstDetalle("CTD_PRODUCTO").Value, 8) & Space(2) & _
                      left(Trim(rstDetalle("DEV").Value) & Space(40), 40)
                      dblCtdCodigo = dblCtdCodigo + Val(rstDetalle("COD_PRODUCTO").Value)
                      dblCtdUni = dblCtdUni + Val(rstDetalle("CTD_PRODUCTO").Value)

        rstDetalle.MoveNext
    Wend
    Printer.Print Space(3) & String(130, "-")
    Printer.Print
    Printer.Print Space(20) & "Suma de Códigos =  " & right(Space(8) & CStr(dblCtdCodigo), 8)
    Printer.Print Space(20) & "Suma de Cant. U =  " & right(Space(8) & CStr(dblCtdUni), 8)
    Printer.Print
    Printer.Print Space(30) & "Usuario : " & rstCabecera("USU_EMISOR").Value & " - " & rstCabecera("USUARIO").Value
    Printer.Print
    Printer.Print " ORIGINADA POR DOC. " & rstCabecera("TIP_DOCUMENTO").Value & " - " & rstCabecera("NUM_DOCUMENTO").Value & "  RELACIONADA CON LA OC: " & rstCabecera("NUM_ORDEN_COMPRA").Value
    Printer.EndDoc
   
    On Error GoTo 0
    GoTo final
 
ErrorImpresora:
 
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
    End If
final:
End Sub


Public Function ListaDestinos(ByVal vstrCia As String, _
                              ByVal vstrCodLocal As String, _
                              ByVal vsrtCodTipoDev As String, _
                              ByVal vstrCodMotivoDev As String) As oraDynaset
    Set ListaDestinos = gclsOracle.FN_Cursor("BTLPROD.PKG_ORDEN_DEVOLUCION.FN_LISTA_DESTINO", 0, vstrCia, vstrCodLocal, vsrtCodTipoDev, vstrCodMotivoDev)
End Function
