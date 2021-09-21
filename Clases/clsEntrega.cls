VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista(ByVal CodigoLocal As String, ByVal FechaInicio As String, ByVal FechaFin As String, ByVal numGuia As String, ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA", 0, CodigoLocal, FechaInicio, FechaFin, numGuia, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.Lista", Err.Description
End Function
Public Function Lista_Sobran(ByVal CodigoLocal As String, ByVal FechaInicio As String, ByVal FechaFin As String, ByVal numGuia As String, ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set Lista_Sobran = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA", 0, CodigoLocal, FechaInicio, FechaFin, numGuia, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.Lista_Sobran", Err.Description
End Function

Public Function ListaDetalle(ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaDetalle = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA_DETALLE", 0, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaDetalle", Err.Description
End Function

Public Function ListaCabGuia(ByVal numGuia As String, ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaCabGuia = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA_CABGUIA", 0, numGuia, CodigoLocal)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaCabGuia", Err.Description
End Function
Public Function ListaCabGuia_Sob(ByVal numGuia As String, ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaCabGuia_Sob = gclsOracle.FN_Cursor("BTLPROD.PKG_PROCE_READ_SOB_FAL_DEV.fn_lista_cabguia_sob", 0, numGuia, CodigoLocal)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaCabGuia_Sob", Err.Description
End Function

Public Function ListaDetGuia(ByVal numGuia As String, busca As String, ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaDetGuia = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA_DETGUIA", 0, numGuia, busca, CodigoLocal)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaDetGuia", Err.Description
End Function
Public Function ListaDetGuiaSob(ByVal numGuia As String, busca As String, ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaDetGuiaSob = gclsOracle.FN_Cursor("BTLPROD.PKG_PROCE_READ_SOB_FAL_DEV.fn_lista_detguia_sob", 0, numGuia, busca, CodigoLocal)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaDetGuia", Err.Description
End Function

Public Function ListaDetGuia_Sob(ByVal numGuia As String, busca As String, ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaDetGuia_Sob = gclsOracle.FN_Cursor("BTLPROD.PKG_PROCE_READ_SOB_FAL_DEV.fn_lista_detguia_sob", 0, numGuia, busca, CodigoLocal)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaDetGuia", Err.Description
End Function


Public Function ListaPendiente(ByVal CodigoLocal As String, ByVal idEntrega As String, ByVal numGuia As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaPendiente = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA_PENDIENTES", 0, CodigoLocal, idEntrega, numGuia)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaPendiente", Err.Description
End Function
Public Function ListaPend_Sob(ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaPend_Sob = gclsOracle.FN_Cursor("BTLPROD.PKG_PROCE_READ_SOB_FAL_DEV.fn_lista_sob_pendien", 0, CodigoLocal)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaPend_Sob", Err.Description
End Function


Public Function ListaAsociados(ByVal CodigoLocal As String, ByVal idEntrega As String, ByVal numGuia As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaAsociados = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.fn_lista_asociados", 0, CodigoLocal, idEntrega, numGuia)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaAsociados", Err.Description
End Function

Public Function ListaProducto(ByVal idEntrega As String, ByVal Busqueda As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaProducto = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_BUSCA_PRODUCTO", 0, idEntrega, Busqueda)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaProducto", Err.Description
End Function



Public Function GrabaEntrega(ByRef Entrega As String, _
                             ByVal numGuia As String, _
                             ByVal CodigoUsuario As String, _
                             ByVal CodigoLocal As String, _
                             ByVal CodigoEmpresa As String, _
                             ByVal Chofer As String, _
                             ByVal placa As String, _
                             ByVal Bultos As String, _
                             ByVal Precintos As String, _
                             ByVal Glosa As String _
                             ) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(Entrega, _
                         numGuia, _
                         CodigoUsuario, CodigoLocal, _
                        CodigoEmpresa, Chofer, placa, Bultos, Precintos, Glosa)
 
    gvarIO = Array(Entrada_Salida, _
                    entrada, _
                    entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
                
                    
    GrabaEntrega = gclsOracle.SP("BTLPROD.PKG_ENTREGA.SP_GRABA_ENTREGA", _
                              gvarValores, _
                              gvarIO)
    Entrega = gvarValores(0)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsLocal.Graba", Err.Description
End Function

Public Function GrabaTransportista(ByRef Entrega As String, _
                                   ByVal CodigoUsuario As String, _
                                   ByVal CodigoLocal As String, _
                                   ByVal CodigoEmpresa As String, _
                                   ByVal Chofer As String, _
                                   ByVal placa As String, _
                                   ByVal Bultos As String, _
                                   ByVal Precintos As String, _
                                   ByVal Glosa As String _
                             ) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(Entrega, _
                        CodigoUsuario, CodigoLocal, _
                        CodigoEmpresa, Chofer, placa, Bultos, Precintos, Glosa)
 
    gvarIO = Array(Entrada_Salida, _
                    entrada, _
                    entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
                
                    
    GrabaTransportista = gclsOracle.SP("BTLPROD.PKG_ENTREGA.sp_graba_transportista", _
                              gvarValores, _
                              gvarIO)
    Entrega = gvarValores(0)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsLocal.GrabaTransportista", Err.Description
End Function

Public Function GrabaGuias(ByVal Entrega As String, _
                             ByVal numGuia As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(Entrega, numGuia)
 
    gvarIO = Array(entrada, entrada)
                
                    
    GrabaGuias = gclsOracle.SP("BTLPROD.PKG_ENTREGA.sp_graba_guias", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsLocal.GrabaGuias", Err.Description
End Function

Public Function GrabaGuias_Sob(Optional ByVal Entrega As String, _
                               Optional ByVal numGuia As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(Entrega, numGuia)
 
    gvarIO = Array(entrada, entrada)
                
                    
    GrabaGuias_Sob = gclsOracle.SP("btlprod.pkg_proce_read_sob_fal_dev.sp_graba_guias_sob", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsLocal.GrabaGuias_Sob", Err.Description
End Function

Public Function ListaTransportista(ByVal IdTransportista As String, _
                          ByVal FlagEstado As String, _
                          ByVal Texto As String, _
                          ByVal busca As String) As OracleInProcServer.oraDynaset
On Error GoTo Control

    Set ListaTransportista = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA_TRANSPORTISTA", 0, IdTransportista, FlagEstado, Texto, busca)
    Exit Function
Control:
    Err.Raise Err.Number, "clsAprobacion.ListaTransportista", Err.Description
End Function


Public Function ListaProducto2(ByVal idEntrega As String, ByVal Busqueda As String, ByVal esReconteo As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaProducto2 = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_BUSCA_PRODUCTO_2", 0, idEntrega, Busqueda, esReconteo)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaProducto2", Err.Description
End Function

Public Function ListaConteo(ByVal idEntrega As String, ByVal Reconteo As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaConteo = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_INGRESA_PRODUCTO", 0, idEntrega, Reconteo)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaConteo", Err.Description

End Function

Public Function ListaConteoOrden(ByVal idEntrega As String, ByVal Reconteo As String, ByVal codProducto As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaConteoOrden = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_ORDENA_PRODUCTO", 0, idEntrega, Reconteo, codProducto)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaConteo", Err.Description

End Function

Public Function AgregaProducto(ByVal idEntrega As String, ByVal Codigo As String, ByVal cantidad As String, ByVal Reconteo As String, Optional tp_reconteo As String = "1") As oraDynaset
    On Error GoTo CtrlError
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, Codigo, cantidad, Reconteo, tp_reconteo)
 
    gvarIO = Array(entrada, entrada, entrada, entrada, entrada)
                
                    
     gclsOracle.SP "BTLPROD.PKG_ENTREGA.SP_INGRESA_PRODUCTO", _
                              gvarValores, _
                              gvarIO
    
    Set AgregaProducto = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_INGRESA_PRODUCTO", 0, idEntrega, Reconteo)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.AgregaProducto", Err.Description
End Function


Public Function Cierra(ByVal idEntrega As String, ByVal usuario As String, ByVal FlgReconteo As String, CodMaquina As String, ByRef numGuiaDev As String) As String
    On Error GoTo CtrlError
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, usuario, FlgReconteo, CodMaquina, numGuiaDev)
 
    gvarIO = Array(entrada, entrada, entrada, entrada, Salida)
                
                    
     Cierra = gclsOracle.SP("BTLPROD.PKG_ENTREGA.SP_CIERRA_ENTREGA", _
                              gvarValores, _
                              gvarIO)

numGuiaDev = gvarValores(4)

    If Cierra <> "" Then GoTo CtrlError
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.Cierra", Cierra
End Function

Public Function ListaDiferencias(ByVal idEntrega As String, Tipo As String) As oraDynaset
    On Error GoTo CtrlError
    
    'Set ListaDiferencias = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.FN_LISTA_DIFERENCIA", 0, identrega, Tipo)
    'Set ListaDiferencias = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA_ADIC_001.FN_LISTA_DIFERENCIA", 0, idEntrega, Tipo)
    Set ListaDiferencias = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA_ADIC_001.FN_LISTA_DIFERENCIA2", 0, idEntrega, Tipo)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaDiferencias", Err.Description
End Function

Public Function GrabaFaltantesYSobrantes(ByVal idEntrega As String, Tipo As String) As String
    On Error GoTo CtrlError
    Dim gvarValores As Variant
    Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, Tipo)
 
    gvarIO = Array(entrada, entrada)
    
    GrabaFaltantesYSobrantes = gclsOracle.SP("BTLPROD.PKG_ENTREGA_ADIC_001.SP_SAVE_SOBRAN_FALT_ENTREG", gvarValores, gvarIO)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.GrabaFaltantesYSobrantes", Err.Description
End Function

Public Function EditaLoteEnSobrantes(ByVal idEntrega As String, ByVal NumLote As String, ByVal codProducto As String, ByVal numGuia As String, numEntrega As String) As String
    On Error GoTo CtrlError
    Dim gvarValores As Variant
    Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, NumLote, codProducto, numGuia, numEntrega)
 
    gvarIO = Array(entrada, entrada, entrada, entrada, entrada)
    
    EditaLoteEnSobrantes = gclsOracle.SP("BTLPROD.PKG_ENTREGA_ADIC_001.SP_UPD_LOTE_SOBRAN", gvarValores, gvarIO)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.EditaLoteEnSobrantes", Err.Description
End Function

Public Function ListaConteoAux(ByVal idEntrega As String, momento As Integer) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaConteoAux = gclsOracle.FN_Cursor("BTLPROD.pkg_aux_conteo.fn_lista_conteo", 0, idEntrega, momento)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaConteoAux", Err.Description
End Function

Public Function momento(ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set momento = gclsOracle.FN_Cursor("BTLPROD.pkg_aux_conteo.fn_Momento", 0, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.Momento", Err.Description
End Function

Public Function ConsolidaConteoAux(ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ConsolidaConteoAux = gclsOracle.FN_Cursor("BTLPROD.pkg_aux_conteo.fn_Consolida_conteo", 0, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ConsolidaConteoAux", Err.Description
End Function

Public Function GrabaConteoAux(ByRef idEntrega As String, _
                             ByVal codProducto As String, _
                             ByVal CtdProducto As Integer, momento As Integer, Bloqueado As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, codProducto, CtdProducto, momento, Bloqueado)
 
    gvarIO = Array(entrada, entrada, entrada, entrada, entrada)
                
    GrabaConteoAux = gclsOracle.SP("BTLPROD.pkg_aux_conteo.sp_ingresa_conteo", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsEntrega.GrabaConteoAux", Err.Description
End Function

Public Function EditaConteoAux(ByRef idEntrega As String, _
                             ByVal codProducto As String, _
                             ByVal CtdProducto As Integer, flg_bloqueado As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, codProducto, CtdProducto, flg_bloqueado)
 
    gvarIO = Array(entrada, entrada, entrada, entrada)
                
    EditaConteoAux = gclsOracle.SP("BTLPROD.pkg_aux_conteo.sp_update_conteo", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsEntrega.EditaConteoAux", Err.Description
End Function

Public Function EliminaConteoAux(ByRef idEntrega As String, _
                             ByVal codProducto As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, codProducto)
    gvarIO = Array(entrada, entrada)
                
    EliminaConteoAux = gclsOracle.SP("BTLPROD.pkg_aux_conteo.sp_del_conteo", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsEntrega.EliminaConteoAux", Err.Description
End Function

Public Function ImprimeTransportista(ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ImprimeTransportista = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.fn_imp_transportista", 0, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ImprimeTransportista", Err.Description
End Function

Public Function Desasociar(ByVal idEntrega As String, _
                             ByVal numGuia As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, numGuia)
 
    gvarIO = Array(entrada, entrada)
    
    Desasociar = gclsOracle.SP("BTLPROD.PKG_ENTREGA.sp_desasociar_guias", gvarValores, gvarIO)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsLocal.desasociar", Err.Description
End Function

Public Function ListaProdDev(ByVal idEntrega As String, Optional ByVal codigoproducto As String = "") As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaProdDev = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.fn_lista_prod_dev", 0, idEntrega, codigoproducto)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaProdDev", Err.Description
End Function

Public Function GrabaDetDev(ByVal idEntrega As String, _
                             ByVal codProducto As String, _
                             ByVal CtdProducto As Integer, Lote As String, fecven As String, momento As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, codProducto, CtdProducto, Lote, fecven, momento)
 
    gvarIO = Array(entrada, entrada, entrada, entrada, entrada, entrada)
                
    GrabaDetDev = gclsOracle.SP("BTLPROD.pkg_aux_conteo.sp_ingresa_det_dev", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsEntrega.GrabaDetDev", Err.Description
End Function

Public Function EditaDetDev(ByVal idEntrega As String, _
                             ByVal codProducto As String, _
                             ByVal CtdProducto As Integer, Lote As String, fecven As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, codProducto, CtdProducto, Lote, fecven)
 
    gvarIO = Array(entrada, entrada, entrada, entrada, entrada)
                
    EditaDetDev = gclsOracle.SP("BTLPROD.pkg_aux_conteo.sp_update_detDev", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsEntrega.EditaDetDev", Err.Description
End Function

Public Function EliminaDetDev(ByVal idEntrega As String, _
                             ByVal codProducto As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, codProducto)
    gvarIO = Array(entrada, entrada)
                
    EliminaDetDev = gclsOracle.SP("BTLPROD.pkg_aux_conteo.sp_del_detdev", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsEntrega.EliminaDetDev", Err.Description
End Function

Public Function momentoDetDev(ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set momentoDetDev = gclsOracle.FN_Cursor("BTLPROD.pkg_aux_conteo.fn_MomentoDev", 0, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.momentoDetDev", Err.Description
End Function

Public Function ListaDetDev(ByVal idEntrega As String, momento As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaDetDev = gclsOracle.FN_Cursor("BTLPROD.pkg_aux_conteo.fn_lista_detdev", 0, idEntrega, momento)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaDetDev", Err.Description
End Function

Public Function MaxStockProd(ByVal idEntrega As String, CodProd As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set MaxStockProd = gclsOracle.FN_Cursor("BTLPROD.pkg_entrega.fn_MaxStock", 0, idEntrega, CodProd)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.MaxStockProd", Err.Description
End Function

Public Function ListaDatosProv(ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set ListaDatosProv = gclsOracle.FN_Cursor("BTLPROD.pkg_aux_conteo.fn_lista_DatosProv", 0, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.ListaDatosProv", Err.Description
End Function

Public Function UpdDevueltosRecep(ByVal idEntrega As String, ByVal CadCodigo As String, ByVal Cadcantidad As String) As oraDynaset
On Error GoTo CtrlError
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, CadCodigo, Cadcantidad)
 
    gvarIO = Array(entrada, entrada, entrada)
                    
    gclsOracle.SP "BTLPROD.PKG_ENTREGA.sp_upd_cantDev", _
                              gvarValores, _
                              gvarIO
        
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.UpdDevueltos", Err.Description
End Function

Public Function UpdTransportistaGuia(ByVal numGuia As String, ByVal numRuc As String, ByVal direccion As String, placa As String, ByVal CodLocal As String) As oraDynaset
On Error GoTo CtrlError
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(numGuia, numRuc, direccion, placa, CodLocal)
 
    gvarIO = Array(entrada, entrada, entrada, entrada, entrada)
                    
    gclsOracle.SP "BTLPROD.PKG_ENTREGA.sp_upd_Transportista", gvarValores, gvarIO
        
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.UpdTransportistaGuia", Err.Description
End Function

Public Function BuscaTransportista(ByVal numRuc As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set BuscaTransportista = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.fn_busca_Transportista", 0, numRuc)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.BuscaTransportista", Err.Description
End Function

Public Function UpdFlgGeneraGuia(ByVal idEntrega As String, ByVal tipoGuia As String) As oraDynaset
On Error GoTo CtrlError
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega, tipoGuia)
 
    gvarIO = Array(entrada, entrada)
                    
    gclsOracle.SP "BTLPROD.PKG_ENTREGA.sp_upd_GuiaDev", gvarValores, gvarIO
        
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.UpdFlgGeneraGuia", Err.Description
End Function

Public Function BuscaFlgGeneraGuia(ByVal idEntrega As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set BuscaFlgGeneraGuia = gclsOracle.FN_Cursor("BTLPROD.PKG_ENTREGA.fn_busca_flggeneraguia", 0, idEntrega)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsEntrega.BuscaFlgGeneraGuia", Err.Description
End Function

Public Function LimpiaDetDev(idEntrega As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(idEntrega)
    gvarIO = Array(entrada)
                
    LimpiaDetDev = gclsOracle.SP("BTLPROD.pkg_aux_conteo.sp_limpia_detdev", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsEntrega.LimpiaDetDev", Err.Description
End Function


