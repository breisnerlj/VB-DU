VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPromocion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ListaTipo(Optional Codigo As String = "") As oraDynaset
    Set ListaTipo = gclsOracle.FN_Cursor("BTLPROD.PKG_PROMOCION.FN_LISTA_TIPO", 0, Codigo)
End Function
Public Function ListaCriterio() As oraDynaset
    gclsOracle.ODataBase.Parameters.Add "CUR_GUIAS", "", ORAPARM_OUTPUT, ORATYPE_CURSOR
    gclsOracle.ODataBase.ExecuteSQL " BEGIN BTLPROD.PKG_PROMOCION.SP_LISTA_COMBO(:CUR_GUIAS); END;"
    Set ListaCriterio = gclsOracle.ODataBase.Parameters("CUR_GUIAS").Value
    gclsOracle.ODataBase.Parameters.Remove "CUR_GUIAS"
End Function

Public Function Lista(Optional Codigo As String = "") As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_PROMOCION.FN_LISTA", 0, Codigo)
End Function
Public Function ListaLocal(Optional Codigo As String = "", Optional Cia As String = "", Optional CodigoLocal As String = "") As oraDynaset
    Set ListaLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_PROMOCION.FN_LISTA", 0, Codigo, Cia, CodigoLocal)
End Function


Public Function ListaOrigen(Optional Codigo As String = "") As oraDynaset
    Set ListaOrigen = gclsOracle.FN_Cursor("BTLPROD.PKG_PROMOCION.FN_LISTA_ORIGEN", 0, Codigo)
End Function

Public Function ListaRegalo(Cia As String, CodigoLocal As String, Codigo As String) As oraDynaset
    Set ListaRegalo = gclsOracle.FN_Cursor("BTLPROD.PKG_PROMOCION.FN_LISTA_REGALOS", 0, Cia, CodigoLocal, Codigo)
End Function



Public Function ListaCanjeProducto(Optional CodPromocion As String = "") As oraDynaset
    On Error GoTo CtrlErr
    
    Set ListaCanjeProducto = gclsOracle.FN_Cursor("BTLPROD.PKG_PROMOCION.FN_LISTA_CANJE_PRODUCTO", 0, CodPromocion)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPromocion", Err.Description

End Function
