VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHelpDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

''''Public Function Grabar(ByVal CodigoLocal As String, _
''''                       ByVal CodigoCategoria As String, _
''''                       ByVal CodigoSubCategoria As String, _
''''                       ByVal Prioridad As String, _
''''                       ByVal Mensaje As String, _
''''                       ByVal CodigoUsuario As String, _
''''                       ByRef NumeroTicket As String _
''''                       ) As String
''''    On Error GoTo pase
''''    Dim arrValores As Variant
''''    Dim arrDireccion As Variant
''''
''''    arrValores = Array(CodigoLocal, CodigoCategoria, CodigoSubCategoria, Prioridad, CodigoUsuario, Mensaje, NumeroTicket)
''''    arrDireccion = Array(entrada, entrada, entrada, entrada, entrada, entrada, entrada, Salida)
''''
''''    Grabar = gclsOracle.SP("BTLPROD.PKG_HELPDESK.SP_GRABA", arrValores, arrDireccion)
''''
''''    Exit Function
''''pase:
''''    Err.Raise Err.Number, "clsCobranza", Err.Description
''''End Function
Public Function Grabar(ByVal CodigoLocal As String, _
                       ByVal CodigoCategoria As String, _
                       ByVal CodigoSubCategoria As String, _
                       ByVal Prioridad As String, _
                       ByVal Mensaje As String, _
                       ByVal CodigoUsuario As String, _
                       ByRef NumeroTicket As String _
                    ) As String
      Grabar = gclsOracle.FN_Valor("BTLPROD.PKG_HELPDESK.FN_GRABA", CodigoLocal, CodigoCategoria, CodigoSubCategoria, Prioridad, CodigoUsuario, Mensaje)
End Function

Public Function ListaCategoria(Optional CodigoCategoria As String = "") As oraDynaset
      Set ListaCategoria = gclsOracle.FN_Cursor("BTLPROD.PKG_HELPDESK.FN_LISTA_CATEGORIA", 0, CodigoCategoria)
End Function
Public Function ListaSubCategoria(Optional CodigoCategoria As String, Optional CodigoSubCategoria As String = "") As oraDynaset
      Set ListaSubCategoria = gclsOracle.FN_Cursor("BTLPROD.PKG_HELPDESK.FN_LISTA_SUBCATEGORIA", 0, CodigoCategoria, CodigoSubCategoria)
End Function
Public Function ListaPrioridad() As oraDynaset
      Set ListaPrioridad = gclsOracle.FN_Cursor("BTLPROD.PKG_HELPDESK.FN_LISTA_PRIORIDAD", 0, "")
End Function

