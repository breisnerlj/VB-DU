VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCntCompetencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Grabar(ByRef NumeroControl As String, _
                       ByVal CodigoLocal As String, _
                       ByVal Observacion As String, _
                       ByVal CodigoUsuario As String, _
                       ByVal ValePromedio As String, _
                       ByVal CadRucCompetencia As String, _
                       ByVal CadLocalCompetencia As String, _
                       ByVal CadCaja As String, _
                       ByVal CadNumDoc1 As String, _
                       ByVal CadNumDoc2 As String, _
                       ByVal CadNumDoc3 As String, _
                       ByVal CadNumDoc4 As String, _
                       ByVal CadComentario As String, _
                       ByVal FechaControl As String _
                    ) As String
                    Dim strNumeroControl As String
      strNumeroControl = gclsOracle.FN_Valor("BTLPROD.PKG_CNT_COMPETENCIA.FN_GRABA", NumeroControl, CodigoLocal, Observacion, CodigoUsuario, ValePromedio, CadRucCompetencia, CadLocalCompetencia, CadCaja, CadNumDoc1, CadNumDoc2, CadNumDoc3, CadNumDoc4, CadComentario, FechaControl)
      NumeroControl = strNumeroControl
      Grabar = strNumeroControl
      
End Function

Public Function Lista(ByVal CodigoLocal As String, _
                       ByVal CodigoPeriodo As String, ByVal FechaInicio As String, ByVal FechaFin As String) As oraDynaset
      Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_CNT_COMPETENCIA.FN_LISTA", 0, CodigoLocal, CodigoPeriodo, FechaInicio, FechaFin)
End Function

Public Function ListaCompetencia(ByVal CodigoLocal As String, Optional ByVal CodigoProveedor As String) As oraDynaset
      Set ListaCompetencia = gclsOracle.FN_Cursor("BTLPROD.PKG_CNT_COMPETENCIA.FN_LISTA_COMPETENCIA", 0, CodigoLocal, CodigoProveedor)
End Function

Public Function ListaCajas(ByVal CodigoLocal As String, Optional ByVal CodigoProveedor As String) As oraDynaset
      Set ListaCajas = gclsOracle.FN_Cursor("BTLPROD.PKG_CNT_COMPETENCIA.fn_lista_cajas", 0, CodigoLocal, CodigoProveedor)
End Function

Public Function ListaEstadistica(ByVal NumeroCompetencia As String, ByVal Periodo As String) As oraDynaset
      Set ListaEstadistica = gclsOracle.FN_Cursor("BTLPROD.PKG_CNT_COMPETENCIA.FN_LISTA_ESTADISTICA", 0, NumeroCompetencia, Periodo, objUsuario.CodigoLocal)
End Function

Public Function DatosCaja(ByVal CodigoEmpresa As String, ByVal CodigoLocal As String, ByVal CodigoCaja As String, ByVal Fecha As String) As oraDynaset
      Set DatosCaja = gclsOracle.FN_Cursor("BTLPROD.PKG_CNT_COMPETENCIA.FN_LISTA_CAJA", 0, CodigoEmpresa, CodigoLocal, CodigoCaja, Fecha)
End Function

Public Function Anula(ByVal CodigoUsuario As String, _
                      ByVal NumeroControl As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    
    gvarValores = Array(NumeroControl, CodigoUsuario)
    gvarIO = Array(entrada, entrada)

    Anula = gclsOracle.SP("BTLPROD.PKG_CNT_COMPETENCIA.SP_ANULA", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsCntCompetencia.Anula", Err.Description
End Function
