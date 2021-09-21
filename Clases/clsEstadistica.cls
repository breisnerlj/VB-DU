VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista(ByVal vstrCodLocal As String, _
                      ByVal vstrCodProducto As String, _
                      Optional ByVal vstrDia As String = "") As oraDynaset

    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_ESTADISTICA_VTA_SMM.FN_LISTA", 0, vstrCodLocal, vstrCodProducto, vstrDia)
End Function

Public Sub Limpia()
    
    gclsOracle.Execute "BEGIN BTLPROD.PKG_ESTADISTICA_VTA_SMM.SP_LIMPIA; END; "
                                                               
End Sub

Public Function Lista_Cantidades(ByVal vstrCodLocal As String, _
                                 ByVal vstrCodProducto As String, _
                                 Optional ByVal vstrDia As String = "") As oraDynaset

    Set Lista_Cantidades = gclsOracle.FN_Cursor("BTLPROD.PKG_ESTADISTICA_VTA_SMM.FN_LISTA2", 0, vstrCodLocal, vstrCodProducto, vstrDia)
    
End Function



