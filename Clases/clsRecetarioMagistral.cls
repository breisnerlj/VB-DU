VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecetarioMagistral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaUndCapacidad() As oraDynaset
    
    On Error GoTo CntrlError
    
    Set ListaUndCapacidad = gclsOracle.FN_Cursor("BTLPROD.PKG_RECETARIO_MAGISTRAL.FN_LST_UND_CAPACIDAD", 0)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsRecetarioMagistral.ListaUndCapacidad", Err.Description
    
End Function



Public Function ListaTipoInsumo() As oraDynaset
    
    On Error GoTo CntrlError
    
    Set ListaTipoInsumo = gclsOracle.FN_Cursor("BTLPROD.PKG_RECETARIO_MAGISTRAL.FN_LST_TIPO_INSUMO", 0)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsRecetarioMagistral.ListaTipoInsumo", Err.Description
    
End Function


Public Property Get CodInsumoBase() As String
    CodInsumoBase = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_INSUMO_BASE")
End Property
