VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public vCodMedico As String

Public Function Lista(ByVal vstrDato As String) As oraDynaset
    On Error GoTo CnttlError
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_MEDICO.FN_LISTA", 0, vstrDato)
    
    Exit Function
CnttlError:
    Err.Raise Err.Number, "clsMedico.Lista", Err.Description
    
End Function


Public Function ListaCMP(ByVal vstrDato As String) As oraDynaset
    On Error GoTo CnttlError
    
    Set ListaCMP = gclsOracle.FN_Cursor("BTLPROD.PKG_MEDICO.FN_LISTA_CMP", 0, vstrDato)
    
    Exit Function
CnttlError:
    Err.Raise Err.Number, "clsMedico.ListaCMP", Err.Description
    
End Function

Public Function Graba(ByVal vstrCodMedico As String, _
                               ByVal vstrNombre As String, _
                               ByVal vstrApellido As String, _
                               ByVal vstrDireccion As String, _
                               ByVal vstrNumCMP As String, _
                               ByVal vstrSexo As String, _
                               ByVal vstrFchNacimiento As String, _
                               ByVal vintActivo As String, _
                               ByVal vstrUSUARIO As String)
   On Error GoTo CtrlErr
   
   Dim gvarValores As Variant
   Dim gvarIO  As Variant
   
   
   
   gvarValores = Array(vstrCodMedico, vstrNombre, _
                       vstrApellido, vstrDireccion, _
                       vstrNumCMP, vstrSexo, _
                       vstrFchNacimiento, vintActivo, _
                       vstrUSUARIO)
                          
   gvarIO = Array(entrada, entrada, _
                  entrada, entrada, _
                  entrada, entrada, _
                  entrada, entrada, _
                  entrada)
                  
   Graba = gclsOracle.SP("BTLPROD.PKG_MEDICO.SP_GRABA", _
                                    gvarValores, _
                                    gvarIO)
                  
   Exit Function
CtrlErr:
   Err.Raise Err.Number, "clsMedico.Graba", Err.Description
        
        
End Function

Public Function buscaCMP(ByVal CMP As String) As String
On Error GoTo CtrlErr
    buscaCMP = "" & gclsOracle.FN_Valor("BTLPROD.PKG_MEDICO.FN_BUSQUEDA_MEDICO", CMP)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function ListaPorCmp(ByVal CMP As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaPorCmp = gclsOracle.FN_Cursor("BTLPROD.PKG_MEDICO.FN_LISTA_X_CMP", 1, CMP)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function getDatosMedico(ByVal CodMedico As String) As String
On Error GoTo CtrlErr
    getDatosMedico = "" & gclsOracle.FN_Valor("BTLPROD.PKG_MEDICO.FN_DATOS_MEDICO", CodMedico)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "Error", Err.Description
End Function
