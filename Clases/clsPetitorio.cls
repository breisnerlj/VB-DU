VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPetitorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista_Convenio(Optional ByVal CodigoPetitorio As String = "") As oraDynaset
    On Error GoTo CtrlErr
    If CodigoPetitorio = "" Then
        Set Lista_Convenio = gclsOracle.FN_Cursor("BTLPROD.PKG_PETITORIO.FN_LISTA_CONVENIOS", 0)
    Else
        Set Lista_Convenio = gclsOracle.FN_Cursor("BTLPROD.PKG_PETITORIO.FN_LISTA_CONVENIO", 0, CodigoPetitorio)
    End If
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPetitorio.Lista_Convenio", Err.Description
End Function

Public Function Lista_Petitorio_Convenio(ByVal vstrCodConvenio As String) As oraDynaset
    On Error GoTo CtrlErr
     Set Lista_Petitorio_Convenio = gclsOracle.FN_Cursor("BTLPROD.PKG_PETITORIO.FN_LISTA_PETITORIO_X_CNV", 0, vstrCodConvenio)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPetitorio.Lista_Petitorio_Convenio", Err.Description
End Function

Public Function Graba_Petitorio_Cnv(ByVal vstrCodPetitorio As String, _
                                    ByVal vstrCodProducto As String, _
                                    ByVal vstrCodUsuario As String, _
                                    ByVal vstrFchRegistra As String)
    On Error GoTo CtrlErr
    Dim arrValores As Variant
    Dim arrDireccion As Variant
        
    arrValores = Array(vstrCodPetitorio, vstrCodProducto, _
                       vstrCodUsuario, vstrFchRegistra)
                       
    arrDireccion = Array(entrada, entrada, _
                         entrada, entrada)
    
    Graba_Petitorio_Cnv = gclsOracle.SP("BTLPROD.PKG_PETITORIO.SP_AGREGA_PRODUCTO_PETITORIO", arrValores, arrDireccion)
    
    Exit Function
CtrlErr:
    Graba_Petitorio_Cnv = "Error al adicionar el petitorio al convenio"
End Function

Public Function HabilitaProducto(ByVal CodigoConvenio As String, _
                                    ByVal CodigoProducto As String, _
                                    ByVal CodigoPetitorio As String, _
                                    ByVal CodigoUsuario As String)
    On Error GoTo CtrlErr
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim textoHtml As String
        
    arrValores = Array(CodigoConvenio, CodigoProducto, _
                       CodigoPetitorio, CodigoUsuario, textoHtml)
                       
    arrDireccion = Array(entrada, entrada, _
                         entrada, entrada, Salida)
    
    HabilitaProducto = gclsOracle.SP("BTLPROD.PKG_PETITORIO.SP_HABILITA_PRODUCTO", arrValores, arrDireccion)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPetitorio.Graba_Petitorio_Cnv", Err.Description
End Function



