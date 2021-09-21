VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''Public Function Lista(ByVal Cia As String, vstrDato As String) As oraDynaset
'''On Error GoTo CnttlError
'''    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA", 0, Cia, vstrDato)
'''Exit Function
'''CnttlError:
'''    Err.Raise Err.Number, "clsLocal.Lista", Err.Description
'''End Function

Public Function Lista(ByVal Cia As String, _
                      Optional vstrDato As String, _
                      Optional ByVal vstrTipo As String = "", _
                      Optional ByVal Opcion As String = "0") As oraDynaset
    On Error GoTo CnttlError
    
    If Trim(vstrTipo) = "" Then
        If Opcion = "0" Then
            Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA", 0, Cia, vstrDato)
        Else
            Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA", 0, Cia, vstrDato, Opcion)
        End If
    Else
        Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_X_TIPO", 0, Cia, vstrTipo)
    End If
    
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal::Lista", Err.Description
End Function

Public Function ListaXTipo(ByVal Cia As String, vstrDato As String) As oraDynaset
On Error GoTo CnttlError
    Set ListaXTipo = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_X_TIPO", 0, Cia, vstrDato)
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal.Lista", Err.Description
End Function

Public Function Lista_Inc_Todos(ByVal Cia As String) As oraDynaset
On Error GoTo CnttlError
    Set Lista_Inc_Todos = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_INC_TODOS", 0, Cia)
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal.Lista_Inc_Todos", Err.Description
End Function

Public Function ListaXConvenio(ByVal Cia As String, vstrConvenio As String) As oraDynaset
On Error GoTo CnttlError
    Set ListaXConvenio = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_X_CONVENIO", 0, Cia, vstrConvenio)
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal.Lista", Err.Description
End Function

Public Function LstLocalSNZona(ByVal Cia As String) As oraDynaset
On Error GoTo CnttlError
    Set LstLocalSNZona = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_LOCA_SN_ZONA", 0, Cia)
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal.LstLocalSNZona", Err.Description
End Function

Public Function LstLocalCNZona(ByVal Cia As String, ByVal ZONA As String) As oraDynaset
On Error GoTo CnttlError
    Set LstLocalCNZona = gclsOracle.FN_Cursor("BTLPROD.PKG_ZONA.FN_LISTA_LOCA_CN_ZONA", 0, Cia, ZONA)
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal.LstLocalSNZona", Err.Description
End Function

Public Function ListaLocalPredetDLV(ByVal Cia As String, Optional ByVal Ubigeo As String = "") As oraDynaset
On Error GoTo CtrlErr

    Set ListaLocalPredetDLV = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_LOCAL_PREDET_DLV", 0, Cia, Ubigeo)
    
    
Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLocal.ListaLocalPredetDLV", Err.Description

End Function


Public Property Get NombreImpresoraAlmacen() As String
    NombreImpresoraAlmacen = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_IMPRESORA_ALM")
End Property




Public Function ListaXUbigeo(ByVal Cia As String, ByVal CodigoUbigeo As String, ByVal EstadoLocal As String, ByVal TipoLocal As String, ByVal Excepciones As String) As oraDynaset
On Error GoTo CnttlError

    Set ListaXUbigeo = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_UBIGEO", 0, Cia, CodigoUbigeo, EstadoLocal, TipoLocal, Excepciones)
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal.ListaXUbigeo", Err.Description
End Function

Public Function ActualizaLocalxMaquina(ByVal COD_LOCAL As String, _
                      ByVal COD_MAQUINA As String, _
                      ByVal usuario As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(COD_LOCAL, _
                         COD_MAQUINA, _
                         usuario)
                         

    gvarIO = Array(entrada, _
                    entrada, _
                    entrada)
                
                    
    ActualizaLocalxMaquina = gclsOracle.SP("BTLPROD.PKG_LOCAL.SP_ACT_LOCAL_X_MAQUINA", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsLocal.Graba", Err.Description
End Function

'ECASTILLO 06.05.2020
Public Function EvaluaLocalInkMig(ByVal COD_LOCAL As String) As String
    On Error GoTo Err
    EvaluaLocalInkMig = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_EST_LOCAL_INKFARMA_POSU", COD_LOCAL)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.Evalua", Err.Description
End Function
'ECASTILLO 27.07.2020
Public Function GetCodInka(ByVal CodLocal As String) As String
On Error GoTo Err
    GetCodInka = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_COD_LOCAL_INKFARMA", CodLocal)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.GetCodInka", Err.Description
End Function

'ECASTILLO 05.07.2020
Public Function GetCodPosu(ByVal CodLocal As String) As String
On Error GoTo Err
    GetCodPosu = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_COD_SAP_LOCAL", CodLocal)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.GetCodPosu", Err.Description
End Function

'ECASTILLO 05.07.2020
Public Function GetCodBTL(ByVal CodLocal As String) As String
On Error GoTo Err
    GetCodBTL = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_COD_LOCAL", CodLocal)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.GetCodBTL", Err.Description
End Function

'ECASTILLO 30.07.2020
Public Function GetEstConfig(ByVal Cia As String, ByVal CodLocal As String, ByVal config As String) As String
On Error GoTo Err
    GetEstConfig = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_EST_LOCAL_FUNCIONALIDAD", Cia, CodLocal, config)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.GetEstConfig", Err.Description
End Function

'ECASTILLO 22.10.2020
Public Function GetIndCDCAP(ByVal CodLocal As String) As String
On Error GoTo Err
    GetIndCDCAP = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_IND_CD_CAPA_BTL", CodLocal)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.GetIndCDCAP", Err.Description
End Function


'ECASTILLO 17.09.2021
Public Function GetMarcaLocal(ByVal CodLocal As String, ByVal pTipo As String) As String
On Error GoTo Err
    GetMarcaLocal = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_MARCA_LOCAL", CodLocal, pTipo)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.GetMarcaLocal", Err.Description
End Function

'ECASTILLO 17.09.2021
Public Function GetMarcaLocalPosu(ByVal CodLocal As String, ByVal pTipo As String) As String
On Error GoTo Err
    GetMarcaLocalPosu = "" & gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_GET_MARCA_LOCAL_POSU", CodLocal, pTipo)
    Exit Function
Err:
    Err.Raise Err.Number, "clsLocal.GetMarcaLocalPosu", Err.Description
End Function
