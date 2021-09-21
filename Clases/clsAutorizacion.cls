VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Lista(Optional ByVal strCodAutorizacion As String = "") As OraDynaset

    On Error GoTo CtrlErr
        Set Lista = gclsOracle.FN_Cursor("CMR.PKG_AUTORIZACION.FN_LISTA", 0, strCodAutorizacion)

      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAutorizacion.Lista", Err.Description
End Function


Public Function AutorizaUsuario(ByVal strCodAutorizacion As String) As OraDynaset
    On Error GoTo CtrlErr
        Set AutorizaUsuario = gclsOracle.FN_Cursor("CMR.PKG_AUTORIZACION.FN_AUTORIZ_USUARIO", 0, strCodAutorizacion)
        
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAutorizacion.AutorizaUsuario", Err.Description
End Function

Public Function ListaPermisos(ByVal Aplicacion As String, _
                              ByVal Usuario As String, _
                              Optional ByVal Padre As String = "") As OraDynaset
    On Error GoTo CtrlErr
        If Padre = "" Then
            Set ListaPermisos = gclsOracle.FN_Cursor("NUEVO.PKG_PERMISO.FN_LISTA_MENU", 0, Aplicacion, Usuario)
        Else
            Set ListaPermisos = gclsOracle.FN_Cursor("NUEVO.PKG_PERMISO.FN_LISTA_MENU_HIJO", 0, Aplicacion, Usuario, Padre)
        End If
        
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAutorizacion.Lista", Err.Description
End Function

