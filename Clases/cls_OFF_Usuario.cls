VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_OFF_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------
'Autor : Juan Arturo Escate Espichan
'Fecha : 07/05/2008
'----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private strCia As String
Private strCodLocal As String
Private strCodMaquina As String
Private strCodigoUsuario As String
Private strSYSDATE As String
Private strNombreUsuario As String
Private strCodigoPerfil As String
Private strDescripcionPerfil As String
Private strFlgContingencia As Boolean
Private strEmpresa As String
Private strSmbMoneda As String
Private strDesLgMoneda As String
Private strRucEmpresa As String
Private strDireccionEmpresa As String
Private strDireccionLocal As String
Private strCodSerieEtiq As String
Private strUsuModPrecio As String
Private objIni As New cls_ArchivoIni
Private xMoneda As New XArrayDB
Private strTipoCambio As String
Private strUltDocEmi As String



Public Function Login(ByVal CodigoUsuario As String) As Boolean
On Error GoTo handle

    Dim objDocumento As cls_OFF_Documento
    Dim rs As New ADODB.Recordset
    Dim strSQL  As String
    Dim a As Long
    
    If Len(CodigoUsuario) < 1 Or Len(CodigoUsuario) > 5 Then GoTo final

    Set objDocumento = New cls_OFF_Documento

    Call objDocumento.CreaArchivoTexto
    Set objDocumento = Nothing

'    strSql = "Select * from usuario.txt "
'    rs.Open strSql, gstrConexion, adOpenStatic, adLockOptimistic, adCmdText
    rs.CursorLocation = adUseClient
    rs.Open strUsuariosXML, gstrConexion, adOpenForwardOnly, adLockReadOnly
    
    rs.Filter = "COD_USUARIO = '" & CodigoUsuario & "'"
        
    If rs.RecordCount = 1 Then
        strCia = objIni.LeerIni(gstrIni, "general", "CIA", "")
        strCodLocal = objIni.LeerIni(gstrIni, "general", "LOCAL", "")
        strFlgContingencia = IIf(Val(objIni.LeerIni(gstrIni, "general", "FLG_CONTINGENCIA", "0")) = 1, True, False)
        strNombreUsuario = "" & rs("DES_USUARIO").Value
        strCodigoPerfil = Format("" & rs("COD_PERFIL").Value, "0000")
        strDescripcionPerfil = "" & rs("DES_PERFIL").Value
        strDescripcionPerfil = "" & rs("DES_PERFIL").Value
        strTipoCambio = "" & rs("TIPO_CAMBIO").Value
        strEmpresa = objIni.LeerIni(gstrIni, "general", "EMPRESA", "")
        strRucEmpresa = objIni.LeerIni(gstrIni, "general", "RUC", "")
        strDireccionEmpresa = objIni.LeerIni(gstrIni, "general", "DIRECCION", "")
        strDireccionLocal = objIni.LeerIni(gstrIni, "general", "DIRECCION_LOCAL", "")
        strCodSerieEtiq = objIni.LeerIni(gstrIni, "general", "COD_SERIE_ETIQ", "")
        strSYSDATE = Format(Now, "dd/mm/yyyy HH:mm:ss")
        strCodigoUsuario = CodigoUsuario
        strCodMaquina = sGetNombrePC
        CargaDatosMoneda
        Login = True
    Else
        Login = False
    End If
    Set rs = Nothing
    
    Exit Function
handle:
    Err.Raise Err.Number, "cls_OFF_Usuario", Err.Description
final:
    Set rs = Nothing
    Login = False

End Function

Public Property Get Cia() As String
    Cia = strCia
End Property

Public Property Let Cia(ByVal lstrCia As String)
    strCia = lstrCia
End Property


Public Property Get CodLocal() As String
    CodLocal = strCodLocal
End Property

Public Property Let CodLocal(ByVal lstrCodLocal As String)
    strCodLocal = lstrCodLocal
End Property

Public Property Get CodMaquina() As String
CodMaquina = strCodMaquina

End Property

Public Property Let CodMaquina(ByVal lstrCodMaquina As String)
strCodMaquina = lstrCodMaquina
End Property


Public Property Get CodigoUsuario() As String
CodigoUsuario = strCodigoUsuario

End Property

Public Property Let CodigoUsuario(ByVal lstrCodigoUsuario As String)
strCodigoUsuario = lstrCodigoUsuario
End Property


Public Property Get NombreUsuario() As String
NombreUsuario = strNombreUsuario

End Property

Public Property Let NombreUsuario(ByVal lstrNombreUsuario As String)
strNombreUsuario = lstrNombreUsuario
End Property

Public Property Get sysdate() As String
sysdate = strSYSDATE

End Property

Public Property Let sysdate(ByVal lstrSYSDATE As String)
strSYSDATE = lstrSYSDATE
End Property


Private Sub Class_Initialize()
    BuscaDatosMoneda
End Sub

Private Sub Class_Terminate()
    Set objIni = Nothing
End Sub

Public Property Get CodigoPerfil() As String
CodigoPerfil = strCodigoPerfil

End Property

Public Property Let CodigoPerfil(ByVal lstrCodigoPerfil As String)
    strCodigoPerfil = lstrCodigoPerfil
End Property


Public Property Get DescripcionPerfil() As String
    DescripcionPerfil = strDescripcionPerfil
End Property

Public Property Let DescripcionPerfil(ByVal lstrDescripcionPerfil As String)
    strDescripcionPerfil = lstrDescripcionPerfil
End Property
'''
Public Property Get Contingencia() As Boolean
    Contingencia = strFlgContingencia
End Property

Public Property Let Contingencia(ByVal lstrContingencia As Boolean)
    strFlgContingencia = lstrContingencia
End Property




Public Property Get Empresa() As String
    Empresa = strEmpresa
End Property

Public Property Let Empresa(ByVal lstrEmpresa As String)
    strEmpresa = lstrEmpresa
End Property



Public Property Get SmbMoneda() As String
    SmbMoneda = strSmbMoneda
End Property

Public Property Let SmbMoneda(ByVal lstrSmbMoneda As String)
    strSmbMoneda = lstrSmbMoneda
End Property



Public Property Get DesLgMoneda() As String
    DesLgMoneda = strDesLgMoneda
End Property

Public Property Let DesLgMoneda(ByVal lstrDesLgMoneda As String)
    strDesLgMoneda = lstrDesLgMoneda
End Property



Private Sub BuscaDatosMoneda()
Dim arrCodMoneda As Variant
Dim arrDesMoneda As Variant
Dim arrSmbMoneda As Variant


arrCodMoneda = Split(objIni.LeerIni(gstrIni, "MONEDA", "COD_MONEDA", ""), ",")
arrDesMoneda = Split(objIni.LeerIni(gstrIni, "MONEDA", "DES_MONEDA", ""), ",")
arrSmbMoneda = Split(objIni.LeerIni(gstrIni, "MONEDA", "SMB_MONEDA", ""), ",")
Set xMoneda = ArrayaXarray(arrCodMoneda, arrDesMoneda, arrSmbMoneda)


End Sub


Public Property Get Moneda() As XArrayDB
    Set Moneda = xMoneda
End Property




Public Property Get RucEmpresa() As String
    RucEmpresa = strRucEmpresa
End Property

Public Property Let RucEmpresa(ByVal lstrRucEmpresa As String)
    strRucEmpresa = lstrRucEmpresa
End Property




Public Property Get DireccionEmpresa() As String
    DireccionEmpresa = strDireccionEmpresa
End Property

Public Property Let DireccionEmpresa(ByVal lstrDireccionEmpresa As String)
    strDireccionEmpresa = lstrDireccionEmpresa
End Property

Public Property Get DireccionLocal() As String
    DireccionLocal = strDireccionLocal
End Property

Public Property Let DireccionLocal(ByVal lstrDireccionLocal As String)
    strDireccionLocal = lstrDireccionLocal
End Property


Public Property Get CodSerieEtiq() As String
    CodSerieEtiq = strCodSerieEtiq
End Property

Public Property Let CodSerieEtiq(ByVal lstrCodSerieEtiq As String)
    strCodSerieEtiq = lstrCodSerieEtiq
End Property


Public Property Get TipoCambio() As String
    TipoCambio = strTipoCambio
End Property

Public Property Let TipoCambio(ByVal lstrTipoCambio As String)
    strTipoCambio = lstrTipoCambio
End Property

Public Function DesUsuarioVenta(ByVal pstrCodUsuarioVenta) As String
Dim cnn As ADODB.Connection
Dim strSQL As String
Dim rsUsuarioVenta As New ADODB.Recordset

On Error GoTo CtrlErr

'        Set cnn = New ADODB.Connection
'        cnn.Open gstrConexion
        
'        strSql = "select * from usuario.txt"
'        rsUsuarioVenta.Open strSql, cnn, adOpenStatic, adLockOptimistic, adCmdText
        rsUsuarioVenta.CursorLocation = adUseClient
        rsUsuarioVenta.Open strUsuariosXML, gstrConexion, adOpenForwardOnly, adLockReadOnly
        rsUsuarioVenta.Filter = "COD_USUARIO = '" & pstrCodUsuarioVenta & "'"
        If rsUsuarioVenta.RecordCount > 0 Then
            DesUsuarioVenta = rsUsuarioVenta!DES_USUARIO
        Else
            Err.Raise vbObjectError + 513, "cls_OFF_Documento.ImprimePorDocumento", "No se encontró usuario de la venta " & pstrCodUsuarioVenta
        End If
        rsUsuarioVenta.Close
'        cnn.Close
Exit Function

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Usuario.DesUsuarioVenta", Err.Description

End Function

Public Property Get UsuModPrecio() As String
    UsuModPrecio = strUsuModPrecio
End Property

Public Property Let UsuModPrecio(ByVal pstrUsuModPrecio As String)
    strUsuModPrecio = pstrUsuModPrecio
End Property


Public Sub CargaDatosMoneda()

Dim a As Long

On Error GoTo CtrlErr

    a = Moneda.Find(0, 0, "1")
    strDesLgMoneda = Moneda(a, 1)
    strSmbMoneda = Moneda(a, 2)

Exit Sub
CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Usuario.DevMoneda", Err.Description
End Sub


Public Function DescripcionMoneda(ByVal strCodMoneda As String)
Dim a As Long


On Error GoTo CtrlErr

    a = Moneda.Find(0, 0, strCodMoneda)
    DescripcionMoneda = Moneda(a, 1)
    

Exit Function
CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Usuario.DevMoneda", Err.Description


End Function

Public Property Get UltDocEmi() As String
    UltDocEmi = strUltDocEmi
End Property

Public Property Let UltDocEmi(ByVal lstrUltDocEmi As String)
    strUltDocEmi = lstrUltDocEmi
End Property


Public Sub ActualizaUltDocEmitido(Optional ByVal strCodTipoDocumento As String = "")
Dim objArchivoIni As cls_ArchivoIni

On Error GoTo CtrlErr

    Set objArchivoIni = New cls_ArchivoIni

    
    If strCodTipoDocumento = "" Then
        UltDocEmi = objIni.LeerIni(gstrIni, "general", "ULT_DOC_EMITIDO", "")
    Else
        objArchivoIni.GuardarIni gstrIni, "general", "ULT_DOC_EMITIDO", strCodTipoDocumento
        UltDocEmi = strCodTipoDocumento
    End If
    
    Set objArchivoIni = Nothing
    
    

Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Usuario.ActualizaUltDocEmitido", Err.Description

End Sub


Public Function AgregaMoneda(ByVal pstrCodigo As String, _
                    ByVal pstrDescripcion As String, _
                    ByVal pstrSmb As String) As XArrayDB
    


    Dim ultimo As Integer
    Dim aux As Integer
    If xMoneda.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xMoneda.Count(1)
    While i < aux
        If xMoneda(i, 0) = pstrCodigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xMoneda.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xMoneda.AppendRows
    End If
j:
    If xMoneda.Count(1) = 0 Then ultimo = 0: xMoneda.AppendRows
    
    xMoneda(ultimo, 0) = pstrCodigo
    xMoneda(ultimo, 1) = pstrDescripcion
    xMoneda(ultimo, 2) = pstrSmb
    Set AgregaMoneda = xMoneda




End Function
