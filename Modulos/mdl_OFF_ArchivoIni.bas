Attribute VB_Name = "mdl_OFF_ArchivoIni"
Option Explicit

Private strArchivo As String
Private strFlgContingencia As String
Private strAdminUser As String
Private strAdminPass As String
Private strGlosa As String
Private strSecVenta As String
Private strCia As String
Private strCodLocal As String
Private strDocumentos As String
Private strNumFac As String
Private intLinFac As Integer
Private strDesFac As String
Private strNumBol As String
Private intLinBol As Integer
Private strDesBol As String
'-------------------------------------------------------------------------------
' PROPIEDADES
'-------------------------------------------------------------------------------
Public Property Get Archivo() As String
    Archivo = strArchivo
End Property

Public Property Let Archivo(ByVal newValue As String)
    strArchivo = newValue
End Property

Public Property Get FlgContingencia() As String
    FlgContingencia = strFlgContingencia
End Property

Public Property Let FlgContingencia(ByVal newValue As String)
    strFlgContingencia = newValue
End Property

Public Property Get AdminUser() As String
    AdminUser = strAdminUser
End Property

Public Property Let AdminUser(ByVal newValue As String)
    strAdminUser = newValue
End Property

Public Property Get AdminPass() As String
    AdminPass = strAdminPass
End Property

Public Property Let AdminPass(ByVal newValue As String)
    strAdminPass = newValue
End Property

Public Property Get Glosa() As String
    Glosa = strGlosa
End Property

Public Property Let Glosa(ByVal newValue As String)
    strGlosa = newValue
End Property

Public Property Get SecVenta() As String
    SecVenta = strSecVenta
End Property

Public Property Let SecVenta(ByVal newValue As String)
    strSecVenta = newValue
End Property

Public Property Get Cia() As String
    Cia = strCia
End Property

Public Property Let Cia(ByVal newValue As String)
    strCia = newValue
End Property

Public Property Get CodLocal() As String
    CodLocal = strCodLocal
End Property

Public Property Let CodLocal(ByVal newValue As String)
    strCodLocal = newValue
End Property

Public Property Get Documentos() As String
    Documentos = strDocumentos
End Property

Public Property Let Documentos(ByVal newValue As String)
    strDocumentos = newValue
End Property

Public Property Get NumFac() As String
    NumFac = strNumFac
End Property

Public Property Let NumFac(ByVal newValue As String)
    strNumFac = newValue
End Property

Public Property Get LinFac() As Integer
    LinFac = intLinFac
End Property

Public Property Let LinFac(ByVal newValue As Integer)
    intLinFac = newValue
End Property

Public Property Get DesFac() As String
    DesFac = strDesFac
End Property

Public Property Let DesFac(ByVal newValue As String)
    strDesFac = newValue
End Property

Public Property Get NumBol() As String
    NumBol = strNumBol
End Property

Public Property Let NumBol(ByVal newValue As String)
    strNumBol = newValue
End Property

Public Property Get LinBol() As Integer
    LinBol = intLinBol
End Property

Public Property Let LinBol(ByVal newValue As Integer)
    intLinBol = newValue
End Property

Public Property Get DesBol() As String
    DesBol = strDesBol
End Property

Public Property Let DesBol(ByVal newValue As String)
    strDesBol = newValue
End Property


'-------------------------------------------------------------------------------
' METODOS
'-------------------------------------------------------------------------------
Public Sub CargarArchivoIni()
    
    Dim strSeccion As String
    
    Dim objArchivoIni As cls_ArchivoIni
    Set objArchivoIni = New cls_ArchivoIni
    
    If Archivo = "" Then
        Set objArchivoIni = Nothing
        Exit Sub
    End If
    
    'Seccion GENERAL
    strSeccion = "GENERAL"
    FlgContingencia = objArchivoIni.LeerIni(Archivo, strSeccion, "FLG_CONTINGENCIA")
    AdminUser = objArchivoIni.LeerIni(Archivo, strSeccion, "ADMIN_USER")
    AdminPass = objArchivoIni.LeerIni(Archivo, strSeccion, "ADMIN_PASS")
    SecVenta = objArchivoIni.LeerIni(Archivo, strSeccion, "SEC_VENTA")
    Cia = objArchivoIni.LeerIni(Archivo, strSeccion, "CIA")
    CodLocal = objArchivoIni.LeerIni(Archivo, strSeccion, "LOCAL")
    Documentos = objArchivoIni.LeerIni(Archivo, strSeccion, "DOCUMENTOS")
    
    'Seccion DOCUMENTOS
    strSeccion = "DOCUMENTOS"
    NumFac = objArchivoIni.LeerIni(Archivo, strSeccion, "NUM_FAC")
    LinFac = objArchivoIni.LeerIni(Archivo, strSeccion, "LIN_FAC")
    DesFac = objArchivoIni.LeerIni(Archivo, strSeccion, "DES_FAC")
    NumBol = objArchivoIni.LeerIni(Archivo, strSeccion, "NUM_BOL")
    LinBol = objArchivoIni.LeerIni(Archivo, strSeccion, "LIN_BOL")
    DesBol = objArchivoIni.LeerIni(Archivo, strSeccion, "DES_BOL")
    
    Set objArchivoIni = Nothing
    
End Sub

Public Sub GuardarArchivoIni()

    Dim strSeccion As String

    Dim objArchivoIni As cls_ArchivoIni
    Set objArchivoIni = New cls_ArchivoIni
    
    If Archivo = "" Then
        Set objArchivoIni = Nothing
        Exit Sub
    End If
    
    'Seccion GENERAL
    strSeccion = "GENERAL"
    objArchivoIni.GuardarIni Archivo, strSeccion, "FLG_CONTINGENCIA", FlgContingencia
    objArchivoIni.GuardarIni Archivo, strSeccion, "ADMIN_USER", AdminUser
    objArchivoIni.GuardarIni Archivo, strSeccion, "ADMIN_PASS", AdminPass
    objArchivoIni.GuardarIni Archivo, strSeccion, "SEC_VENTA", SecVenta
    objArchivoIni.GuardarIni Archivo, strSeccion, "CIA", Cia
    objArchivoIni.GuardarIni Archivo, strSeccion, "LOCAL", CodLocal
    objArchivoIni.GuardarIni Archivo, strSeccion, "DOCUMENTOS", Documentos
    
    'Seccion DOCUMENTOS
    strSeccion = "DOCUMENTOS"
    objArchivoIni.GuardarIni Archivo, strSeccion, "NUM_FAC", NumFac
    objArchivoIni.GuardarIni Archivo, strSeccion, "LIN_FAC", LinFac
    objArchivoIni.GuardarIni Archivo, strSeccion, "DES_FAC", DesFac
    objArchivoIni.GuardarIni Archivo, strSeccion, "NUM_BOL", NumBol
    objArchivoIni.GuardarIni Archivo, strSeccion, "LIN_BOL", LinBol
    objArchivoIni.GuardarIni Archivo, strSeccion, "DES_BOL", DesBol
    
    Set objArchivoIni = Nothing

End Sub


