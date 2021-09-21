VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ArchivoIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If Win32 Then
    'Declaraciones para 32 bits
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
    'Declaraciones para 16 bits
    Private Declare Function GetPrivateProfileString Lib "Kernel" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "Kernel" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lplFileName As String) As Integer
#End If


'----------------------------------------------------------------------------
'Función equivalente a GetSetting de VB4.
'GetSetting     En VB4/32bits usa el registro.
'               En VB4/16bits usa un archivo de texto.
'Pero al usar las llamadas del API, siempre se escriben en archivos de texto.
'----------------------------------------------------------------------------
Public Function LeerIni(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
    'Los parámetros son:
    'lpFileName:    La Aplicación (fichero INI)
    'lpAppName:     La sección que suele estar entrre corchetes
    'lpKeyName:     Clave
    'vDefault:      Valor opcional que devolverá
    '               si no se encuentra la clave.
    '
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String

    'Si no se especifica el valor por defecto,
    'asignar incialmente una cadena vacía
    
    
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If

    sRetVal = String$(255, 0)

    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        LeerIni = lpString
    Else
        LeerIni = left(sRetVal, LTmp)
    End If
    
    
    Exit Function
    
CtrlErr:
    Err.Raise Err.Number, "cls_ArchivoIni.LeerIni", Err.Description

End Function


'----------------------------------------------------------------------------
'Procedimiento equivalente a SaveSetting de VB4.
'SaveSetting    En VB4/32bits usa el registro.
'               En VB4/16bits usa un archivo de texto.
'Pero al usar las llamadas del API, siempre se escriben en archivos de texto.
'----------------------------------------------------------------------------
Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    'Guarda los datos de configuración
    'Los parámetros son los mismos que en LeerIni
    'Siendo lpString el valor a guardar
    '
    Dim LTmp As Long

    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub

