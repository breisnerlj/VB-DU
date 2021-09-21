Attribute VB_Name = "mdlIni"
Option Explicit
'------
Declare Function GetComputerName Lib "KERNEL32" Alias _
       "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) _
       As Long
Declare Function GetPrivateProfileInt Lib "KERNEL32" Alias "GetPrivateProfileIntA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, _
          ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
          ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" ( _
         ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

Function pfstr_Leer_Cadena_INI(ByVal vstrNombreSeccion$, ByVal vstrNombreCampo$, ByVal vstrDirectorioINI$, Optional ByVal vstrDefecto$) As String
Dim strAuxiliar As String * 400, lngRes As Long
    lngRes = GetPrivateProfileString(vstrNombreSeccion, vstrNombreCampo, vstrDefecto, strAuxiliar, Len(strAuxiliar), vstrDirectorioINI)
    pfstr_Leer_Cadena_INI = left$(strAuxiliar, lngRes)
End Function

Function pflng_Escribir_Cadena_INI(ByVal vstrNombreSeccion$, ByVal vstrNombreCampo$, ByVal vstrValor$, ByVal vstrDirectorioINI$) As Long
    pflng_Escribir_Cadena_INI = WritePrivateProfileString(vstrNombreSeccion, vstrNombreCampo, vstrValor, vstrDirectorioINI)
End Function

Function pfstr_GetNombrePC() As String
Dim strNombre As String
Dim intLength As Integer
Dim lngLength As Long
 
   strNombre = Space$(256)
   lngLength = 255
   intLength = GetComputerName(strNombre, lngLength)
   pfstr_GetNombrePC = left(strNombre, lngLength)
   strNombre = ""
End Function

