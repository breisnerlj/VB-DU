Attribute VB_Name = "mdlVariables"
Option Explicit
'**Modified Martinnets-Resalta Objeto
Public gb_resaltar As Byte
Public gb_Salir As Boolean
Public objresalta As Object
'******************
'**Modified Martinnets -Scotiabank
''''''Public Const Scotiabank = "03"
'**************
Global gclsOracle As New clsOracle
Global objUsuario As New clsUsuario
Global objVenta As New clsVenta
Global xProductoRegaloBK As New XArrayDB ' Variable que se devuelve como propiedad los producto
Global objLiquidacion As New clsLiquidacion

Global objMedico As New clsMedico

Global gintDec As Double
Global gintDecTot As Double
'-----------------------------------------------------'
Global gosesVentas As OraSession
Global godbVentas As OraDatabase
'''''''''''''-----------------------------------------------------'
Global gstrAplicacion$, gstrVersion$
Global gvarUSUARIO$, gvarPASSWORD$, gvarTNSNAME$
'ECASTILLO 19.03.2020
'---------------------------------------
Global gvarUSUARIO2$, gvarPASSWORD2$, gvarTNSNAME2$
Global gosesVentas2 As OraSession
Global godbVentas2 As OraDatabase
'---------------------------------------
Global gxdbNC As New XArrayDB
Global OSVERSIONINFO As OSVERSIONINFO

Global Ver As OSVERSIONINFO

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const VER_PRINCIPAL_WIN98 = 4
Public Const VER_PRINCIPAL_WINXP = 5

Public gstrCodAreaUsuario As String

Public gstrCodUsuario As String
Public gstrPassword As String

Public gstrDesAreaUsuario As String

Public gvarHabilitadoContingencia As Variant

 
Public gstrValidaCotizacion As String
Global gstrCodTarjetaFid As String
Global gstrCodTarjetaMon As String
Public gintFidelizado As Integer
Public gstrUriFolderImagen As String
Public gstrPathLog As String


Public gstrVarURL1 As String
Public gstrVarURL2 As String
Public gstrVarURL3 As String
Public gstrVarKEY1 As String
Public gstrFlagLogFile1 As String 'web service
Public gstrFlagLogFile2 As String 'oracle
Public gstrFlagLogFile3 As String 'gmaps
Public gstrFlagLogBD1 As String
Public gstrFlagLogBD2 As String
Public gstrFlagLogBD3 As String
Public gstrFlagSufDir As String
Public gstrFlagValRut As String
Public gstrFlagReclamo As String
Public gstrFlagReservaCap As String
Public gstrIndRAv3 As String
Public gstrIndRAv4 As String
Public gstrIndDCSAP As String
Public gstrIndCreaLogError As String

Public garrCallGoogleMaps As New XArrayDB
Public gintFilesLog As Integer
Public gstrSpecialChrsToWS0 As String
Public gstrSpecialChrsToWS1 As String



''''''''''''
''''''''''''Global gstrCodBtl$
''''''''''''Global gstrCodUsuario$
''''''''''''Global gstrNomUsuario$
''''''''''''Global gstrMaquina$
''''''''''''Global gstrNomMaquina$
''''''''''''Global gstrIgv$
''''''''''''Global gstrTC$
''''''''''''Global gstrEmpresa$
''''''''''''
''''''''''''
''''''''''''Global gstrIDKeito As String
''''''''''''Global gstrFPago As String
'''''''''''''---------------------------------'
'''''''''''''-- gstrFlgRM = ' ' => Magistral Inactivo
'''''''''''''-- gstrFlgRM = '1' => Magistral Activo
'Global gstrFlgRM As String
'''''''''''''---------------------------------'
'Global gvarValores As Variant
'Global gvarIO As Variant
'Global gvarError As Variant
'Global gvarTipo As Variant
'Global gvarNroElem As Variant
'Global gstrMensaje As Variant


Sub psub_Grilla_Traslate(ByRef rgrd As Object, ByVal vvarColumn As Variant, ByVal vvarValue As Variant, ByVal vvarDisplayValue As Variant)
Dim ValueItem As New TrueDBGrid70.ValueItem
'On Error GoTo Error
    'rgrd.Columns(vvarColumn).ValueItems.Clear
    
    rgrd.Columns(vvarColumn).ValueItems.Translate = True
    ValueItem.DisplayValue = vvarDisplayValue
    ValueItem.Value = vvarValue
    rgrd.Columns(vvarColumn).ValueItems.Add ValueItem
    Set ValueItem = Nothing
'Error:
 '   On Error GoTo 0

End Sub

'Valida las fechas del MaskEdit'
Public Function fbln_Valida_Fecha(ByVal vstrFormato$, ByRef rstrError$, Optional MaxDate, Optional MinDate, Optional ValVacio) As Boolean
Dim obj As MaskEdBox
    fbln_Valida_Fecha = False
    Set obj = Screen.ActiveControl
    If IsMissing(ValVacio) Then ValVacio = False
    If ValVacio And val(obj.Text) = 0 Then fbln_Valida_Fecha = True: Exit Function
    If Not IsDate(Format(obj.Text, vstrFormato)) Then
        rstrError = "Valor de Fecha No Válida"
        Exit Function
    End If
    If Not IsMissing(MaxDate) Then
        If CDate(Format(obj.Text, vstrFormato)) > CDate(Format(MaxDate, vstrFormato)) Then
            rstrError = "Fecha Máxima << " & Format(MaxDate, vstrFormato) & " >>"
            Exit Function
        End If
    End If
    If Not IsMissing(MinDate) Then
        If CDate(Format(obj.Text, vstrFormato)) < CDate(Format(MinDate, vstrFormato)) Then
            rstrError = "Fecha Mínima << " & Format(MinDate, vstrFormato) & " >>"
            Exit Function
        End If
    End If
    fbln_Valida_Fecha = True
End Function







