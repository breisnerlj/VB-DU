Attribute VB_Name = "mdlPrincial"

Option Explicit
Global gvarUSUARIO1$, gvarPASSWORD1$, gvarTNSNAME1$
Global gclsOracle1 As New clsOracle
Global gosesVentas1 As OraSession
Global godbVentas1 As OraDatabase

Function conectaOracle( _
            ByVal a_ccod_local_in As String, _
            ByVal A_cnumpedido_in As String, _
            ByVal a_ctelefono_in As String, _
            ByVal A_cmonto_in As String, _
            ByVal A_cusu_crea_in As String, _
            ByVal a_cterminal_in As String, _
            ByVal A_ctipo_rcd_in As String _
) As String
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrorConexion
    Dim vstrTNSNAME  As String
    vstrTNSNAME = "BTLRAC"
    Dim strSQL$
    If gclsOracle1.Conexion(vstrTNSNAME, A_cusu_crea_in, A_cusu_crea_in) <> "" Then Exit Function
    
    
    gclsOracle1.Execute "BEGIN DBMS_APPLICATION_INFO.SET_MODULE('Recarga Telefonica Directa','" & App.Major & "." & App.Minor & "." & App.Revision & "'); END ;"
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim GrabarVentaFallida As String
                

    arrValores = Array(a_ccod_local_in, A_cnumpedido_in, a_ctelefono_in, A_cmonto_in, A_cusu_crea_in, a_cterminal_in)
    'arrDireccion = Array(entrada, entrada, entrada, entrada, entrada, entrada)
    conectaOracle = gclsOracle1.FN_Valor("BTLPROD.fn_GRABA_RECARGA", a_ccod_local_in, A_cnumpedido_in, a_ctelefono_in, A_cmonto_in, A_cusu_crea_in, a_cterminal_in, A_ctipo_rcd_in)
    'GrabarVentaFallida = gclsOracle.SP("BTLPROD.btlprod.fn_GRABA_RECARGA", arrValores, arrDireccion)
        gclsOracle1.cerrar
    Screen.MousePointer = vbDefault

    Exit Function
    'If GrabarVentaFallida = "" Then
     '   conectaOracle = GrabarVentaFallida
    'Exit Function
    'End If
    On Error GoTo 0
    GoTo Final
ErrorConexion:
    Screen.MousePointer = vbDefault
    MsgBox "Existe un Problema con la Conexión" & Chr(13) & "Error :" & Err.Description, vbCritical, App.ProductName
    Exit Function
Final:
    conectaOracle = True
    Screen.MousePointer = vbDefault
    
End Function


Function Respuesta(ByVal a_ccod_local_in As String, _
                   ByVal C_TRACNAVSAT As String _
        ) As String
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrorConexion
    Dim vstrTNSNAME  As String
    vstrTNSNAME = "BTLRAC"
    Dim strSQL$
    If gclsOracle1.Conexion(vstrTNSNAME, "BTL" & a_ccod_local_in, "BTL" & a_ccod_local_in) <> "" Then Exit Function
    
    
    gclsOracle1.Execute "BEGIN DBMS_APPLICATION_INFO.SET_MODULE('Recarga Telefonica Directa','" & App.Major & "." & App.Minor & "." & App.Revision & "'); END ;"
    
    Respuesta = "" & gclsOracle1.FN_Valor("BTLPROD.fn_respuesta_recarga", a_ccod_local_in, C_TRACNAVSAT)
    gclsOracle1.cerrar
    Screen.MousePointer = vbDefault
    
    On Error GoTo 0
    GoTo Final
ErrorConexion:
    Screen.MousePointer = vbDefault
    MsgBox "Existe un Problema con la Conexión" & Chr(13) & "Error :" & Err.Description, vbCritical, App.ProductName
    Exit Function
Final:
    'Respuesta = ""
    Screen.MousePointer = vbDefault
    
End Function

'Public Function fncPalote(strTexto$, Optional intFlag% = 0, Optional strCar$ = "|") As String
'Dim intA
'    strTexto = CStr(strTexto)
'    intA = InStr(strTexto, strCar)
'    If intA = 0 Then fncPalote = Trim(strTexto): Exit Function
'    If intFlag = 0 Then fncPalote = Trim(left(strTexto, intA - 1))
'    If intFlag = 1 Then fncPalote = Trim(right(strTexto, Len(strTexto) - intA))
'End Function

