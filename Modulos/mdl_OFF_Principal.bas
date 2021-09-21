Attribute VB_Name = "mdl_OFF_Principal"
Option Explicit
Global gstrConexion As String
Global strUsuariosXML As String
Global strPreciosXML As String
Global strDetalleVentaXML As String
Global strPagoVentaXML As String
Global gstrIni As String
 
Sub OFF_Main()

    Ver.dwOSVersionInfoSize = Len(Ver)
    
    GetVersionEx Ver

'    gstrConexion = " Provider = Microsoft.Jet.OLEDB.4.0;" _
'                 & " Data Source=" & App.Path & "; " _
'                 & " Extended Properties=""Text;HDR=YES;FMT=Delimited(;)"" "

'    strUsuariosXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "usuarios.xml"
'    strPreciosXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "precios.xml"
'    strDetalleVentaXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "detalleventa.xml"
'    strPagoVentaXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "pagoventa.xml"

'    gstrConexion = "Provider=MSPersist;"

    If frm_OFF_CheckDate.CheckDate = vbNo Then End
    
    frm_OFF_Logeo.Caption = frm_OFF_Logeo.Caption & " " & gstrAplicacion & " * Ver: " & gstrVersion & " - " & gvarTNSNAME
    frm_OFF_Logeo.Show
End Sub
'''
''''Autor : Arturo Escate Espichan
''''Fecha : 02/05/2008
'''Public Sub setteaFormulario(Formulario As Form)
'''With frm_OFF_Principal
'''    Formulario.Width = 7250
'''    Formulario.Height = 7625
'''    Formulario.left = .left + .Picture2.left + 200
'''    Formulario.top = .top + .Picture2.top + 400
'''End With
'''End Sub
'''


Public Sub KillProcess(ByVal processName As String)
On Error GoTo ErrHandler
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename

    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")

    For Each oService In oServices
        servicename = _
            LCase(Trim(CStr(oService.name) & ""))

        If InStr(1, servicename, _
            LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next

    Set oServices = Nothing
    Set oWMI = Nothing
    Exit Sub
ErrHandler:
    Err.Clear
End Sub



