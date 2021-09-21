VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_OFF_Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private objIni As New cls_ArchivoIni
Private xDocumentos As New XArrayDB
Private xFormato As New XArrayDB
Private xImprimeDetalle As New XArrayDB
Private xImprimePago As New XArrayDB
Private xConsultaDoc As New XArrayDB
Private xConsultaDocumento As New XArrayDB
Private xTipoDocumento As New XArrayDB

Public Property Get TipoDocumento() As XArrayDB
    Set TipoDocumento = xTipoDocumento
End Property


Function ListaTipoDocumento(Optional ByVal FlgExcluirSinCorrelativo As Integer = 0) As XArrayDB
    Set ListaTipoDocumento = BuscaDatos(FlgExcluirSinCorrelativo)
End Function

Private Sub Class_Initialize()
    xTipoDocumento.ReDim 0, -1, 0, 4
    xImprimeDetalle.ReDim 0, -1, 0, 25
    xImprimePago.ReDim 0, -1, 0, 10
    xConsultaDoc.ReDim 0, -1, 0, 6
    xConsultaDocumento.ReDim 0, -1, 0, 31
    xFormato.ReDim 0, -1, 0, 3
    BuscaDatosFormato
    BuscaDatos
End Sub

Private Sub Class_Terminate()
    Set objIni = Nothing
End Sub



Function UltimoCorrelativo(ByVal TipoDocumento As String) As String
    Dim xTemp As New XArrayDB
    Dim a As Long
    a = BuscaDatos.Find(0, 0, TipoDocumento)
    UltimoCorrelativo = xDocumentos(a, 2)
End Function

Function ActualizaCorrelativo(ByVal TipoDocumento As String, Optional ByVal NumeroDocumento As String = "", Optional ByVal flgBlanco As Boolean = False) As String
    
    Dim xTemp As New XArrayDB
    Dim a As Long
    Dim NDigitos As Integer
    a = BuscaDatos.Find(0, 0, TipoDocumento)
    
    If Not NumeroDocumento = "" Then
        NDigitos = Len(NumeroDocumento)
        xDocumentos(a, 2) = right(String(NDigitos, "0") & NumeroDocumento, NDigitos)
    Else
        If flgBlanco = False Then
            NDigitos = Len(xDocumentos(a, 2))
            xDocumentos(a, 2) = right(String(NDigitos, "0") & (Val(xDocumentos(a, 2)) + 1), NDigitos)
        Else
            xDocumentos(a, 2) = ""
        End If
    End If
    
    Dim g As Integer
    Dim arrtmp As Variant
    g = 0
    ReDim arrtmp(xDocumentos.Count(1))
    While g < xDocumentos.Count(1)
        arrtmp(g) = xDocumentos(g, 2)
        g = g + 1
    Wend
    'Debug.Print left(Join(arrtmp, ","), Len(Join(arrtmp, ",")) - 1)
    objIni.GuardarIni gstrIni, "general", "NUM_DOCUMENTO", left(Join(arrtmp, ","), Len(Join(arrtmp, ",")) - 1)
End Function

Function NumeroLineas(ByVal TipoDocumento As String) As String
    Dim xTemp As New XArrayDB
    Dim a As Long
    a = BuscaDatos.Find(0, 0, TipoDocumento)
    NumeroLineas = xDocumentos(a, 3)
End Function


Function AnchoDocumento(ByVal TipoDocumento As String) As Integer
    Dim xTemp As New XArrayDB
    Dim a As Long
    a = BuscaDatos.Find(0, 0, TipoDocumento)
    AnchoDocumento = Val(xDocumentos(a, 4))

End Function

'Autor : Juan Arturo Escate Espichan
'Fecha : 08/05/2008
'Proposito : Es para cargar los datos de los documento
'---------------------------------------------------------------------------------------------------------
Private Function BuscaDatos(Optional ByVal FlgExcluirSinCorrelativo As Integer = 0) As XArrayDB
    
    Dim arrTipoCodDocumento As Variant
    Dim arrTipoDesDocumento As Variant
    Dim arrNumDocumento As Variant
    Dim arrLinDocumento As Variant
    Dim arrAnchoDocumento As Variant
    'Dim xArrTmp As New XArrayDB
    
    Dim strTipoCodDocumento As String
    Dim strTipoDesDocumento As String
    Dim strNumDocumento As String
    Dim strLinDocumento As String
    Dim strAnchoDocumento As String
    
    Dim strIniCodDocumento As String
    Dim strIniDesDocumento As String
    Dim strIniNumDocumento As String
    Dim strIniNumLinDocumento As String
    Dim strIniAnchoDocumento As String
    
    Dim strCorrelativo As String
    Dim intPos As Integer
    Dim blnSalir As Boolean
    
    strIniCodDocumento = objIni.LeerIni(gstrIni, "general", "COD_DOCUMENTO", "")
    strIniDesDocumento = objIni.LeerIni(gstrIni, "general", "DES_DOCUMENTO", "")
    strIniNumDocumento = IIf(Replace(objIni.LeerIni(gstrIni, "general", "NUM_DOCUMENTO", ""), ",", "") = "", objIni.LeerIni(gstrIni, "general", "SER_DOCUMENTO", ""), objIni.LeerIni(gstrIni, "general", "NUM_DOCUMENTO", ""))
    strIniNumLinDocumento = objIni.LeerIni(gstrIni, "general", "NUM_LIN_DOCUMENTO", "")
    strIniAnchoDocumento = objIni.LeerIni(gstrIni, "general", "ANCHO_DOCUMENTO", "")
    
    strTipoCodDocumento = ""
    strTipoDesDocumento = ""
    strNumDocumento = ""
    strLinDocumento = ""
    strAnchoDocumento = ""
    
    If FlgExcluirSinCorrelativo <> 0 Then
        
        blnSalir = False
        While Not blnSalir
            
            intPos = InStr(strIniNumDocumento, ",")
            
            If intPos > 0 Then
                
                strCorrelativo = Mid(strIniNumDocumento, 1, intPos - 1)
                
                If Len(Trim(strCorrelativo)) > 0 Then
                    strTipoCodDocumento = strTipoCodDocumento + IIf(Len(strTipoCodDocumento) > 0, ",", "") + Mid(strIniCodDocumento, 1, InStr(strIniCodDocumento, ",") - 1)
                    strTipoDesDocumento = strTipoDesDocumento + IIf(Len(strTipoDesDocumento) > 0, ",", "") + Mid(strIniDesDocumento, 1, InStr(strIniDesDocumento, ",") - 1)
                    strNumDocumento = strNumDocumento + IIf(Len(strNumDocumento) > 0, ",", "") + Mid(strIniNumDocumento, 1, InStr(strIniNumDocumento, ",") - 1)
                    strLinDocumento = strLinDocumento + IIf(Len(strLinDocumento) > 0, ",", "") + Mid(strIniNumLinDocumento, 1, InStr(strIniNumLinDocumento, ",") - 1)
                    strAnchoDocumento = strAnchoDocumento + IIf(Len(strAnchoDocumento) > 0, ",", "") + Mid(strIniAnchoDocumento, 1, InStr(strIniAnchoDocumento, ",") - 1)
                End If
                
                strIniCodDocumento = Mid(strIniCodDocumento, InStr(strIniCodDocumento, ",") + 1)
                strIniDesDocumento = Mid(strIniDesDocumento, InStr(strIniDesDocumento, ",") + 1)
                strIniNumDocumento = Mid(strIniNumDocumento, InStr(strIniNumDocumento, ",") + 1)
                strIniNumLinDocumento = Mid(strIniNumLinDocumento, InStr(strIniNumLinDocumento, ",") + 1)
                strIniAnchoDocumento = Mid(strIniAnchoDocumento, InStr(strIniAnchoDocumento, ",") + 1)
            
            Else
                
                If Len(Trim(strIniNumDocumento)) > 0 Then
                    strTipoCodDocumento = strTipoCodDocumento + IIf(Len(strTipoCodDocumento) > 0, ",", "") + strIniCodDocumento
                    strTipoDesDocumento = strTipoDesDocumento + IIf(Len(strTipoDesDocumento) > 0, ",", "") + strIniDesDocumento
                    strNumDocumento = strNumDocumento + IIf(Len(strNumDocumento) > 0, ",", "") + strIniNumDocumento
                    strLinDocumento = strLinDocumento + IIf(Len(strLinDocumento) > 0, ",", "") + strIniNumLinDocumento
                    strAnchoDocumento = strAnchoDocumento + IIf(Len(strAnchoDocumento) > 0, ",", "") + strIniAnchoDocumento
                End If
                
                blnSalir = True
            
            End If
        
        Wend
        
    Else
    
        strTipoCodDocumento = strIniCodDocumento
        strTipoDesDocumento = strIniDesDocumento
        strNumDocumento = strIniNumDocumento
        strLinDocumento = strIniNumLinDocumento
        strAnchoDocumento = strIniAnchoDocumento
    
    End If
    
'    arrTipoCodDocumento = Split(objIni.LeerIni(gstrIni, "general", "COD_DOCUMENTO", ""), ",")
'    arrTipoDesDocumento = Split(objIni.LeerIni(gstrIni, "general", "DES_DOCUMENTO", ""), ",")
'    arrNumDocumento = Split(objIni.LeerIni(gstrIni, "general", "NUM_DOCUMENTO", ""), ",")
'    arrLinDocumento = Split(objIni.LeerIni(gstrIni, "general", "NUM_LIN_DOCUMENTO", ""), ",")
'    arrAnchoDocumento = Split(objIni.LeerIni(gstrIni, "general", "ANCHO_DOCUMENTO", ""), ",")
    
    arrTipoCodDocumento = Split(strTipoCodDocumento, ",")
    arrTipoDesDocumento = Split(strTipoDesDocumento, ",")
    arrNumDocumento = Split(strNumDocumento, ",")
    arrLinDocumento = Split(strLinDocumento, ",")
    arrAnchoDocumento = Split(strAnchoDocumento, ",")
    
    Set xDocumentos = ArrayaXarray(arrTipoCodDocumento, arrTipoDesDocumento, arrNumDocumento, arrLinDocumento, arrAnchoDocumento)
    
    Set xTipoDocumento = xDocumentos
    Set BuscaDatos = xDocumentos

End Function

Private Sub BuscaDatosFormato()
Dim arrCodFormato As Variant
Dim arrDesFormato As Variant
Dim arrCtdAncho As Variant
Dim arrCtdAlto As Variant


arrCodFormato = Split(objIni.LeerIni(gstrIni, "IMPRESION", "COD_FORMATO", ""), ",")
arrDesFormato = Split(objIni.LeerIni(gstrIni, "IMPRESION", "DES_FORMATO", ""), ",")
arrCtdAncho = Split(objIni.LeerIni(gstrIni, "IMPRESION", "CTD_ANCHO", ""), ",")
arrCtdAlto = Split(objIni.LeerIni(gstrIni, "IMPRESION", "CTD_ALTO", ""), ",")

Set xFormato = ArrayaXarray(arrCodFormato, arrDesFormato, arrCtdAncho, arrCtdAlto)



End Sub



Public Function Formato() As XArrayDB
    Set Formato = xFormato
End Function





Private Sub Boleta(ByVal intNumeroLineas As Integer, ByVal xTmpDet As XArrayDB, ByVal xTmpFP As XArrayDB)

Dim dblTotal  As Double
Dim objArchivoIni As cls_ArchivoIni
Dim j As Integer
Dim i As Integer
'**************
Dim strCodigo As String
Dim strCantidad As String
Dim dblSubTotal As Double
Dim dblImporte As Double
Dim strProducto As String
Dim strDireccion As String
Dim strCodMaquina  As String
Dim strCodUsuarioDependiente As String
Dim strMensaje As String
'**************
Dim p As Printer
Dim intTamañoEspacio  As Integer
'**************
Dim sngCtv As Single
Dim ImporteLetras As String
Dim strSql As String


On Error GoTo ErrorImpresora


    For Each p In Printers
       If UCase(p.Port) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
    
    
    
   If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
        If Not cargaParametros("004") Then
            Exit Sub
        End If

   Else
        Printer.Width = 225 * 56.7
        Printer.Height = (280 / 3) * 56.7
        Printer.Print "."
   End If
    
   
   
   
    dblTotal = 0
    
    If xTmpDet.Count(1) > 0 Then
        j = 0
       Printer.Font.name = "Draft 17cpi"
       Printer.Print
       Printer.Print Space(6) & "" & left(xTmpDet(j, 22) & Space(25), 25) & _
                      Space(1) & left(xTmpDet(j, 20) & Space(15), 15) & xTmpDet(j, 18) & " " & xTmpDet(j, 5)
    
       Printer.Print
       Printer.Print
       strDireccion = xTmpDet(j, 21)
       strCodMaquina = xTmpDet(j, 4)
       strCodUsuarioDependiente = xTmpDet(j, 17)
    
       For i = 0 To xTmpDet.UpperBound(1)
            dblTotal = dblTotal + xTmpDet(i, 16)
            strCodigo = xTmpDet(i, 7)
            strCantidad = IIf(xTmpDet(i, 10) = "1", "F" & str(xTmpDet(i, 9)), str(xTmpDet(i, 9)))
            dblSubTotal = xTmpDet(i, 16)
            dblImporte = xTmpDet(i, 12)
            strProducto = xTmpDet(i, 8)
            Printer.Print left(Trim(strCodigo) & "     ", 5) & " " & _
                         left(strProducto & Space(40), 40) & _
                         right("      " & strCantidad, 5) & " " & _
                         right("        " & Format(dblImporte, "####0.#0"), 8) & " " & _
                         right("        " & Format(dblSubTotal, "####0.#0"), 8)
       Next i
       For j = i To intNumeroLineas
            If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
                If j = intNumeroLineas Then Exit For Else Printer.Print
            Else
                Printer.Print
            End If
       Next j
       
       Printer.Print Trim(objOFFUsuario.Empresa) & Space(11) & "Total Venta " & Trim(objOFFUsuario.SmbMoneda) & right("        " & Format(dblTotal, "####0.#0"), 8)
       

       sngCtv = Val(Replace(dblTotal, " ", ""))
       ImporteLetras = UCase(Letra(Int(sngCtv), 80))
       strSql = right(Format(sngCtv - Int(sngCtv), ".00"), 2)
       Printer.Print "SON : " & Trim(ImporteLetras) & " Y " & strSql & "/100" & " " & objOFFUsuario.DesLgMoneda
        
        
       
       Printer.Print strDireccion
       Printer.Print Space(46) + strCodMaquina + "   " + strCodUsuarioDependiente

       Set objArchivoIni = New cls_ArchivoIni
       strMensaje = objArchivoIni.LeerIni(gstrIni, "GLOSA", "LINE1", "")
       Set objArchivoIni = Nothing
        
        

       intTamañoEspacio = Int((AnchoDocumento(COD_TIPO_BOL) - Len(IIf(IsNull(strMensaje), "", strMensaje))) / 2)
       Printer.Print Space(intTamañoEspacio) + Chr(34) + IIf(IsNull(strMensaje), "", strMensaje) + Chr(34)
       Printer.EndDoc
       
       
       
   Else
        Err.Raise vbObjectError + 513, "cls_OFF_Documento.Boleta", "No se encontró datos para el documento"
   End If
   
   

   Exit Sub
ErrorImpresora:
    Err.Raise Err.Number, "cls_OFF_Documento.Boleta", Err.Description

End Sub


Public Sub Imprime(ByVal intSecVta As Integer)
    Dim cnn As ADODB.Connection
    Dim strSql As String
    Dim rsDetalleVenta As New ADODB.Recordset
    Dim rsPagoVenta As New ADODB.Recordset
    '**************
    Dim strNumDocumento As String
    Dim strTipoDocumento As String

    On Error GoTo CtrlErr

        Set cnn = New ADODB.Connection
        cnn.Open gstrConexion
        
'        strSQL = "select * from pagoventa.txt"
'        rsPagoVenta.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
        rsPagoVenta.CursorLocation = adUseClient
        rsPagoVenta.Open strPagoVentaXML, gstrConexion, adOpenStatic, adLockReadOnly
        
        If Not rsPagoVenta.EOF Then
            rsPagoVenta.Filter = "SEC_VENTA = " & intSecVta
        End If
        
        If rsPagoVenta.RecordCount > 0 Then
            rsPagoVenta.MoveFirst
            LimpiaImprimePago
            While Not rsPagoVenta.EOF
                Call AgregaImprimePago(rsPagoVenta!SEC_VENTA, _
                        rsPagoVenta!NUM_ITEM, _
                        rsPagoVenta!COD_PAGO, _
                        rsPagoVenta!COD_MONEDA, _
                        rsPagoVenta!MTO_IMPORTE, _
                        rsPagoVenta!MTO_SOLES, _
                        rsPagoVenta!MTO_VUELTO, _
                        IIf(IsNull(rsPagoVenta!NUM_TARJETA), "", rsPagoVenta!NUM_TARJETA), _
                        IIf(IsNull(rsPagoVenta!FCH_VENCIMIENTO), "", rsPagoVenta!FCH_VENCIMIENTO), _
                        rsPagoVenta!NUM_CUOTAS, _
                        rsPagoVenta!TIPO_CAMBIO)
                rsPagoVenta.MoveNext
            Wend
        Else
            Err.Raise vbObjectError + 513, "cls_OFF_Documento.Imprime", "No se encontró forma de pago para la secuencia  " & intSecVta
        End If
                
'        strSQL = "select * from detalleventa.txt"
'        rsDetalleVenta.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
        rsDetalleVenta.CursorLocation = adUseClient
        rsDetalleVenta.Open strPagoVentaXML, gstrConexion, adOpenStatic, adLockReadOnly
        rsDetalleVenta.Filter = "SEC_VENTA = " & intSecVta
        
        If rsDetalleVenta.RecordCount > 0 Then
            rsDetalleVenta.MoveFirst
            While Not rsDetalleVenta.EOF
                LimpiaImprimeDetalle
                strNumDocumento = rsDetalleVenta!NUM_DOCUMENTO
                strTipoDocumento = rsDetalleVenta!COD_TIPO_DOCUMENTO
                Do While Not rsDetalleVenta.EOF _
                            And strNumDocumento = rsDetalleVenta!NUM_DOCUMENTO _
                            And strTipoDocumento = rsDetalleVenta!COD_TIPO_DOCUMENTO
                        Call AgregaImprimeDetalle(rsDetalleVenta!SEC_VENTA, rsDetalleVenta!COD_CIA, _
                                    rsDetalleVenta!COD_LOCAL, _
                                    rsDetalleVenta!COD_TIPO_DOCUMENTO, _
                                    rsDetalleVenta!COD_MAQUINA, _
                                    rsDetalleVenta!NUM_DOCUMENTO, _
                                    rsDetalleVenta!NUM_ITEM, _
                                    rsDetalleVenta!COD_PRODUCTO, _
                                    rsDetalleVenta!DES_PRODUCTO, _
                                    rsDetalleVenta!CTD_PRODUCTO, _
                                    rsDetalleVenta!FLG_FRACCION, _
                                    rsDetalleVenta!PCT_DESCUENTO, _
                                    rsDetalleVenta!PRC_UNITARIO, _
                                    rsDetalleVenta!PRC_ORIGINAL, _
                                    rsDetalleVenta!MTO_IGV, _
                                    rsDetalleVenta!MTO_EXONERADO, _
                                    rsDetalleVenta!MTO_SUBTOTAL, _
                                    rsDetalleVenta!USU_EMISION, _
                                    rsDetalleVenta!FCH_EMISION, _
                                    IIf(IsNull(rsDetalleVenta!TIP_CLIENTE), "0", rsDetalleVenta!TIP_CLIENTE), _
                                    IIf(IsNull(rsDetalleVenta!NUM_RUC), "", rsDetalleVenta!NUM_RUC), _
                                    IIf(IsNull(rsDetalleVenta!DIR_CLIENTE), "", rsDetalleVenta!DIR_CLIENTE), _
                                    IIf(IsNull(rsDetalleVenta!NOM_CLIENTE), "", rsDetalleVenta!NOM_CLIENTE), _
                                    rsDetalleVenta!PCT_IGV, _
                                    IIf(IsNull(rsDetalleVenta!PARTIDA_ARANCELARIA), "", rsDetalleVenta!PARTIDA_ARANCELARIA), _
                                    rsDetalleVenta!CTD_FRACCIONA)
                    rsDetalleVenta.MoveNext
                    If rsDetalleVenta.EOF Then Exit Do
                Loop
                
                Call ImprimeDoc(strTipoDocumento, ImprimeDetalle, ImprimePago)
            Wend
        
        Else
            Err.Raise vbObjectError + 513, "cls_OFF_Documento.Imprime", "No se encontró items para la secuencia  " & intSecVta
        End If
        
    rsPagoVenta.Filter = adFilterNone
    rsPagoVenta.Close
    rsDetalleVenta.Filter = adFilterNone
    rsDetalleVenta.Close
    'cnn.Close
    Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.Imprime", Err.Description
End Sub




Public Sub ImprimeDoc(ByVal strTipoDocumento As String, _
                        ByVal xTempDetalle As XArrayDB, _
                        ByVal xTempFormpaPago As XArrayDB)
                        
Dim intNumeroLineas As Integer
On Error GoTo CtrlErr
                            
   intNumeroLineas = Val(NumeroLineas(strTipoDocumento))
                                    
        Select Case strTipoDocumento
            Case COD_TIPO_BOL
                Call Boleta(intNumeroLineas, xTempDetalle, xTempFormpaPago)
            Case COD_TIPO_FAC
                Call Factura(intNumeroLineas, xTempDetalle, xTempFormpaPago)
            Case COD_TIPO_TKB
                Call Ticket(intNumeroLineas, xTempDetalle, xTempFormpaPago)
                'Call TicketNew(intNumeroLineas, xTempDetalle, xTempFormpaPago)
            Case COD_TIPO_TKF
                Call Ticket(intNumeroLineas, xTempDetalle, xTempFormpaPago)
                'Call TicketNew(intNumeroLineas, xTempDetalle, xTempFormpaPago)
        End Select
Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.ImprimeDoc", Err.Description

End Sub


Public Sub ImprimePorDocumento(ByVal pstrTipoDocumento As String, ByVal pstrNumDocumento As String)
    Dim intNumeroLineas As Integer
'    Dim cnn As ADODB.Connection
    Dim strSql As String
    Dim rsDetalleVenta As New ADODB.Recordset
    Dim rsPagoVenta As New ADODB.Recordset
    Dim intSecuencia As Integer

    On Error GoTo CtrlErr
   
    intNumeroLineas = Val(NumeroLineas(pstrTipoDocumento))
   
'    Set cnn = New ADODB.Connection
'    cnn.Open gstrConexion
   
   
'    strSQL = "select * from detalleventa.txt"
'    rsDetalleVenta.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
    rsDetalleVenta.CursorLocation = adUseClient
    rsDetalleVenta.Open strDetalleVentaXML, gstrConexion, adOpenDynamic, adLockOptimistic
    rsDetalleVenta.Filter = "COD_TIPO_DOCUMENTO = '" & pstrTipoDocumento & "' AND NUM_DOCUMENTO = '" & pstrNumDocumento & "' AND TIP_MOVIMIENTO = '" & COD_TIP_MOV_VENTA & "'"
    
    If rsDetalleVenta.RecordCount > 0 Then
        intSecuencia = rsDetalleVenta!SEC_VENTA
'        strSQL = "select * from pagoventa.txt"
'        rsPagoVenta.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
        rsPagoVenta.CursorLocation = adUseClient
        rsPagoVenta.Open strPagoVentaXML, gstrConexion, adOpenDynamic, adLockOptimistic
        rsPagoVenta.Filter = "SEC_VENTA = " & intSecuencia
        
        If rsPagoVenta.RecordCount > 0 Then
            rsPagoVenta.MoveFirst
            LimpiaImprimePago
            While Not rsPagoVenta.EOF
                Call AgregaImprimePago(rsPagoVenta!SEC_VENTA, _
                                       rsPagoVenta!NUM_ITEM, _
                                       rsPagoVenta!COD_PAGO, _
                                       rsPagoVenta!COD_MONEDA, _
                                       rsPagoVenta!MTO_IMPORTE, _
                                       rsPagoVenta!MTO_SOLES, _
                                       rsPagoVenta!MTO_VUELTO, _
                                       IIf(IsNull(rsPagoVenta!NUM_TARJETA), "", rsPagoVenta!NUM_TARJETA), _
                                       IIf(IsNull(rsPagoVenta!FCH_VENCIMIENTO), "", rsPagoVenta!FCH_VENCIMIENTO), _
                                       rsPagoVenta!NUM_CUOTAS, _
                                       rsPagoVenta!TIPO_CAMBIO)
                rsPagoVenta.MoveNext
            Wend
        Else
            Err.Raise vbObjectError + 513, "cls_OFF_Documento.ImprimePorDocumento", "No se encontró forma de pago para la secuencia  " & intSecuencia
        End If
            
        rsDetalleVenta.MoveFirst
        LimpiaImprimeDetalle
        While Not rsDetalleVenta.EOF
            Call AgregaImprimeDetalle(rsDetalleVenta!SEC_VENTA, rsDetalleVenta!COD_CIA, _
                                      rsDetalleVenta!COD_LOCAL, _
                                      rsDetalleVenta!COD_TIPO_DOCUMENTO, _
                                      rsDetalleVenta!COD_MAQUINA, _
                                      rsDetalleVenta!NUM_DOCUMENTO, _
                                      rsDetalleVenta!NUM_ITEM, _
                                      rsDetalleVenta!COD_PRODUCTO, _
                                      rsDetalleVenta!DES_PRODUCTO, _
                                      rsDetalleVenta!CTD_PRODUCTO, _
                                      rsDetalleVenta!FLG_FRACCION, _
                                      rsDetalleVenta!PCT_DESCUENTO, _
                                      rsDetalleVenta!PRC_UNITARIO, _
                                      rsDetalleVenta!PRC_ORIGINAL, _
                                      rsDetalleVenta!MTO_IGV, _
                                      rsDetalleVenta!MTO_EXONERADO, _
                                      rsDetalleVenta!MTO_SUBTOTAL, _
                                      rsDetalleVenta!USU_EMISION, _
                                      rsDetalleVenta!FCH_EMISION, _
                                      IIf(IsNull(rsDetalleVenta!TIP_CLIENTE), "0", rsDetalleVenta!TIP_CLIENTE), _
                                      IIf(IsNull(rsDetalleVenta!NUM_RUC), "", rsDetalleVenta!NUM_RUC), _
                                      IIf(IsNull(rsDetalleVenta!DIR_CLIENTE), "", rsDetalleVenta!DIR_CLIENTE), _
                                      IIf(IsNull(rsDetalleVenta!NOM_CLIENTE), "", rsDetalleVenta!NOM_CLIENTE), _
                                      rsDetalleVenta!PCT_IGV, _
                                      IIf(IsNull(rsDetalleVenta!PARTIDA_ARANCELARIA), "", rsDetalleVenta!PARTIDA_ARANCELARIA), _
                                      rsDetalleVenta!CTD_FRACCIONA)
            rsDetalleVenta.MoveNext
        Wend
    Else
        Err.Raise vbObjectError + 513, "cls_OFF_Documento.Imprime", "No se encontró items para la secuencia  " & intSecuencia
    End If
    
    rsPagoVenta.Filter = adFilterNone
    rsPagoVenta.Close
    rsDetalleVenta.Filter = adFilterNone
    rsDetalleVenta.Close
'    cnn.Close
   
    Select Case pstrTipoDocumento
        Case COD_TIPO_BOL
            Call Boleta(intNumeroLineas, ImprimeDetalle, ImprimePago)
            Case COD_TIPO_FAC
                Call Factura(intNumeroLineas, ImprimeDetalle, ImprimePago)
            Case COD_TIPO_TKB
                Call Ticket(intNumeroLineas, ImprimeDetalle, ImprimePago)
            Case COD_TIPO_TKF
                Call Ticket(intNumeroLineas, ImprimeDetalle, ImprimePago)
    End Select
    Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.ImprimeDoc", Err.Description
End Sub

Private Sub Factura(ByVal intNumeroLineas As Integer, ByVal xTmpDet As XArrayDB, ByVal xTmpFP As XArrayDB)

    Dim dblPctIgv As Double
    Dim i As Integer
    Dim j As Integer

    Dim arrMeses(12) As String
    Dim dblDonacion As Double
    Dim strLetras As String
    Dim dblMtoBaseImp As Double
    Dim dblMtoExonerado As Double
    Dim dblMtoImpuesto As Double
    Dim dblMtoTotal As Double
    Dim strMensaje As String
    Dim intTamañoEspacio  As Integer
    Dim intFormato  As Integer
    Dim p As Printer
    Dim strGlosaConv As String
    Dim dblMtoTotalRef As Double
    Dim strDesAuxCliNombre As String
    Dim strCodMaquina As String
    Dim strCodUsuarioDependiente As String
    Dim strCodConvenio As String
    Dim strCodigo As String
    Dim strCantidad As String
    Dim strProducto As String
    Dim dblImporte As Double
    Dim dblSubTotal As Double
    Dim dblTotal As Double
    Dim dblPctDscto As Double
    '**************
    Dim sngCtv As Single
    Dim ImporteLetras As String
    Dim strSql As String
    Dim objArchivoIni As cls_ArchivoIni

    On Error GoTo CtrlErr

    arrMeses(1) = "Enero"
    arrMeses(2) = "Febrero"
    arrMeses(3) = "Marzo"
    arrMeses(4) = "Abril"
    arrMeses(5) = "Mayo"
    arrMeses(6) = "Junio"
    arrMeses(7) = "Julio"
    arrMeses(8) = "Agosto"
    arrMeses(9) = "Setiembre"
    arrMeses(10) = "Octubre"
    arrMeses(11) = "Noviembre"
    arrMeses(12) = "Diciembre"

    For Each p In Printers
       If UCase(p.Port) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
   
   If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
        If Not cargaParametros("002") Then
            Exit Sub
        End If

   Else
        Printer.Width = 225 * 56.7
        Printer.Height = 153 * 56.7
   End If
   
   
   
   
   If xTmpDet.Count(1) > 0 Then
        j = 0
        i = 0
        Printer.Print: Printer.Print
        Printer.FontName = "Draft 17cpi"
        Printer.Print "."
        Printer.FontName = "Draft 10cpi"
        Printer.Print Space(55) & "NIT " & xTmpDet(j, 20)
        Printer.FontName = "Draft 12cpi"
        
        
        dblPctIgv = xTmpDet(j, 23) / 100
        
        
        
        
        dblMtoBaseImp = 0
        dblMtoExonerado = 0
        dblMtoImpuesto = 0
        dblMtoTotal = 0
        
        strCodMaquina = xTmpDet(j, 4)
        strCodUsuarioDependiente = xTmpDet(j, 17)
        
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print Space(18) & strCodMaquina & "   " & strCodUsuarioDependiente & "   " & objOFFUsuario.CodigoUsuario
        Printer.Print Space(8) & CStr(Day(xTmpDet(j, 18))) & " de " & arrMeses(Month(xTmpDet(j, 18))) & " de " & CStr(Year(xTmpDet(j, 18)))
        Printer.Print Space(6) & left(xTmpDet(j, 22) & Space(60), 60) & " " & xTmpDet(j, 5)
        Printer.Print Space(8) & xTmpDet(j, 20)
        Printer.Print Space(8) & xTmpDet(j, 21)

        Printer.Print
        Printer.Print
        Printer.Print
        
        Printer.Print
        Printer.FontName = "Draft 17cpi"
        
        For i = 0 To xTmpDet.UpperBound(1)
        
            dblMtoExonerado = dblMtoExonerado + xTmpDet(i, 15)
            dblMtoImpuesto = dblMtoImpuesto + xTmpDet(i, 14)
            dblMtoTotal = dblMtoTotal + xTmpDet(i, 16)
        
        
            strCodigo = xTmpDet(i, 7)
            strCantidad = IIf(xTmpDet(i, 10) = "1", "F" & str(xTmpDet(i, 9)), str(xTmpDet(i, 9)))
            strProducto = xTmpDet(i, 8)
            dblImporte = xTmpDet(i, 12)
            dblSubTotal = Round(xTmpDet(i, 12) / xTmpDet(i, 25), 2)
            dblTotal = xTmpDet(i, 16)
            dblPctDscto = xTmpDet(i, 11)
        
        
        
        
        
            Printer.Print right(Space(7) & Trim(strCantidad), 7) & Space(2) & left(strCodigo & " " & strProducto & Space(40), 40) & _
                        Space(2) & right(Space(8) & Format(dblImporte, "#####0.#0"), 8) & _
                       Space(2) & right(Space(8) & Format(dblSubTotal, "#####0.#0"), 8) & _
                       Space(2) & right(Space(6) & Format(dblPctDscto, "#####0.#0"), 6) & _
                       Space(2) & right(Space(8) & Format(dblTotal, "#####0.#0"), 8)



            If dblPctIgv = 0 Then
               If xTmpDet(i, 24) <> "" Then
                  Printer.Print "        P.A. : " & xTmpDet(i, 24)
                  i = i + 1
               End If
            End If


        Next i
        
        dblMtoBaseImp = dblMtoTotal - dblMtoImpuesto
        
        
        Printer.FontName = "Draft 12cpi"
        For j = 1 To intNumeroLineas - i
           Printer.Print
        Next j

        Printer.Print
        Printer.Print
        Printer.Print


       sngCtv = Val(Replace(dblMtoTotal, " ", ""))
       ImporteLetras = UCase(Letra(Int(sngCtv), 80))
       strSql = right(Format(sngCtv - Int(sngCtv), ".00"), 2)
       

       Printer.Print "SON : " & Trim(ImporteLetras) & " Y " & strSql & "/100" & " " & objOFFUsuario.DesLgMoneda




    Printer.Print

    Printer.Print Space(40) & "%  " & right(Space(11) & Format(dblPctIgv * 100, "##0.00"), 11)
    Printer.Print Space(8) & Trim(objOFFUsuario.SmbMoneda) & right(Space(11) & _
                    Format(IIf(dblMtoBaseImp > 0, dblMtoBaseImp, dblMtoExonerado), "#### ##0.#0"), 11) & Space(18) & _
                    Trim(objOFFUsuario.SmbMoneda) & right(Space(11) & Format(dblMtoImpuesto, "#### ##0.#0"), 11) & Space(25) & _
                    Trim(objOFFUsuario.SmbMoneda) & right(Space(11) & Format(dblMtoTotal, "#### ##0.#0"), 11)

    Printer.Print
    Printer.Print Space(56) + strCodMaquina + "   " + strCodUsuarioDependiente
    
    ''AQUI VA EL MENSAJE


       Set objArchivoIni = New cls_ArchivoIni
       strMensaje = objArchivoIni.LeerIni(gstrIni, "GLOSA", "LINE1", "")
       Set objArchivoIni = Nothing


    intTamañoEspacio = Int((AnchoDocumento(COD_TIPO_FAC) - Len(IIf(IsNull(strMensaje), "", strMensaje))) / 2)
    Printer.Print Space(intTamañoEspacio) + Chr(34) + IIf(IsNull(strMensaje), "", strMensaje) + Chr(34)
    Printer.EndDoc
   
   Else
        Err.Raise vbObjectError + 513, "cls_OFF_Documento.Factura", "No se encontraron datos del documento"
   End If



Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.Factura", Err.Description


End Sub



Private Sub Ticket(ByVal intNumeroLineas As Integer, ByVal xTmpDet As XArrayDB, ByVal xTmpFP As XArrayDB)

Dim dblDonacion As Double
Dim strGlosaConv As String
Dim strMensaje As String
Dim dblTotal As Double
Dim i As Integer
Dim j As Integer
'''''''''''''''''''''''''''''''''''''''''''''''
Dim p As Printer
Dim intRespuesta As Integer
Dim dblMtoTotalRef As Double
Dim intTotItems As Integer
Dim vInd As Integer
'**************
Dim TipoDocumento As String
'**************
Dim objFormaPago As cls_OFF_FormaPago
Dim objArchivoIni As cls_ArchivoIni
Dim strDesFormaPago As String
Dim intPctIgv As Integer
Dim dblMtoExonerado As Double
Dim dblMtoImpuesto As Double
Dim dblMtoTotal As Double
Dim dblMtoBaseImp As Double

Dim strCodigo As String
Dim strCantidad As String
Dim dblSubTotal As Double
Dim dblImporte As Double
Dim strProducto As String



    vInd = 0
    For Each p In Printers
       If UCase(p.Port) = "LPT1:" Then
          Set Printer = p
          vInd = "1"
          Exit For
       End If
    Next p
    
    
    If xTmpDet.Count(1) > 0 Then
    j = 0
    i = 0
    TipoDocumento = xTmpDet(j, 3)
    intPctIgv = xTmpDet(j, 23)
        
       
       
    On Error GoTo 0
        On Error GoTo ErrorImpresora
            On Error GoTo Mal_Seteo
            GoTo OK_Seteo
Mal_Seteo:
            MsgBox Err.Description, vbExclamation, "Error de Ticketera.....????"
            Exit Sub
OK_Seteo:

            '***********************************************************************'
            Open "lpt1" For Output As #1
            '''Open "COM3" For Output As #1
            '**** 16-DOT PITCH FONT *******'
            Print #1, Chr$(27); Chr$(58);
            '-------------------------------'
            Print #1, "    BOTICAS BTL    "
            '****** Standar ********'
            Print #1, Chr$(27); Chr$(77);
            '-----------------------------'
            Print #1, "   BOTICA TORRES DE LIMATAMBO SAC    "
            Print #1, Space(5) & objOFFUsuario.DireccionEmpresa
            Print #1, Space(8) & "RUC :" & objOFFUsuario.RucEmpresa
            If TipoDocumento = "TKB" Then
             Print #1, left(xTmpDet(j, 22) & Space(25), 25)
            End If
            Print #1, "BTL" & objOFFUsuario.CodLocal & Space(2) & CadenaDirecc_Part1(objOFFUsuario.DireccionLocal)
            Print #1, Space(8) & CadenaDirecc_Part2(objOFFUsuario.DireccionLocal)
            Print #1, "FECHA:" & xTmpDet(j, 18) & Space(1) & "CAJA:" & " " & xTmpDet(j, 4)
            Print #1, "Ticket" & " " & "Nº" & Space(1) & xTmpDet(j, 5) & " " & "Serie:" & " " & objOFFUsuario.CodSerieEtiq
            Print #1, ""
            Print #1, left("Código" & Space(5), 5) & Space(1) & _
                      left("Descripción" & Space(20), 20) & Space(1) & _
                      right(Space(5) & "Cant.", 5) & Space(1) & _
                      right(Space(8) & "Importe", 8) & Space(1)
            Print #1, String(42, "-")
            Printer.FontBold = False
            
            
            dblMtoExonerado = 0
            dblMtoImpuesto = 0
            dblMtoTotal = 0
            
            For i = 0 To xTmpDet.UpperBound(1)
                dblMtoExonerado = dblMtoExonerado + xTmpDet(i, 15)
                dblMtoImpuesto = dblMtoImpuesto + xTmpDet(i, 14)
                dblMtoTotal = dblMtoTotal + xTmpDet(i, 16)
                strCodigo = xTmpDet(i, 7)
                strCantidad = IIf(xTmpDet(i, 10) = "1", "F" & str(xTmpDet(i, 9)), str(xTmpDet(i, 9)))
                dblSubTotal = xTmpDet(i, 16)
                dblImporte = xTmpDet(i, 12)
                strProducto = xTmpDet(i, 8)
            
               If Len(strProducto) > 20 Then ' SI LA DESCRIPCION DE PRODUCTO ES MAYOR A 20 CARACTERES'
                        Print #1, left(strCodigo & Space(5), 5) & Space(1) & _
                                         left(Mid(strProducto, 1, 20) & Space(20), 20) & Space(1) & _
                                         right(Space(5) & strCantidad, 5) & Space(2) & _
                                         right(Space(7) & dblSubTotal, 7) & Space(1)
                        Print #1, Space(5) & left(Mid(strProducto, 21, 40) & Space(20), 20) & Space(3)

                  Else
                         Print #1, left(strCodigo & Space(5), 5) & Space(1) & _
                                   left(Mid(strProducto, 1, 30) & Space(30), 30) & Space(2) & _
                                   Space(30) & right(Space(5) & strCantidad, 5) & Space(1) & _
                                   right(Space(7) & dblSubTotal, 7) & Space(1)
                End If
            Next i
            
            intTotItems = xTmpDet.Count(1)
            Print #1, ""
            Print #1, Space(5) & "Total Productos" & " " & intTotItems
            Print #1, ""
          
            
            
            
            If TipoDocumento = "TKF" Then
               dblMtoBaseImp = dblMtoTotal - dblMtoImpuesto
               If dblMtoExonerado = 0 Then
                  Print #1, Space(15) & "Valor venta   S/. " & " " & Format(dblMtoBaseImp, "#0.00")
               Else
                  Print #1, Space(15) & "Valor venta   S/. " & " " & Format(dblMtoExonerado, "#0.00")
               End If
               Print #1, Space(15) & "IGV " & intPctIgv & "%.      S/. " & " " & Format(dblMtoImpuesto, "#0.00")
            End If
                  Print #1, Space(15) & "Total         S/. " & " " & Format(dblMtoTotal, "#0.00")
                  Print #1, ""
                  
                  
                  Dim dblMonNac As Double
                  Dim dblVuelto As Double
                  
                  dblMonNac = 0: dblVuelto = 0
                  
                  If xTmpFP.Count(1) > 0 Then
                    For i = 0 To xTmpFP.UpperBound(1)
                        
                        Set objFormaPago = New cls_OFF_FormaPago
                        strDesFormaPago = objFormaPago.DescripcionFormaPago(xTmpFP(i, 2))
                        Set objFormaPago = Nothing
                        
                        Print #1, left(strDesFormaPago & Space(25), 25) & Space(4) & _
                                  "S/." & " " & right(Space(8) & Format(xTmpFP(i, 5), "#0.00"), 8)
                        
                        dblMonNac = dblMonNac + xTmpFP(i, 5)
                        dblVuelto = dblVuelto + xTmpFP(i, 6)
                    Next i
                  Else
                    Err.Raise vbObjectError + 513, "cls_OFF_Documento.Ticket", "No se encontró datos para la forma de pago"
                  End If
                  
                  
                                      
                  Print #1, "Vuelto                       S/.   " & " " & Format(dblVuelto, "#0.00")
                                      
            If TipoDocumento = "TKF" Then
               Print #1, "Ruc Ciente   :" & " " & xTmpDet(j, 20)
               Print #1, "Razón Social :" & " " & xTmpDet(j, 22)
            End If
               Print #1, "Cajero:" & Space(1) & objOFFUsuario.DesUsuarioVenta(xTmpDet(j, 17))
               Print #1, String(42, "-")
               Print #1, "************* www.btl.com.pe *************" & _
                         "           Gracias por su compra          "
               Print #1, "       DELIVERY 612-5000 LAS 24 HORAS     "
             
              
       Set objArchivoIni = New cls_ArchivoIni
       strMensaje = objArchivoIni.LeerIni(gstrIni, "GLOSA", "LINE1", "")
       Set objArchivoIni = Nothing
              
            ''AQUI VA EL MENSAJE
            If Len(strMensaje) < 40 Then
                    Print #1, strMensaje
              Else
                For i = 0 To Val(Len(strMensaje)) / 40
                    Print #1, strMensaje
                Next i
            End If
            
            For j = 1 To 5
               Print #1, ""
            Next j
      
            Print #1, Chr(27) + "d" + "1"

            '*********************************************************************************************************
            Close #1
            
    Else
        Err.Raise vbObjectError + 513, "cls_OFF_Documento.Ticket", "No se encontró datos para el documento"
    End If
                  
    On Error GoTo 0
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
          Resume
       Else
          Exit Sub
     End If




End Sub

Public Sub AgregaImprimeDetalle(pintSecuenciaVenta As Integer, pstrCodCia As String, _
                                pstrCodLocal As String, _
                                pstrCodTipoDocumento As String, _
                                pstrCodMaquina As String, _
                                pstrNumDocumento As String, _
                                pintNumItem As Integer, _
                                pstrCodProducto As String, _
                                pstrDesProducto As String, _
                                pintCtdProducto As Integer, _
                                pstrFlgFraccion As String, _
                                pdblPctDescuento As Double, _
                                pdblPrcUnitario As Double, _
                                pdblPrcOriginal As Double, _
                                pdblMtoIGV As Double, _
                                pdblMtoExonerado As Double, _
                                pdblMtoSubTotal As Double, _
                                pstrUsuEmision As String, _
                                pstrFechaEmision As String, _
                                pstrTipoCliente As String, _
                                pstrNumRuc As String, _
                                pstrDirCliente As String, _
                                pstrNomCliente As String, _
                                pdblPctIGV As Double, _
                                pstrPartArancel As String, _
                                pintCtdFraccion As Integer)
    Dim ultimo As Integer
    Dim i As Integer
    Dim encontro As Boolean
    
    If xImprimeDetalle.Count(1) < 0 Then Exit Sub
    
    
    

    encontro = False
    ultimo = xImprimeDetalle.Count(1)

    If encontro = False Then
        xImprimeDetalle.AppendRows
    End If


    If xImprimeDetalle.Count(1) = 0 Then ultimo = 0: xImprimeDetalle.AppendRows
    
    
    xImprimeDetalle(ultimo, 0) = pintSecuenciaVenta
    xImprimeDetalle(ultimo, 1) = pstrCodCia
    xImprimeDetalle(ultimo, 2) = pstrCodLocal
    xImprimeDetalle(ultimo, 3) = pstrCodTipoDocumento
    xImprimeDetalle(ultimo, 4) = pstrCodMaquina
    xImprimeDetalle(ultimo, 5) = pstrNumDocumento
    xImprimeDetalle(ultimo, 6) = pintNumItem
    xImprimeDetalle(ultimo, 7) = pstrCodProducto
    xImprimeDetalle(ultimo, 8) = pstrDesProducto
    xImprimeDetalle(ultimo, 9) = pintCtdProducto
    xImprimeDetalle(ultimo, 10) = pstrFlgFraccion
    xImprimeDetalle(ultimo, 11) = pdblPctDescuento
    xImprimeDetalle(ultimo, 12) = pdblPrcUnitario
    xImprimeDetalle(ultimo, 13) = pdblPrcOriginal
    xImprimeDetalle(ultimo, 14) = pdblMtoIGV
    xImprimeDetalle(ultimo, 15) = pdblMtoExonerado
    xImprimeDetalle(ultimo, 16) = pdblMtoSubTotal
    xImprimeDetalle(ultimo, 17) = pstrUsuEmision
    xImprimeDetalle(ultimo, 18) = pstrFechaEmision
    xImprimeDetalle(ultimo, 19) = pstrTipoCliente
    xImprimeDetalle(ultimo, 20) = pstrNumRuc
    xImprimeDetalle(ultimo, 21) = pstrDirCliente
    xImprimeDetalle(ultimo, 22) = pstrNomCliente
    xImprimeDetalle(ultimo, 23) = pdblPctIGV
    xImprimeDetalle(ultimo, 24) = pstrPartArancel
    xImprimeDetalle(ultimo, 25) = pintCtdFraccion
    
End Sub


Public Function LimpiaImprimeDetalle()
    xImprimeDetalle.ReDim 0, -1, 0, 25

End Function

Public Function LimpiaImprimePago()
    xImprimePago.ReDim 0, -1, 0, 10

End Function

Public Property Get ImprimePago() As XArrayDB
    Set ImprimePago = xImprimePago
End Property



Public Property Get ImprimeDetalle() As XArrayDB
    Set ImprimeDetalle = xImprimeDetalle
End Property


Public Sub AgregaImprimePago(pintSecuenciaVenta As Integer, _
                                pintNumItem As Integer, _
                                pstrCodPago As String, _
                                pstrCodMoneda As String, _
                                pdblMtoImporte As Double, _
                                pdblMtoSoles As Double, _
                                pdblMtoVuelto As Double, _
                                pstrNumTarjeta As String, _
                                pstrFechaVencimiento As String, _
                                pintNumCuotas As Integer, _
                                pdblTipoCambio As Double)
    Dim ultimo As Integer
    Dim i As Integer
    Dim encontro As Boolean
    
    If xImprimePago.Count(1) < 0 Then Exit Sub
    
    
    

    encontro = False
    ultimo = xImprimePago.Count(1)

    If encontro = False Then
        xImprimePago.AppendRows
    End If


    If xImprimePago.Count(1) = 0 Then ultimo = 0: xImprimePago.AppendRows
    
    
    xImprimePago(ultimo, 0) = pintSecuenciaVenta
    xImprimePago(ultimo, 1) = pintNumItem
    xImprimePago(ultimo, 2) = pstrCodPago
    xImprimePago(ultimo, 3) = pstrCodMoneda
    xImprimePago(ultimo, 4) = pdblMtoImporte
    xImprimePago(ultimo, 5) = pdblMtoSoles
    xImprimePago(ultimo, 6) = pdblMtoVuelto
    xImprimePago(ultimo, 7) = pstrNumTarjeta
    xImprimePago(ultimo, 8) = pstrFechaVencimiento
    xImprimePago(ultimo, 9) = pintNumCuotas
    xImprimePago(ultimo, 10) = pdblTipoCambio
    
    
    
    
End Sub


Private Function cargaParametros(ByVal pstrCodFormato As String) As Boolean
Dim objImpresion As clsImpresion
Dim a As Long
Dim strNombre As String
Dim dblAncho As Double
Dim dblAlto As Double

    On Error GoTo Error
        
        a = Formato.Find(0, 0, pstrCodFormato)
        
        strNombre = Formato(a, 1)
        dblAncho = CDbl(Formato(a, 2))
        dblAlto = CDbl(Formato(a, 3))
        
    
        Set objImpresion = New clsImpresion
        cargaParametros = objImpresion.pfbln_SetFormOFF(strNombre, dblAncho, dblAlto)
        Set objImpresion = Nothing
    
    Exit Function
Error:
    Err.Raise Err.Number, "clsFormato", Err.Description
End Function



Public Sub LimpiaConsultaDoc()
    xConsultaDoc.ReDim 0, -1, 0, 6
End Sub


Public Sub LimpiaFormato()
    xFormato.ReDim 0, -1, 0, 3

End Sub

Public Function AgregaConsultaDoc(ByVal pstrTipoDocumento As String, _
                    ByVal pstrNumDocumento As String, _
                    ByVal strFchEmisión As String, _
                    ByVal dblImporte As Double, _
                    ByVal strNomCliente As String, _
                    ByVal strEstadoDoc As String, _
                    ByVal strUsuario As String) As XArrayDB
    


    Dim ultimo As Integer
    Dim aux As Integer
    If xConsultaDoc.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xConsultaDoc.Count(1)
    While i < aux
        If xConsultaDoc(i, 0) = pstrTipoDocumento And xConsultaDoc(i, 1) = pstrNumDocumento Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xConsultaDoc.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xConsultaDoc.AppendRows
    End If
j:
    If xConsultaDoc.Count(1) = 0 Then ultimo = 0: xConsultaDoc.AppendRows
    
    xConsultaDoc(ultimo, 0) = pstrTipoDocumento
    xConsultaDoc(ultimo, 1) = pstrNumDocumento
    xConsultaDoc(ultimo, 2) = strFchEmisión
    xConsultaDoc(ultimo, 3) = dblImporte
    xConsultaDoc(ultimo, 4) = strNomCliente
    xConsultaDoc(ultimo, 5) = strEstadoDoc
    xConsultaDoc(ultimo, 6) = strUsuario
    Set AgregaConsultaDoc = xConsultaDoc




End Function


Public Property Get ConsultaDoc() As XArrayDB
    Set ConsultaDoc = xConsultaDoc
End Property



Public Function Anula(ByVal pstrTipoDocumento As String, _
                        ByVal pstrNumDocumento As String) As Boolean
    Dim bolResultado As Boolean
    Dim cnn As ADODB.Connection
    Dim strSql As String
    Dim rsDetalleVenta As New ADODB.Recordset
    Dim i As Integer

    On Error GoTo CtrlErr
    bolResultado = False


'        Set cnn = New ADODB.Connection
'        cnn.Open gstrConexion
'        strSQL = "select * from detalleventa.txt"
'        rsDetalleVenta.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
    rsDetalleVenta.CursorLocation = adUseClient
    rsDetalleVenta.Open strDetalleVentaXML, gstrConexion, adOpenDynamic, adLockOptimistic
        
    rsDetalleVenta.Filter = "COD_TIPO_DOCUMENTO = '" & pstrTipoDocumento & _
                            "' AND NUM_DOCUMENTO = '" & pstrNumDocumento & _
                            "' AND TIP_MOVIMIENTO = '" & COD_TIP_MOV_VENTA & "'"
    If rsDetalleVenta.RecordCount > 0 Then
        rsDetalleVenta.MoveFirst
        While Not rsDetalleVenta.EOF
            Call AgregaConsultaDocumento(rsDetalleVenta!SEC_VENTA, _
                                rsDetalleVenta!COD_CIA, rsDetalleVenta!COD_LOCAL, rsDetalleVenta!COD_TIPO_DOCUMENTO, rsDetalleVenta!COD_MAQUINA, _
                                rsDetalleVenta!NUM_DOCUMENTO, rsDetalleVenta!NUM_ITEM, rsDetalleVenta!COD_PRODUCTO, rsDetalleVenta!DES_PRODUCTO, rsDetalleVenta!CTD_PRODUCTO, rsDetalleVenta!FLG_FRACCION, _
                                rsDetalleVenta!PCT_DESCUENTO, rsDetalleVenta!PRC_UNITARIO, rsDetalleVenta!PRC_ORIGINAL, rsDetalleVenta!MTO_IGV, rsDetalleVenta!MTO_EXONERADO, rsDetalleVenta!MTO_SUBTOTAL, _
                                rsDetalleVenta!USU_EMISION, rsDetalleVenta!FCH_EMISION, objOFFUsuario.CodigoUsuario, objOFFUsuario.sysdate, rsDetalleVenta!FLG_MODPRECIO, rsDetalleVenta!FLG_REGALO, _
                                COD_ESTADO_ANU, IIf(IsNull(rsDetalleVenta!TIP_CLIENTE), "", rsDetalleVenta!TIP_CLIENTE), IIf(IsNull(rsDetalleVenta!NUM_RUC), "", rsDetalleVenta!NUM_RUC), IIf(IsNull(rsDetalleVenta!DIR_CLIENTE), "", rsDetalleVenta!DIR_CLIENTE), IIf(IsNull(rsDetalleVenta!NOM_CLIENTE), "", rsDetalleVenta!NOM_CLIENTE), rsDetalleVenta!PCT_IGV, IIf(IsNull(rsDetalleVenta!PARTIDA_ARANCELARIA), "", rsDetalleVenta!PARTIDA_ARANCELARIA), _
                                rsDetalleVenta!CTD_FRACCIONA, IIf(IsNull(rsDetalleVenta!USU_MOD_PRECIO), "", rsDetalleVenta!USU_MOD_PRECIO))
            rsDetalleVenta.MoveNext
        Wend
        For i = 0 To ConsultaDocumento.UpperBound(1)
            rsDetalleVenta.AddNew
            rsDetalleVenta!SEC_VENTA = ConsultaDocumento(i, 0)
            rsDetalleVenta!COD_CIA = ConsultaDocumento(i, 1)
            rsDetalleVenta!COD_LOCAL = ConsultaDocumento(i, 2)
            rsDetalleVenta!COD_TIPO_DOCUMENTO = ConsultaDocumento(i, 3)
            rsDetalleVenta!COD_MAQUINA = ConsultaDocumento(i, 4)
            rsDetalleVenta!NUM_DOCUMENTO = ConsultaDocumento(i, 5)
            rsDetalleVenta!NUM_ITEM = ConsultaDocumento(i, 6)
            rsDetalleVenta!COD_PRODUCTO = ConsultaDocumento(i, 7)
            rsDetalleVenta!DES_PRODUCTO = ConsultaDocumento(i, 8)
            rsDetalleVenta!CTD_PRODUCTO = ConsultaDocumento(i, 9)
            rsDetalleVenta!FLG_FRACCION = ConsultaDocumento(i, 10)
            rsDetalleVenta!PCT_DESCUENTO = ConsultaDocumento(i, 11)
            rsDetalleVenta!PRC_UNITARIO = ConsultaDocumento(i, 12)
            rsDetalleVenta!PRC_ORIGINAL = ConsultaDocumento(i, 13)
            rsDetalleVenta!MTO_IGV = ConsultaDocumento(i, 14)
            rsDetalleVenta!MTO_EXONERADO = ConsultaDocumento(i, 15)
            rsDetalleVenta!MTO_SUBTOTAL = ConsultaDocumento(i, 16)
            rsDetalleVenta!USU_EMISION = ConsultaDocumento(i, 17)
            rsDetalleVenta!FCH_EMISION = ConsultaDocumento(i, 18)
            rsDetalleVenta!USU_ANULACION = ConsultaDocumento(i, 19)
            rsDetalleVenta!FCH_ANULACION = ConsultaDocumento(i, 20)
            rsDetalleVenta!FLG_MODPRECIO = ConsultaDocumento(i, 21)
            rsDetalleVenta!FLG_REGALO = ConsultaDocumento(i, 22)
            rsDetalleVenta!COD_ESTADO = ConsultaDocumento(i, 23)
            rsDetalleVenta!TIP_CLIENTE = ConsultaDocumento(i, 24)
            rsDetalleVenta!NUM_RUC = ConsultaDocumento(i, 25)
            rsDetalleVenta!DIR_CLIENTE = ConsultaDocumento(i, 26)
            rsDetalleVenta!NOM_CLIENTE = ConsultaDocumento(i, 27)
            rsDetalleVenta!PCT_IGV = ConsultaDocumento(i, 28)
            rsDetalleVenta!PARTIDA_ARANCELARIA = ConsultaDocumento(i, 29)
            rsDetalleVenta!CTD_FRACCIONA = ConsultaDocumento(i, 30)
            rsDetalleVenta!USU_MOD_PRECIO = ConsultaDocumento(i, 31)
            rsDetalleVenta!TIP_MOVIMIENTO = COD_TIP_MOV_ANULACION
            rsDetalleVenta!COD_SERIE_REL = objOFFUsuario.CodSerieEtiq
            
            rsDetalleVenta.Update
        Next i
        bolResultado = True
    Else
        Err.Raise vbObjectError + 513, "cls_OFF_Documento.Anula", "No se encontró el documento."
    End If
    
    rsDetalleVenta.Filter = adFilterNone
    rsDetalleVenta.Save strDetalleVentaXML, adPersistXML
    rsDetalleVenta.Close
    'cnn.Close
    Anula = bolResultado
    Exit Function

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.Anula", Err.Description
End Function

Private Sub LimpiaConsultaDocumento()
    xConsultaDocumento.ReDim 0, -1, 0, 31
End Sub

Private Sub LimpiaTipoDocumento()
    xTipoDocumento.ReDim 0, -1, 0, 4
End Sub


Public Property Get ConsultaDocumento() As XArrayDB
    Set ConsultaDocumento = xConsultaDocumento
End Property


Public Sub AgregaConsultaDocumento(ByVal intSecVenta As Integer, ByVal strCodCia As String, ByVal strCodLocal As String, ByVal strCodTipoDocumento As String, _
                                ByVal strCodMaquina As String, ByVal strNumDocumento As String, ByVal intNumItem As Integer, ByVal strCodProducto As String, _
                                ByVal strDesProducto As String, ByVal intCtdProducto As Integer, ByVal strFlgFraccion As String, ByVal dblPctDescuento As Double, _
                                ByVal dblPrcUnitario As Double, ByVal dblPrcOriginal As Double, ByVal dblMtoIgv As Double, ByVal dblMtoExonerado As Double, _
                                ByVal dblMtoSubtotal As Double, ByVal strUsuEmision As String, ByVal strFchEmision As String, ByVal strUsuAnulacion As String, _
                                ByVal strFchAnulacion As String, ByVal strFlgModPrecio As String, ByVal strFlgRegalo As String, ByVal strCodEstado As String, _
                                ByVal strTipCliente As String, ByVal strNumRuc As String, ByVal strDirCliente As String, ByVal strNomCliente As String, _
                                ByVal dblPctIgv As Double, ByVal strPartidaArancelaria As String, ByVal intCtdFracciona As Integer, ByVal strUsuModPrecio As String)
                                
                                
                                
    Dim ultimo As Integer
    Dim i As Integer
    Dim encontro As Boolean
    
    If xConsultaDocumento.Count(1) < 0 Then Exit Sub
    
    
    

    encontro = False
    ultimo = xConsultaDocumento.Count(1)

    If encontro = False Then
        xConsultaDocumento.AppendRows
    End If


    If xConsultaDocumento.Count(1) = 0 Then ultimo = 0: xConsultaDocumento.AppendRows
    
    
    xConsultaDocumento(ultimo, 0) = intSecVenta
    xConsultaDocumento(ultimo, 1) = strCodCia
    xConsultaDocumento(ultimo, 2) = strCodLocal
    xConsultaDocumento(ultimo, 3) = strCodTipoDocumento
    xConsultaDocumento(ultimo, 4) = strCodMaquina
    xConsultaDocumento(ultimo, 5) = strNumDocumento
    xConsultaDocumento(ultimo, 6) = intNumItem
    xConsultaDocumento(ultimo, 7) = strCodProducto
    xConsultaDocumento(ultimo, 8) = strDesProducto
    xConsultaDocumento(ultimo, 9) = intCtdProducto
    xConsultaDocumento(ultimo, 10) = strFlgFraccion
    xConsultaDocumento(ultimo, 11) = dblPctDescuento
    xConsultaDocumento(ultimo, 12) = dblPrcUnitario
    xConsultaDocumento(ultimo, 13) = dblPrcOriginal
    xConsultaDocumento(ultimo, 14) = dblMtoIgv
    xConsultaDocumento(ultimo, 15) = dblMtoExonerado
    xConsultaDocumento(ultimo, 16) = dblMtoSubtotal
    xConsultaDocumento(ultimo, 17) = strUsuEmision
    xConsultaDocumento(ultimo, 18) = strFchEmision
    xConsultaDocumento(ultimo, 19) = strUsuAnulacion
    xConsultaDocumento(ultimo, 20) = strFchAnulacion
    xConsultaDocumento(ultimo, 21) = strFlgModPrecio
    xConsultaDocumento(ultimo, 22) = strFlgRegalo
    xConsultaDocumento(ultimo, 23) = strCodEstado
    xConsultaDocumento(ultimo, 24) = strTipCliente
    xConsultaDocumento(ultimo, 25) = strNumRuc
    xConsultaDocumento(ultimo, 26) = strDirCliente
    xConsultaDocumento(ultimo, 27) = strNomCliente
    xConsultaDocumento(ultimo, 28) = dblPctIgv
    xConsultaDocumento(ultimo, 29) = strPartidaArancelaria
    xConsultaDocumento(ultimo, 30) = intCtdFracciona
    xConsultaDocumento(ultimo, 31) = strUsuModPrecio
    
End Sub


Public Function Reactiva(ByVal pstrTipoDocumento As String, ByVal pstrNumDocumento As String) As Boolean
    Dim bolResultado As Boolean
    Dim cnn As ADODB.Connection
    Dim strSql As String
    Dim rsDetalleVenta As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo CtrlErr
    bolResultado = False

'        Set cnn = New ADODB.Connection
'        cnn.Open gstrConexion
'        strSQL = "select * from detalleventa.txt"
'        rsDetalleVenta.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
    rsDetalleVenta.CursorLocation = adUseClient
    rsDetalleVenta.Open strDetalleVentaXML, gstrConexion, adOpenDynamic, adLockOptimistic
        
    rsDetalleVenta.Filter = "COD_TIPO_DOCUMENTO = '" & pstrTipoDocumento & "' AND NUM_DOCUMENTO = '" & pstrNumDocumento & "' AND TIP_MOVIMIENTO = '" & COD_TIP_MOV_VENTA & "'"
    If rsDetalleVenta.RecordCount > 0 Then
        rsDetalleVenta.MoveFirst
        While Not rsDetalleVenta.EOF
            Call AgregaConsultaDocumento(rsDetalleVenta!SEC_VENTA, _
                                rsDetalleVenta!COD_CIA, rsDetalleVenta!COD_LOCAL, rsDetalleVenta!COD_TIPO_DOCUMENTO, rsDetalleVenta!COD_MAQUINA, _
                                rsDetalleVenta!NUM_DOCUMENTO, rsDetalleVenta!NUM_ITEM, rsDetalleVenta!COD_PRODUCTO, rsDetalleVenta!DES_PRODUCTO, rsDetalleVenta!CTD_PRODUCTO, rsDetalleVenta!FLG_FRACCION, _
                                rsDetalleVenta!PCT_DESCUENTO, rsDetalleVenta!PRC_UNITARIO, rsDetalleVenta!PRC_ORIGINAL, rsDetalleVenta!MTO_IGV, rsDetalleVenta!MTO_EXONERADO, rsDetalleVenta!MTO_SUBTOTAL, _
                                rsDetalleVenta!USU_EMISION, rsDetalleVenta!FCH_EMISION, "", "", rsDetalleVenta!FLG_MODPRECIO, rsDetalleVenta!FLG_REGALO, _
                                "EMI", IIf(IsNull(rsDetalleVenta!TIP_CLIENTE), "", rsDetalleVenta!TIP_CLIENTE), IIf(IsNull(rsDetalleVenta!NUM_RUC), "", rsDetalleVenta!NUM_RUC), IIf(IsNull(rsDetalleVenta!DIR_CLIENTE), "", rsDetalleVenta!DIR_CLIENTE), IIf(IsNull(rsDetalleVenta!NOM_CLIENTE), "", rsDetalleVenta!NOM_CLIENTE), rsDetalleVenta!PCT_IGV, IIf(IsNull(rsDetalleVenta!PARTIDA_ARANCELARIA), "", rsDetalleVenta!PARTIDA_ARANCELARIA), _
                                rsDetalleVenta!CTD_FRACCIONA, IIf(IsNull(rsDetalleVenta!USU_MOD_PRECIO), "", rsDetalleVenta!USU_MOD_PRECIO))
            rsDetalleVenta.MoveNext
        Wend
        For i = 0 To ConsultaDocumento.UpperBound(1)
                rsDetalleVenta.AddNew
                rsDetalleVenta!SEC_VENTA = ConsultaDocumento(i, 0)
                rsDetalleVenta!COD_CIA = ConsultaDocumento(i, 1)
                rsDetalleVenta!COD_LOCAL = ConsultaDocumento(i, 2)
                rsDetalleVenta!COD_TIPO_DOCUMENTO = ConsultaDocumento(i, 3)
                rsDetalleVenta!COD_MAQUINA = ConsultaDocumento(i, 4)
                rsDetalleVenta!NUM_DOCUMENTO = ConsultaDocumento(i, 5)
                rsDetalleVenta!NUM_ITEM = ConsultaDocumento(i, 6)
                rsDetalleVenta!COD_PRODUCTO = ConsultaDocumento(i, 7)
                rsDetalleVenta!DES_PRODUCTO = ConsultaDocumento(i, 8)
                rsDetalleVenta!CTD_PRODUCTO = ConsultaDocumento(i, 9)
                rsDetalleVenta!FLG_FRACCION = ConsultaDocumento(i, 10)
                rsDetalleVenta!PCT_DESCUENTO = ConsultaDocumento(i, 11)
                rsDetalleVenta!PRC_UNITARIO = ConsultaDocumento(i, 12)
                rsDetalleVenta!PRC_ORIGINAL = ConsultaDocumento(i, 13)
                rsDetalleVenta!MTO_IGV = ConsultaDocumento(i, 14)
                rsDetalleVenta!MTO_EXONERADO = ConsultaDocumento(i, 15)
                rsDetalleVenta!MTO_SUBTOTAL = ConsultaDocumento(i, 16)
                rsDetalleVenta!USU_EMISION = ConsultaDocumento(i, 17)
                rsDetalleVenta!FCH_EMISION = ConsultaDocumento(i, 18)
                rsDetalleVenta!USU_ANULACION = ConsultaDocumento(i, 19)
                rsDetalleVenta!FCH_ANULACION = ConsultaDocumento(i, 20)
                rsDetalleVenta!FLG_MODPRECIO = ConsultaDocumento(i, 21)
                rsDetalleVenta!FLG_REGALO = ConsultaDocumento(i, 22)
                rsDetalleVenta!COD_ESTADO = ConsultaDocumento(i, 23)
                rsDetalleVenta!TIP_CLIENTE = ConsultaDocumento(i, 24)
                rsDetalleVenta!NUM_RUC = ConsultaDocumento(i, 25)
                rsDetalleVenta!DIR_CLIENTE = ConsultaDocumento(i, 26)
                rsDetalleVenta!NOM_CLIENTE = ConsultaDocumento(i, 27)
                rsDetalleVenta!PCT_IGV = ConsultaDocumento(i, 28)
                rsDetalleVenta!PARTIDA_ARANCELARIA = ConsultaDocumento(i, 29)
                rsDetalleVenta!CTD_FRACCIONA = ConsultaDocumento(i, 30)
                rsDetalleVenta!USU_MOD_PRECIO = ConsultaDocumento(i, 31)
                rsDetalleVenta!TIP_MOVIMIENTO = COD_TIP_MOV_REACTIVACION
                rsDetalleVenta!COD_SERIE_REL = objOFFUsuario.CodSerieEtiq
                rsDetalleVenta.Update
           Next i
            bolResultado = True
        Else
            Err.Raise vbObjectError + 513, "cls_OFF_Documento.Anula", "No se encontró el documento  "
        End If
        rsDetalleVenta.Filter = adFilterNone
        rsDetalleVenta.Save strDetalleVentaXML, adPersistXML
        rsDetalleVenta.Close
'        cnn.Close
    Reactiva = bolResultado

Exit Function
CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.Reactiva", Err.Description

End Function



Public Sub CreaArchivoTexto()
    Dim pp$
    On Error GoTo CtrlErr


        If objVenta.debeRegularizar = False And Not objOFFVenta.Existe(strDetalleVentaXML) = 0 Then
            Kill strDetalleVentaXML
        End If
        If objVenta.debeRegularizar = False And Not objOFFVenta.Existe(strPagoVentaXML) = 0 Then
            Kill strPagoVentaXML
        End If

'    If objOFFVenta.Existe(App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "detalleventa.txt") = 0 Then
    If objOFFVenta.Existe(strDetalleVentaXML) = 0 Then
        ''pp = "SEC_VENTA;COD_CIA;COD_LOCAL;COD_TIPO_DOCUMENTO;COD_MAQUINA;NUM_DOCUMENTO;NUM_ITEM;COD_PRODUCTO;DES_PRODUCTO;CTD_PRODUCTO;FLG_FRACCION;PCT_DESCUENTO;PRC_UNITARIO;PRC_ORIGINAL;MTO_IGV;MTO_EXONERADO;MTO_SUBTOTAL;USU_EMISION;FCH_EMISION;USU_ANULACION;FCH_ANULACION;FLG_MODPRECIO;FLG_REGALO;COD_ESTADO;TIP_CLIENTE;NUM_RUC;DIR_CLIENTE;NOM_CLIENTE;PCT_IGV;PARTIDA_ARANCELARIA;CTD_FRACCIONA;USU_MOD_PRECIO;TIP_MOVIMIENTO"
'        Call EscribeArchivoTexto(App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "detalleventa.txt", pp)
        ''Call EscribeArchivoTexto(strDetalleVentaXML, pp)
        
        Call CreaFormatoDetalleVenta
        
    End If
    
        
    
'    If objOFFVenta.Existe(App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "pagoventa.txt") = 0 Then
    If objOFFVenta.Existe(strPagoVentaXML) = 0 Then
        ''pp = "SEC_VENTA;NUM_ITEM;COD_PAGO;COD_MONEDA;MTO_IMPORTE;MTO_SOLES;MTO_VUELTO;NUM_TARJETA;FCH_VENCIMIENTO;NUM_CUOTAS;TIPO_CAMBIO;TIPO_CUOTA;NUM_AUTORIZACION"
'        Call EscribeArchivoTexto(App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "pagoventa.txt", pp)
        ''Call EscribeArchivoTexto(strPagoVentaXML, pp)
        
        Call CreaFormatoPagoVenta
    Else
    End If

    Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.CreaArchivoTexto", Err.Description
End Sub

Private Sub EscribeArchivoTexto(ByVal pa$, ByVal pp$)
Dim NumArch As Integer
On Error GoTo CtrlErr

    NumArch = FreeFile

    Open pa For Append As NumArch

    Print #NumArch, pp

    Close #NumArch


Exit Sub

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Documento.EscribeArchivoTexto", Err.Description

End Sub



Public Function AgregaTipoDocumento(ByVal pstrCodDocumento As String, _
                    ByVal pstrDesDocumento As String, _
                    ByVal pstrNumDocumento As String, _
                    ByVal pintNumLinea As Integer, _
                    ByVal pintAncho As Integer) As XArrayDB
    


    Dim ultimo As Integer
    Dim aux As Integer
    If xTipoDocumento.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xTipoDocumento.Count(1)
    While i < aux
        If xTipoDocumento(i, 0) = pstrCodDocumento Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xTipoDocumento.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xTipoDocumento.AppendRows
    End If
j:
    If xTipoDocumento.Count(1) = 0 Then ultimo = 0: xTipoDocumento.AppendRows
    
    xTipoDocumento(ultimo, 0) = pstrCodDocumento
    xTipoDocumento(ultimo, 1) = pstrDesDocumento
    xTipoDocumento(ultimo, 2) = pstrNumDocumento
    xTipoDocumento(ultimo, 3) = pintNumLinea
    xTipoDocumento(ultimo, 4) = pintAncho
    Set AgregaTipoDocumento = xTipoDocumento




End Function

Public Function AgregaFormato(ByVal pstrCodFormato As String, _
                    ByVal pstrDesFormato As String, _
                    ByVal pstrAncho As String, _
                    ByVal pstrAlto As String) As XArrayDB
    


    Dim ultimo As Integer
    Dim aux As Integer
    If xFormato.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xFormato.Count(1)
    While i < aux
        If xFormato(i, 0) = pstrCodFormato Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xFormato.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xFormato.AppendRows
    End If
j:
    If xFormato.Count(1) = 0 Then ultimo = 0: xFormato.AppendRows
    
    xFormato(ultimo, 0) = pstrCodFormato
    xFormato(ultimo, 1) = pstrDesFormato
    xFormato(ultimo, 2) = pstrAncho
    xFormato(ultimo, 3) = pstrAlto
    Set AgregaFormato = xTipoDocumento




End Function




Private Sub CreaFormatoPagoVenta()
Dim pp$

On Error GoTo CtrlErr

pp = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "    xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "    xmlns:rs='urn:schemas-microsoft-com:rowset'"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "    xmlns:z='#RowsetSchema'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "<s:Schema id='RowsetSchema'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "    <s:ElementType name='row' content='eltOnly'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='SEC_VENTA' rs:number='1' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='NUM_ITEM' rs:number='2' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='COD_PAGO' rs:number='3' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='COD_MONEDA' rs:number='4' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='MTO_IMPORTE' rs:number='5' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='MTO_SOLES' rs:number='6' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='MTO_VUELTO' rs:number='7' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='NUM_TARJETA' rs:number='8' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='50'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='FCH_VENCIMIENTO' rs:number='9' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='NUM_CUOTAS' rs:number='10' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='TIPO_CUOTA' rs:number='11' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='TIPO_CAMBIO' rs:number='12' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:AttributeType name='NUM_AUTORIZACION' rs:number='13' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "        <s:extends type='rs:rowbase'/>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "    </s:ElementType>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "</s:Schema>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "<rs:data>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "</rs:data>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)
pp = "</xml>"
Call EscribeArchivoTexto(strPagoVentaXML, pp)

Exit Sub
CtrlErr:

    Err.Raise Err.Number, "cls_OFF_Documento.CreaFormatoPagoVenta", Err.Description

End Sub


Private Sub CreaFormatoDetalleVenta()
Dim pp$

On Error GoTo CtrlErr

pp = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "    xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "    xmlns:rs='urn:schemas-microsoft-com:rowset'"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "    xmlns:z='#RowsetSchema'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "<s:Schema id='RowsetSchema'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "    <s:ElementType name='row' content='eltOnly'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='SEC_VENTA' rs:number='1' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='COD_CIA' rs:number='2' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='3'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='COD_LOCAL' rs:number='3' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='3'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='COD_TIPO_DOCUMENTO' rs:number='4' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='3'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='COD_MAQUINA' rs:number='5' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='NUM_DOCUMENTO' rs:number='6' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='NUM_ITEM' rs:number='7' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='COD_PRODUCTO' rs:number='8' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='5'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='DES_PRODUCTO' rs:number='9' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='100'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='CTD_PRODUCTO' rs:number='10' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='FLG_FRACCION' rs:number='11' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='1'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='PCT_DESCUENTO' rs:number='12' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='PRC_UNITARIO' rs:number='13' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='PRC_ORIGINAL' rs:number='14' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='MTO_IGV' rs:number='15' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='MTO_EXONERADO' rs:number='16' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='MTO_SUBTOTAL' rs:number='17' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='USU_EMISION' rs:number='18' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='5'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='FCH_EMISION' rs:number='19' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='USU_ANULACION' rs:number='20' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='5'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='FCH_ANULACION' rs:number='21' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='FLG_MODPRECIO' rs:number='22' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='1'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='FLG_REGALO' rs:number='23' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='1'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='COD_ESTADO' rs:number='24' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='TIP_CLIENTE' rs:number='25' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='1'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='NUM_RUC' rs:number='26' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='DIR_CLIENTE' rs:number='27' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='100'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='NOM_CLIENTE' rs:number='28' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='100'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='PCT_IGV' rs:number='29' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='PARTIDA_ARANCELARIA' rs:number='30' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='100'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='CTD_FRACCIONA' rs:number='31' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='30'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='USU_MOD_PRECIO' rs:number='32' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='5'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"

Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='TIP_MOVIMIENTO' rs:number='33' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='10'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"

Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:AttributeType name='COD_SERIE_REL' rs:number='33' rs:nullable='true' rs:writeunknown='true'>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "            <s:datatype dt:type='string' rs:dbtype='str' dt:maxLength='20'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        </s:AttributeType>"

Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "        <s:extends type='rs:rowbase'/>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "    </s:ElementType>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "</s:Schema>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "<rs:data>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "</rs:data>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
pp = "</xml>"
Call EscribeArchivoTexto(strDetalleVentaXML, pp)
Exit Sub
CtrlErr:

    Err.Raise Err.Number, "cls_OFF_Documento.CreaFormatoPagoVenta", Err.Description


End Sub
