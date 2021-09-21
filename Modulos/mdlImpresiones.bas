Attribute VB_Name = "mdlImpresiones"
Enum TipoAlineacion
    Izquierda = 0
    Derecha = 1
    centro = 2
End Enum

Public pstrDevImpGuia$

Global gNroCopia%

Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
          ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
          
Public intNumImpresion As Integer


Function pfstr_Alineado(ByVal vvarCad As Variant, ByVal vintTam As Integer, Optional ByVal vvarAlineacion As TipoAlineacion = Izquierda, Optional ByVal vstrFormato As String) As String
Dim strSalida As String
Dim strCadena As String
    If vstrFormato <> "" Then strCadena = Format(vvarCad, vstrFormato) Else strCadena = "" & vvarCad
    Select Case vvarAlineacion
        Case TipoAlineacion.centro
            If Len(strCadena) < vintTam Then
                strSalida = Space((vintTam - Len(strCadena)) \ 2) & strCadena & Space(vintTam - Len(strCadena) - (vintTam - Len(strCadena)) \ 2)
            Else
                strSalida = left(strCadena, vintTam)
            End If
        Case TipoAlineacion.Izquierda
            strSalida = left(strCadena & Space(vintTam), vintTam)
        Case TipoAlineacion.Derecha
            strSalida = right(Space(vintTam) & strCadena, vintTam)
    End Select
    pfstr_Alineado = strSalida
End Function

Function fbln_SetImpresora(ByVal vstrDevice$, ByVal vstrImpDoc$) As Boolean
Dim objImpresora As Printer
    fbln_SetImpresora = False
    vstrDevice = pfstr_Leer_Cadena_INI("Impresion", vstrImpDoc, pstrArchivo)
    For Each objImpresora In Printers
      If objImpresora.Devicename = vstrDevice Then
         Set Printer = objImpresora
         fbln_SetImpresora = True
         Exit For
      End If
    Next objImpresora
End Function


Sub psub_Imprime_Guia_Cliente(ByVal vstrNumguia As String)
Dim odynCab As oraDynaset
Dim odynDet As oraDynaset
Dim strSQL As String
Dim Copia As Double
Dim objFormato As New clsImpresion
 
    'Cabecera'
    Set odynCab = gclsOracle.FN_Cursor("BTLCERO.PKG_GUIA_CLIENTE.FN_CABECERA", 0, vstrNumguia)
    If odynCab.RecordCount = 0 Then MsgBox "No se encontró Guía en la BD", vbCritical, "ERROR": Exit Sub
        
    'Detalle'
    Set odynDet = gclsOracle.FN_Cursor("BTLCERO.PKG_GUIA_CLIENTE.FN_DETALLE", 0, vstrNumguia, odynCab("NUM_PEDIDO").Value, "1")
    If odynDet.RecordCount = 0 Then MsgBox "Documento sin detalle", vbCritical, "ERROR": Exit Sub
'--------------------------------------------------------------------------------------------------------------'
    frm_VTA_Impresoras.Show vbModal
    If gNroCopia = 0 Then Exit Sub

    Dim intLinea%
    
  
    Copia = 0
    On Error GoTo 0
    Do While gNroCopia > Copia
        odynDet.MoveFirst
        While Not odynDet.EOF
            'Cabecera
            On Error GoTo Mal_Seteo
               
'            If Not pfbln_SetForm("001") Then
'                Exit Sub
'            End If

    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
              If Not objFormato.cargaParametros("001") Then
                Exit Sub
            End If
    Else
            Printer.Width = 225 * 56.7
            Printer.Height = 153 * 56.7 '152*56.7, ' 93.3 * 56.7 ' 280
    End If
                
            GoTo OK_Seteo
Mal_Seteo:
            MsgBox Err.Description, vbExclamation, "ERROR EN IMPRESORA"
            Exit Sub
OK_Seteo:
            '-----------
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.FontName = "Draft 10cpi"
            Printer.Print Space(55) & Format(odynCab("NUM_GUIA").Value, "000-0000000")
            Printer.FontName = "Draft 12cpi"
            Printer.Print
            Printer.Print
            Printer.Print
                        
            Printer.Print "DROGUERIA"
            Printer.Print
            Printer.Print Space(18) & Format(odynCab("FCH_EMISION").Value, "DD") & " de " & left(Format(odynCab("FCH_EMISION").Value, "MMMM") & Space(10), 10) & " del " & Format(odynCab("FCH_EMISION").Value, "YYYY") & Space(46) & "X"
            Printer.Print
            Printer.Print
            
            Dim strDestinatario$, strOC$
                          
            If odynCab("BROKER").Value <> "*" Then
                strDestinatario = odynCab("BROKER").Value
            End If
            
            If odynCab("OBS_ORDEN_COMPRA").Value <> "*" Then
                strOC = "O/C Cliente : " & odynCab("OBS_ORDEN_COMPRA").Value
            End If
            Printer.Print Space(23) & strDestinatario
            Printer.Print Space(23) & strOC
            Printer.Print Space(23) & odynCab("CLIENTE").Value
            Printer.FontName = "Draft 17cpi"
            Printer.Print Space(32) & odynCab("DIR_ENTREGA_PROD").Value
            Printer.FontName = "Draft 10cpi"
            Printer.Print Space(10) & odynCab("PERSONA").Value
            Printer.FontName = "Draft 17cpi"
            Printer.Print String(130, "-")
            Printer.Print right(Space(2) & "#", 2) & Space(1) & _
                           right(Space(6) & "Cant.", 6) & Space(1) & _
                           left("Código" & Space(6), 6) & Space(2) & _
                           left("Descripción" & Space(60), 60) & Space(2) & _
                           left("Laboratorio" & Space(15), 15) & Space(2) & _
                           left("Lote" & Space(15), 15) & Space(1) & _
                           left("F.Ven" & Space(7), 7) & Space(2) & _
                           right(Space(7) & "Ctd.Lte", 7)
            Printer.Print String(130, "-")
          '---
            intLinea = 0
            Dim strCodProdAnte As String
            Dim j%
            strCodProdAnte = "*"
            j = 1
            While Not odynDet.EOF
                  'Detalle
                  If strCodProdAnte <> odynDet("COD_PRODUCTO").Value Then
                    Printer.Print right(Space(2) & j, 2) & Space(1) & _
                                   right(Space(6) & odynDet("CANTIDADES").Value, 6) & Space(1) & _
                                   left(odynDet("COD_PRODUCTO").Value & Space(6), 6) & Space(2) & _
                                   left(odynDet("DES_PRODUCTO").Value & Space(60), 60) & Space(2) & _
                                   left(odynDet("LAB").Value & Space(15), 15) & Space(2) & _
                                   left(odynDet("NUM_LOTE").Value & Space(15), 15) & Space(1) & _
                                   left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(2) & _
                                   right(Space(7) & odynDet("CANTIDADES_LOTE").Value, 7)
                                   j = j + 1
                  Else
                    Printer.Print right(Space(2) & "", 2) & Space(1) & _
                                   right(Space(6) & "", 6) & Space(1) & _
                                   left("" & Space(6), 6) & Space(2) & _
                                   left("" & Space(60), 60) & Space(2) & _
                                   left("" & Space(15), 15) & Space(2) & _
                                   left(odynDet("NUM_LOTE").Value & Space(15), 15) & Space(1) & _
                                   left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(2) & _
                                   right(Space(7) & odynDet("CANTIDADES_LOTE").Value, 7)
                  End If
                  strCodProdAnte = odynDet("COD_PRODUCTO").Value
                  intLinea = intLinea + 1
                  odynDet.MoveNext
            Wend
            Printer.Print
            Dim i%, intMaxLineas%
            intMaxLineas = Val("" & pfstr_Val_Parametro_Local(gstrLocal, "NUM_LIN_GR"))
            
            For i = 1 To intMaxLineas - intLinea
                Printer.Print
            Next i
            Printer.Print
            Dim strObs$
            If odynCab("OBS_GUIA_CLIENTE").Value <> "*" Then
                strObs = "Obs :" & odynCab("OBS_GUIA_CLIENTE").Value
            End If
            Printer.FontName = "Draft 17cpi"
            Printer.Print strObs
            Printer.Print Space(10) & gclsOracle.FN_Valor("BTLCERO.PKG_PEDIDO.FN_MSJ_BOTIQUIN", odynCab("NUM_GUIA").Value)
            Printer.FontName = "Draft 15cpi"
            
            Printer.EndDoc
        Wend
        Copia = Copia + 1
       Loop
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


Function pfstr_Val_Parametro_Local(ByVal vstrCod_Btl$, ByVal vstrCod_Parametro$) As String
    pfstr_Val_Parametro_Local = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", vstrCod_Btl, vstrCod_Parametro)
End Function

Function pfstr_Val_Parametro(ByVal vstrCod_Parametro$) As String
    pfstr_Val_Parametro = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", vstrCod_Parametro)
End Function


Public Sub psub_Selecionar_Todo(Optional ByRef robjTxt)
Dim objTXT As Object
    Set objTXT = IIf(IsMissing(robjTxt), Screen.ActiveControl, robjTxt)
    On Error GoTo Error
    If TypeName(objTXT) = "TextBox" Or TypeName(objTXT) = "MaskEdBox" Then
    'If Not (TypeName(objTXT) = "TextBox" Or TypeName(objTXT) = "MaskEdBox" Or UCase(TypeName(objTXT)) Like "*COMBO*") Then Exit Sub
    objTXT.SelStart = 0
    objTXT.SelLength = Len(objTXT.Text)
    End If
Error:
    On Error GoTo 0
End Sub

Sub psub_Foco(Optional ByRef rObjText, Optional ByVal vintFontSizeFoco%, Optional vintFontSizeNoFoco%)
Dim objText As Object
    On Error GoTo Salida
    Set objText = IIf(IsMissing(rObjText), Screen.ActiveControl, rObjText)
    If Not (TypeName(objText) = "TextBox" Or TypeName(objText) = "MaskEdBox") Then Exit Sub
    objText.BackColor = IIf(IsMissing(rObjText), gvarColor, vbWhite)
    objText.ForeColor = IIf(IsMissing(rObjText), gvarFColor, vbWindowText)
    objText.Font.Bold = (IsMissing(rObjText) = True)
    objText.Font.Size = IIf(IsMissing(rObjText), IIf(vintFontSizeFoco = 0, 10, vintFontSizeFoco), IIf(vintFontSizeNoFoco = 0, 10, vintFontSizeNoFoco))
Salida:
    On Error GoTo 0
End Sub


Sub psub_Imprime_Doc_Cliente(ByVal vstrCia As String, _
                            ByVal vstrNum_Documento As String, _
                            ByVal vstrTip_Documento As String)
Dim strSQL$, strIgv$, strNumRuc$
Dim odynCab As oraDynaset, odynDet As oraDynaset
Dim strImporteLetras$, strImporteResto$, sngCtv As Single

Dim i%, intNumLineas%, intLinea%
    Set odynCab = gclsOracle.FN_Cursor("BTLCERO.PKG_DOC.FN_CABECERA", 0, vstrCia, vstrTip_Documento, vstrNum_Documento)
    If odynCab.RecordCount = 0 Then MsgBox "No se encontro Documento en la BD", vbCritical, "ERROR": Exit Sub
    Dim strNumLineas$
    
    'PARA EL CALCULO DEL PARAMETRO SEGUN TIPO DE DOCUMENTO Y SU FORMATO
    strNumLineas = IIf(vstrTip_Documento = "BOL", "NUM_LIN_BO", IIf(odynCab("FLG_FACT_CORTA").Value = "1", "NUM_LIN_FC", "NUM_LIN_FA"))
     
    intNumLineas = Val(pfstr_Val_Parametro_Local("000", strNumLineas))   'NUMERO DE LINEAS SEGUN EL PARAMETRO (strNumLineas)
    strNumRuc = pfstr_Val_Parametro("NUMRUC_BTL") 'NUMERO DE RUC DE BTL
    
    Dim blnFacturaCorta As Boolean
    blnFacturaCorta = ("" & odynCab("FLG_FACT_CORTA").Value = "1")
       
    Set odynDet = gclsOracle.FN_Cursor("BTLCERO.PKG_DOC.FN_DETALLE_IMPRIMIR", 0, vstrCia, vstrTip_Documento, vstrNum_Documento)
    If odynDet.RecordCount = 0 Then MsgBox "Documento sin detalle", vbCritical, "ERROR": Exit Sub
    
    strIgv = odynDet("PCT_IGV").Value
    
    
    Dim strSec_Caja$, strFch_Emision$, strNomCliente$, strDirEntrega$, strRuc_Cliente$, _
        strUsuPed$, strUsuDoc$, strTipoVenta$, strGuia$, strOC$
    Dim dblTotal As Double, dblIGV As Double, dblSubTotal As Double
    
    
    strFch_Emision = odynCab("FCH_EMISION").Value
    strNomCliente = odynCab("NOMBRE").Value
    strDirEntrega = IIf(IsNull(odynCab("DIR_ENTREGA_PROD").Value), "", odynCab("DIR_ENTREGA_PROD").Value)
    strRuc_Cliente = "" & odynCab("RUC_CLIENTE").Value
    strUsuPed = IIf(IsNull(odynCab("USU_PEDIDO").Value), "", odynCab("USU_PEDIDO").Value)
    strUsuDoc = IIf(IsNull(odynCab("USU_EMISION").Value), "", odynCab("USU_EMISION").Value)
    dblTotal = Val(odynCab("MTO_TOTAL").Value)
    dblIGV = Val(odynCab("MTO_IMPUESTO").Value)
    dblSubTotal = Val(odynCab("MTO_BASE_IMP").Value) + Val(odynCab("MTO_EXONERADO").Value)
    strTipoVenta = IIf(IsNull(odynCab("CON_VTA").Value), "", odynCab("CON_VTA").Value)
    
'    If odynCab("COD_CLIENTE_IMP").Value = "0000000408" And Not IsNull(odynCab("COD_CONVENIO").Value) Then
'        strTipoVenta = "CREDITO 45 DIAS"
'    End If
    
'    strGuia = "" & IIf(IsNull(odynCab("NUM_GUIA").Value), "", odynCab("NUM_GUIA").Value)
'    strOC = "" & IIf(IsNull(odynCab("OBS_ORDEN_COMPRA").Value), "", odynCab("OBS_ORDEN_COMPRA").Value)
    
    If intNumImpresion = 0 Then
        frm_VTA_Impresoras.Show vbModal
        intNumImpresion = 1
    End If
    
    If gNroCopia = 0 Then Exit Sub
    On Error GoTo ErrorImpresora
    
    
    Select Case vstrTip_Documento
        Case "BOL" 'PARA EL CASO DE BOLETA
            'Printer.Width = 225 * 56.7
            'Printer.Height = (280 / 3) * 56.7
            Printer.Font.name = "Draft 17cpi"

            'MODIFICAR AQUI
            'If Not pfbln_SetForm("004") Then
                'Exit Sub
            'End If

            'Printer.Print "."
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print Space(6) & left("" & Space(40), 40) & _
            Format(strFch_Emision, "DD/MM/YY") & " " & vstrNum_Documento
            Printer.Print Space(2) & strNomCliente

            Printer.Print
            Printer.Print
            intLinea = 0
            While Not odynDet.EOF
                Printer.Print left(Trim(odynDet("COD_PRODUCTO").Value) & Space(5), 5) & Space(1) & _
                left(odynDet("DES_PRODUCTO").Value & Space(40), 40) & _
                right(Space(5) & odynDet("CANTIDADES").Value, 5) & Space(1) & _
                right(Space(1) & " ", 1) & Space(1) & _
                right(Space(10) & Format(odynDet("SUBTOTAL").Value, "###,##0.00"), 10)
                odynDet.MoveNext
                intLinea = intLinea + 1
            Wend

            For i = intLinea To intNumLineas
                Printer.Print
            Next i

            'Printer.Print
            sngCtv = Val(Replace(odynCab("MTO_TOTAL").Value, " ", ""))
            strImporteLetras = UCase(pfstr_Letra(Int(sngCtv), 80))
            strImporteResto = right(Format(sngCtv - Int(sngCtv), ".00"), 2)
            Printer.Print "     SON : " & Trim(strImporteLetras) & " Y " & strImporteResto & "/100 NUEVOS SOLES"

            Printer.Print "Boticas BTL      " & Space(22 - gintDecTot) & "Total Venta S/.  " & right(Space(8 + gintDecTot) & Format(odynCab("MTO_TOTAL").Value, "###,##0." & String(gintDecTot, "0")), 8 + gintDecTot)
            'Printer.Print "Sub total S/. " & Right(Space(7 + gintDec) & Format(dblSubTotal, "#,##0." & String(gintDec, "0")), 7 + gintDec) & _
                          Space(4 - gintDec) & " IGV S/. " & Right(Space(6 + gintDecTot) & Format(dblIGV, "#,##0." & String(gintDec, "0")), 6 + gintDec) & _
                          Space(4 - gintDecTot) & " Total S/. " & Right(Space(8 + gintDecTot) & Format(odynCab("MTO_TOTAL").Value, "###,##0." & String(gintDecTot, "0")), 8 + gintDecTot)
            Printer.Print
            Printer.Print "NIT " & strNumRuc & Space(26) & strSec_Caja & "   " & strUsuPed & "   " & strUsuDoc
            Printer.EndDoc

        Case "FAC"    'PARA EL CASO DE FACTURA
          
                        
            Printer.FontName = "Draft 10cpi"
            If blnFacturaCorta = False Then '---------- PARA LA CABECERA EN CASO NO ES FACTURA CORTA
                             
                Printer.Print
                Printer.Print Space(18) & strSec_Caja & "   " & strUsuPed & "   " & strUsuDoc
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print "DROGUERIA"
                Printer.Print
                Printer.Print Space(8) & Day(strFch_Emision) & " de " & Format(strFch_Emision, "mmmm") & " del " & Year(strFch_Emision)
                Printer.Print Space(6) & left(strNomCliente & Space(45), 45) & Format(strFch_Emision, "dd/mm/yyyy") & " " & left(vstrNum_Documento, 3) & "-" & right(vstrNum_Documento, 7)
                Printer.Print
                Printer.Print Space(8) & strRuc_Cliente & Space(33) & strTipoVenta & Space(3) & Trim(strOC)
                Printer.FontName = "Draft 17cpi"
                Printer.Print Space(8) & left(strDirEntrega & Space(90), 90) & Trim(strGuia)
                Printer.Print
                Printer.Print
                Printer.Print
            Else                           '---------- PARA LA CABECERA EN CASO ES FACTURA CORTA
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print Space(18) & strSec_Caja & "   " & strUsuPed & "   " & strUsuDoc
                Printer.Print "DROGUERIA"
                Printer.Print
                Printer.Print
                Printer.Print Space(8) & Day(strFch_Emision) & " de " & Format(strFch_Emision, "mmmm") & " del " & Year(strFch_Emision)
                Printer.Print Space(48) & Format(strFch_Emision, "dd/mm/yyyy") & " " & left(vstrNum_Documento, 3) & "-" & right(vstrNum_Documento, 7)
                Printer.Print Space(54) & strTipoVenta
                Printer.Print
                Printer.FontName = "Draft 12cpi"
                Printer.Print Space(8) & left(strNomCliente & Space(50), 50) & " OC: " & Trim(strOC)
                Printer.FontName = "Draft 10cpi"
                Printer.Print
                Printer.Print Space(8) & strRuc_Cliente & Space(10) & Trim(strGuia)
                Printer.Print
                Printer.FontName = "Draft 17cpi"
                Printer.Print Space(8) & left(strDirEntrega & Space(80), 80)
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
            End If

            

            intLinea = 0
            Dim strCodProdAnt As String
            strCodProdAnt = "*"
            While Not odynDet.EOF
                
                If blnFacturaCorta = False Then 'PARA EL DETALLE DE LA FACTURA EN CASO NO SEA FACTURA CORTA
                    
                    Printer.FontName = "Draft 17cpi"
                    If strCodProdAnt <> odynDet("COD_PRODUCTO").Value Then
                        Printer.Print left(Trim(odynDet("CANTIDADES").Value) & Space(6), 6) & Space(1) & _
                                left(odynDet("COD_PRODUCTO").Value & " " & odynDet("DES_PRODUCTO").Value & Space(55), 55) & Space(1) & _
                                left("" & odynDet("NUM_LOTE").Value & Space(10), 10) & Space(1) & _
                                left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(1) & _
                                right("" & odynDet("CANTIDADES_LOTE").Value & Space(7), 7) & Space(1) & _
                                right(Space(6 + gintDec) & Format(odynDet("UNITARIO").Value, "# ##0." & String(gintDec, "0")), 6 + gintDec) & Space(18 - gintDec) & _
                                right(Space(10 + gintDec) & Format(odynDet("SUBTOTAL").Value, "# ### ##0." & String(gintDec, "0")), 10 + gintDec)
                    Else
                        Printer.Print left("" & Space(6), 6) & Space(1) & _
                                left("" & Space(55), 55) & Space(1) & _
                                left("" & odynDet("NUM_LOTE").Value & Space(10), 10) & Space(1) & _
                                left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(1) & _
                                right("" & odynDet("CANTIDADES_LOTE").Value & Space(7), 7) & Space(1) & _
                                right(Space(6 + gintDec) & "", 6 + gintDec) & Space(18 - gintDec) & _
                                right(Space(10 + gintDec) & "", 10 + gintDec)

                    End If
                    strCodProdAnt = odynDet("COD_PRODUCTO").Value

                Else  'PARA EL DETALLE DE LA FACTURA EN CASO SEA FACTURA CORTA (FALTA REALIZAR)
                    
                    Printer.FontName = "Draft 12cpi"
                    'Printer.Print right(Space(7) & Trim(odynDet("CANT").Value), 7) & Space(2) & left(odynDet("COD_PRODUCTO").Value & " " & odynDet("OBS_DES_PRODUCTO").Value & Space(49), 49) & _
                     Space(1) & right(Space(6 + gintDec) & Format(odynDet("UNITARIO").Value, "# ##0." & String(gintDec, "0")), 6 + gintDec) & _
                     Space(15 - gintDec) & right(Space(10 + gintDec) & Format(odynDet("SUBTOTAL").Value, "# ### ##0." & String(gintDec, "0")), 10 + gintDec)
                End If
                odynDet.MoveNext
                intLinea = intLinea + 1
            Wend
            For i = intLinea To intNumLineas
                Printer.Print
            Next i
            Printer.Print
            
            sngCtv = Val(Replace(dblTotal, " ", ""))
            strImporteLetras = UCase(pfstr_Letra(Int(sngCtv), 80))
            strImporteResto = right(Format(sngCtv - Int(sngCtv), ".00"), 2)
            Printer.FontName = "Draft 17cpi"
            '5 Printer.Print "" & gclsOracle.FN_Valor("PKG_PEDIDO.FN_MSJ_PD", vstrCia, vstrTip_Documento, vstrNum_Documento)
            '6 Printer.Print "" & gclsOracle.FN_Valor("PKG_PEDIDO.FN_MSJ_BOTIQUIN", vstrCia, vstrTip_Documento, vstrNum_Documento)
            Printer.FontName = "Draft 12cpi"
            Printer.Print "     SON : " & Trim(strImporteLetras) & " Y " & strImporteResto & "/100 NUEVOS SOLES"
            Printer.Print

            If blnFacturaCorta = False Then 'PARA EL PIE DE PAGINA DE LA FACTURA EN CASO NO SEA FACTURA CORTA
                Printer.Print Space$(78 - gintDec) & "S/." & right(Space(10 + gintDec) & Format(dblSubTotal, "# ### ##0." & String(gintDec, "0")), 10 + gintDec)
                Printer.Print
                Printer.Print Space$(62 - gintDec) & "Igv" & Space(2) & Val(strIgv * 100) & "%" & Space(8) & "S/." & right(Space(10 + gintDec) & Format(dblIGV, "# ### ##0." & String(gintDec, "0")), 10 + gintDec)
                Printer.Print
                Printer.Print Space$(76 - gsintDecTot) & "S/." & right(Space(10 + gintDecTot) & Format(dblTotal, "# ### ##0." & String(gintDecTot, "0")), 10 + gintDecTot)
                Printer.Print
            Else 'PARA EL PIE DE PAGINA DE LA FACTURA EN CASO SEA FACTURA CORTA
                Printer.Print Space(13 - gintDec) & "S/." & right(Space(10 + gintDec) & Format(dblSubTotal, "# ### ##0." & String(gintDec, "0")), 10 + gintDec) & _
                              Space(16 - gintDec) & "Igv" & Space(2) & Val(strIgv * 100) & "%" & Space(5) & "S/." & right(Space(10 + gintDec) & Format(dblIGV, "# ### ##0." & String(gintDec, "0")), 10 + gintDec) & _
                              Space(15 - gintDec) & "S/." & right(Space(10 + gintDecTot) & Format(dblTotal, "# ### ##0.#0"), 12)

            End If
        Printer.EndDoc
    End Select
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un problema con la Impresora" & Chr(13) & _
             Err.Description & Chr(13) & _
             "Desea esperar a ser resuelto?", vbQuestion + vbYesNo) = vbYes Then
        Resume
    Else
        Exit Sub
    End If
End Sub

Function pfstr_Letra(ByVal strnum As String, Optional vLo) As String
Dim lngA As Long
Dim Negativo As Boolean
Dim l As Long
Dim Una As Boolean
Dim Millon As Boolean
Dim Millones As Boolean
Dim vez As Long
Dim MaxVez As Long
Dim k As Long
Dim strQ As String
Dim strB As String
Dim strU As String
Dim strD As String
Dim strC As String
Dim iA As Long
    
Dim strN() As String
Dim lo As Long

'Si no se especifica el ancho...
If IsMissing(vLo) Then
  lo = 255
Else
  lo = vLo
End If
        
Dim unidad(0 To 9) As String
Dim decena(0 To 9) As String
Dim centena(0 To 9) As String
Dim deci(0 To 9) As String
Dim otros(0 To 15) As String
'Asignar los valores
unidad(1) = "uno"          ' Dejar en minusculas
unidad(2) = "dos"
unidad(3) = "tres"
unidad(4) = "cuatro"
unidad(5) = "cinco"
unidad(6) = "seis"
unidad(7) = "siete"
unidad(8) = "ocho"
unidad(9) = "nueve"

decena(1) = "diez"
decena(2) = "veinte"
decena(3) = "treinta"
decena(4) = "cuarenta"
decena(5) = "cincuenta"
decena(6) = "sesenta"
decena(7) = "setenta"
decena(8) = "ochenta"
decena(9) = "noventa"
   
centena(1) = "ciento"
centena(2) = "doscientos"
centena(3) = "trescientos"
centena(4) = "cuatrocientos"
centena(5) = "quinientos"
centena(6) = "seiscientos"
centena(7) = "setecientos"
centena(8) = "ochocientos"
centena(9) = "novecientos"
    
deci(1) = "dieci"
deci(2) = "veinti"
deci(3) = "treinti"
deci(4) = "cuarenti"
deci(5) = "cincuenti"
deci(6) = "sesenti"
deci(7) = "setenti"
deci(8) = "ochenti"
deci(9) = "noventi"
    
otros(1) = "1"
otros(2) = "2"
otros(3) = "3"
otros(4) = "4"
otros(5) = "5"
otros(6) = "6"
otros(7) = "7"
otros(8) = "8"
otros(9) = "9"
otros(10) = "10"
otros(11) = "once"
otros(12) = "doce"
otros(13) = "trece"
otros(14) = "catorce"
otros(15) = "quince"

On Error GoTo 0
    
lngA = Abs(Val(strnum))
Negativo = (lngA <> Val(strnum))
strnum = LTrim$(RTrim$(str$(lngA)))
l = Len(strnum)
    
If lngA = 0 Then
  strnum = left$("cero" & Space$(lo), lo)
  pfstr_Letra = strnum
  Exit Function
End If
    '
Una = True
Millon = False
Millones = False
If l < 4 Then Una = False
  If lngA > 999999 Then Millon = True
  If lngA > 1999999 Then Millones = True
  strB = ""
  strQ = strnum
  vez = 0
    
  ReDim strN(1 To 4)
  strQ = right$(String$(12, "0") & strnum, 12)
  For k = Len(strQ) To 1 Step -3
    vez = vez + 1
    strN(vez) = Mid$(strQ, k - 2, 3)
  Next
  MaxVez = 4
  For k = 4 To 1 Step -1
    If strN(k) = "000" Then
      MaxVez = MaxVez - 1
    Else
      Exit For
    End If
  Next
  For vez = 1 To MaxVez
    strU = "": strD = "": strC = ""
    strnum = strN(vez)
    l = Len(strnum)
    k = Val(right$(strnum, 2))
    If right$(strnum, 1) = "0" Then
      k = k \ 10
      strD = decena(k)
      ElseIf k > 10 And k < 16 Then
        k = Val(Mid$(strnum, l - 1, 2))
        strD = otros(k)
      Else
        strU = unidad(Val(right$(strnum, 1)))
        If l - 1 > 0 Then
          k = Val(Mid$(strnum, l - 1, 1))
          strD = deci(k)
        End If
      End If
      If l - 2 > 0 Then
        k = Val(Mid$(strnum, l - 2, 1))
        strC = centena(k) & " "
      End If
      If strU = "uno" And left$(strB, 4) = " mil" Then strU = ""
        strB = strC & strD & strU & " " & strB
        ''If (vez = 1 Or vez = 3) And strN(vez + 1) <> "000" Then strB = " mil " & strB
        '-------------------------------------
      If (vez = 1 Or vez = 3) And strN(vez + 1) <> "000" And Len(CStr(lngA)) = 4 And Val(Mid(CStr(lngA), 1, 1)) = 1 Then
        strB = " un mil " & strB
      Else
        If (vez = 1 Or vez = 3) And strN(vez + 1) <> "000" Then strB = " mil " & strB
        End If
      '-------------------------------------
      If vez = 2 And Millon Then
        If Millones Then
          strB = " millones " & strB
        Else
          strB = "un millón " & strB
        End If
      End If
  Next
  strB = LTrim$(RTrim$(strB))
  If right$(strB, 3) = "uno" Then strB = left$(strB, Len(strB) - 1) & "o"
  Do                              'Quitar los espacios que haya por medio
   iA = InStr(strB, "  ")
   If iA = 0 Then Exit Do
   strB = left$(strB, iA - 1) & Mid$(strB, iA + 1)
  Loop
  If left$(strB, 6) = "uno un" Then strB = Mid$(strB, 5)
  If left$(strB, 7) = "uno mil" Then strB = Mid$(strB, 5)
  If right$(strB, 16) <> "millones mil uno" Then
    iA = InStr(strB, "millones mil uno")
    If iA Then strB = left$(strB, iA + 8) & Mid$(strB, iA + 13)
  End If
  If right$(strB, 6) = "ciento" Then strB = left$(strB, Len(strB) - 2)
  If Negativo Then strB = "menos " & strB
  
  strC = Space$(lo)
  LSet strC = strB
  pfstr_Letra = strC
End Function


Function pfint_SoloNumeros(ByRef rintKey%, Optional ByVal vstrSimbol$) As Integer
    Select Case rintKey
        Case 3, 22, 24, 13, 8, 9, 27, 48 To 57
        Case AscW(IIf(vstrSimbol = "", 0, vstrSimbol))
        Case Else: rintKey = 0
    End Select
    pfint_SoloNumeros = rintKey
End Function

Function pfstr_Val_Direccion_Local(ByVal vstrCod_Local$) As String
    pfstr_Val_Direccion_Local = gclsOracle.FN_Valor("btlprod.PKG_GESTION.FN_DIRECCION_LOCAL", vstrCod_Local)
End Function

Function pfstr_Val_Razon_Social(ByVal vstrCod_Cia$, ByVal vstrCod_Local$) As String
    pfstr_Val_Razon_Social = gclsOracle.FN_Valor("btlprod.PKG_GESTION.FN_RAZON_SOCIAL", vstrCod_Cia, vstrCod_Local)
End Function

