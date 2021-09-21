VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'************Registro MArtinnets
'************Original Cristian Rueda-Setear Formulario en Impresora Predeterminada
' Constants for DEVMODE
Private Const CCHFORMNAME = 32
Private Const CCHDEVICENAME = 32
Private Const DM_FORMNAME As Long = &H10000
Private Const DM_ORIENTATION = &H1&
 
' Custom constants for this sample's SelectForm function
Private Const FORM_NOT_SELECTED = 0
Private Const FORM_SELECTED = 1
Private Const FORM_ADDED = 2
 
' Constants for DocumentProperties() call
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
 
Private Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type
 
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
 
Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As Any) As Long
Dim ldynFormato As oraDynaset
'************************************************************************

Public Function Lista(ByVal CIA As String, _
                      ByVal CodigoLocal As String, _
                      ByVal CodigoMaquinaDestino As String, _
                      Optional ByVal TipoDocumento As String = vbNullString, _
                      Optional ByVal conCheck As String = "0") As oraDynaset
                
On Error GoTo CnrlErr
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_IMPRESION.FN_LISTA", 0, CIA, CodigoLocal, CodigoMaquinaDestino, TipoDocumento, conCheck)
    Exit Function
CnrlErr:
    Err.Raise Err.Number, "clsImpresion", Err.Description
    
End Function



Public Function ListaPendiente(ByVal CIA As String, _
                ByVal CodigoLocal As String, _
                ByVal CodigoMaquinaDestino As String, _
                Optional CodigoTipoDoc As String _
                ) As Integer
                
On Error GoTo CnrlErr
    ListaPendiente = gclsOracle.FN_Valor("BTLPROD.PKG_IMPRESION.FN_COLA_PENDIENTE", CIA, CodigoLocal, CodigoMaquinaDestino, CodigoTipoDoc)
    Exit Function
CnrlErr:
    Err.Raise Err.Number, "clsImpresion", Err.Description
End Function




Public Function EliminaCola(CIA As String, _
                      CodigoLocal As String, _
                      TipoDocumento As String, _
                      NumeroDocumento As String, _
                      NombreMaquina As String) As String
                            
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO As Variant
        gvarValores = Array(CIA, CodigoLocal, TipoDocumento, NumeroDocumento, NombreMaquina)
        gvarIO = Array(entrada, entrada, entrada, entrada, entrada)
        EliminaCola = gclsOracle.SP("BTLPROD.PKG_IMPRESION.SP_ELIMINA", gvarValores, gvarIO)
        
Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsImpresion", Err.Description
End Function


Public Function PuedeImprimir(ByVal CIA As String, _
                    ByVal CodigoMaquina As String, _
                    ByVal TipoDocumento As String _
                ) As Boolean
                
On Error GoTo CnrlErr
    Dim flag As Integer
    flag = gclsOracle.FN_Valor("BTLPROD.PKG_IMPRESION.FN_TIPO_DOCUMENTO", CIA, CodigoMaquina, TipoDocumento)
    If flag = 0 Then
        PuedeImprimir = False
    Else
        PuedeImprimir = True
    End If
    
    Exit Function
CnrlErr:
    Err.Raise Err.Number, "clsImpresion", Err.Description
End Function



'-------------------------------------------------------------------------------------------------
'********************Funcione y Procedimientos Setear Formulario en Impresora
'-------------------------------------------------------------------------------------------------

Private Function SelectForm(ByVal vPrinterName As String, _
                           ByRef rFormName As String, _
                           ByVal vWidth As Long, _
                           ByVal vheight As Long, _
                           ByVal vMyhWnd As Long) _
    As Integer
Dim nSize As Long           ' Size of DEVMODE
Dim pDevMode As DEVMODE
Dim PrinterHandle As Long   ' Handle to printer
Dim hPrtDC As Long          ' Handle to Printer DC
'Dim PrinterName As String
Dim aDevMode() As Byte      ' Working DEVMODE
Dim FormSize As SIZEL
Dim FormNameNew As String
'PrinterName = Printer.DeviceName  ' Current printer
hPrtDC = Printer.hdc              ' hDC for current Printer
SelectForm = FORM_NOT_SELECTED    ' Set for failure unless reset in code.
 
FormNameNew = rFormName
 
' Get a handle to the printer.
If OpenPrinter(vPrinterName, PrinterHandle, 0&) Then
    ' Retrieve the size of the DEVMODE.
    nSize = DocumentProperties(vMyhWnd, PrinterHandle, vPrinterName, 0&, _
            0&, 0&)
    ' Reserve memory for the actual size of the DEVMODE.
    ReDim aDevMode(1 To nSize)
 
    ' Fill the DEVMODE from the printer.
    nSize = DocumentProperties(vMyhWnd, PrinterHandle, vPrinterName, _
            aDevMode(1), 0&, DM_OUT_BUFFER)
    ' Copy the Public (predefined) portion of the DEVMODE.
    Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
    
    ' If FormName is "MyCustomForm", we must make sure it exists
    ' before using it. Otherwise, it came from our EnumForms list,
    ' and we do not need to check first. Note that we could have
    ' passed in a Flag instead of checking for a literal name.
    'If FormName = "MyCustomForm" Then
        ' Use form "MyCustomForm", adding it if necessary.
        ' Set the desired size of the form needed.
        Dim lngDel As Long
        lngDel = DelForm(vPrinterName, rFormName)
        
        'Elimino antes si es que existe el formulario
        If Not (lngDel = 0 Or lngDel = 87 Or lngDel = 1902) Then
            SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
            ClosePrinter (PrinterHandle)
            Exit Function
        End If
        With FormSize   ' Given in thousandths of millimeters
            .cx = vWidth  ' width
            .cy = vheight ' height
        End With
        If GetFormName(PrinterHandle, FormSize, rFormName) = 0 Then
            ' Form not found - Either of the next 2 lines will work.
            'FormName = AddNewForm(PrinterHandle, FormSize, "MyCustomForm")
            'AddNewForm PrinterHandle, FormSize, "MyCustomForm"
            AddNewForm PrinterHandle, FormSize, FormNameNew
            If GetFormName(PrinterHandle, FormSize, rFormName) = 0 Then
                ClosePrinter (PrinterHandle)
                SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
                Exit Function
            Else
                SelectForm = FORM_ADDED  ' Form Added, Selection succeeded!
            End If
        End If
    'End If
    
    ' Change the appropriate member in the DevMode.
    ' In this case, you want to change the form name.
    pDevMode.dmFormName = rFormName & Chr(0)  ' Must be NULL terminated!
    ' Set the dmFields bit flag to indicate what you are changing.
    pDevMode.dmFields = DM_FORMNAME
 
    ' Copy your changes back, then update DEVMODE.
    Call CopyMemory(aDevMode(1), pDevMode, Len(pDevMode))
    nSize = DocumentProperties(vMyhWnd, PrinterHandle, vPrinterName, _
            aDevMode(1), aDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
 
    nSize = ResetDC(hPrtDC, aDevMode(1))   ' Reset the DEVMODE for the DC.
 
    ' Close the handle when you are finished with it.
    ClosePrinter (PrinterHandle)
    ' Selection Succeeded! But was Form Added?
    If SelectForm <> FORM_ADDED Then SelectForm = FORM_SELECTED
Else
    SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
End If
End Function
'-------------------------------------------------------------------------------------------------
Private Function DelForm(ByVal vPrinterName As String, _
                        ByVal vFormName As String) As Long
Dim PrinterHandle As Long   ' Handle to printer
Dim RetVal As Long
 
    If OpenPrinter(vPrinterName, PrinterHandle, 0&) Then
        RetVal = DeleteForm(PrinterHandle, vFormName & Chr(0))
        If RetVal <> 0 Then ' DeleteForm succeeded.
            DelForm = 0
        Else
            DelForm = Err.LastDllError
        End If
        ClosePrinter (PrinterHandle)
    End If
End Function
'-------------------------------------------------------------------------------------------------
Private Function pfbln_SetForm(ByVal vstrCod_formato As String) As Boolean
'Dim objFormato As New clsFormato
Dim intTemp As Integer
pfbln_SetForm = False
       
    intTemp = SelectForm(Printer.Devicename, _
                         Nombre(vstrCod_formato), _
                         Ancho(vstrCod_formato) * 1000, _
                         alto(vstrCod_formato) * 1000, _
                         App.hInstance)
    'intTemp = objFormato.sub_carga(vstrCod_formato)
    'Set objFormato = Nothing
    If intTemp = 0 Then
        MsgBox "Unable to retrieve From name", vbExclamation, "Operation halted!"
        Exit Function
    End If
 
    pfbln_SetForm = True
End Function
Function cargaParametros(ByVal vstrCod_formato As String) As Boolean
    'Set lodynDatos = gclsOracle.FN_Cursor("CMR.PKG_FORMATO.FN_LISTA", 0, vstrCod_formato)
    
    'Dim OraSqlStmt As OraSqlStmt
    On Error GoTo Error
    godbVentas.Parameters.Remove "P_COD_FORMATO"
    godbVentas.Parameters.Add "P_COD_FORMATO", vstrCod_formato, ORAPARM_INPUT
    Set ldynFormato = godbVentas.CreatePlsqlDynaset("BEGIN :SALIDA := CMR.PKG_FORMATO.FN_LISTA(:P_COD_FORMATO);END;", "SALIDA", ORADYN_READONLY)
    cargaParametros = pfbln_SetForm(vstrCod_formato)
    Exit Function
Error:
    Err.Raise Err.Number, "clsFormato", Err.Description
End Function

Private Property Get Nombre(ByVal vstrCod_formato As String) As String
   ldynFormato.FindFirst " COD_FORMATO = '" & vstrCod_formato & "'"
   Nombre = ldynFormato("DES_FORMATO").Value
End Property

Private Property Get Ancho(ByVal vstrCod_formato As String) As Double
   ldynFormato.FindFirst " COD_FORMATO = '" & vstrCod_formato & "'"
   Ancho = ldynFormato("CTD_ANCHO").Value
End Property

Private Property Get alto(ByVal vstrCod_formato As String) As Double
   ldynFormato.FindFirst " COD_FORMATO = '" & vstrCod_formato & "'"
   alto = ldynFormato("CTD_ALTO").Value
End Property
'-------------------------------------------------------------------------------------------------
'*****************************************************************
'-------------------------------------------------------------------------------------------------





'-------------------------------------------------------------------------------------------------
'*********************** Añadido para contingencia de ventas
'*********************** jlopez
'*********************** 9/5/08
'-------------------------------------------------------------------------------------------------


Public Function pfbln_SetFormOFF(ByVal pstrNombre As String, ByVal pdblAncho As Double, ByVal pdblAlto As Double) As Boolean
Dim intTemp As Integer
pfbln_SetFormOFF = False
       
    intTemp = SelectForm(Printer.Devicename, _
                         pstrNombre, _
                         pdblAncho * 1000, _
                         pdblAlto * 1000, _
                         App.hInstance)
    If intTemp = 0 Then
        MsgBox "Unable to retrieve From name", vbExclamation, "Operation halted!"
        Exit Function
    End If
 
    pfbln_SetFormOFF = True
End Function





