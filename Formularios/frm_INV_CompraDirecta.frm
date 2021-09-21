VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_INV_CompraDirecta 
   BorderStyle     =   0  'None
   Caption         =   "Busqueda de Compra Directa"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   615
      Left            =   6240
      Picture         =   "frm_INV_CompraDirecta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdRecepcion 
      Caption         =   "&Recepción"
      Height          =   615
      Left            =   1440
      Picture         =   "frm_INV_CompraDirecta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nueva Recepción"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   3360
      Picture         =   "frm_INV_CompraDirecta.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprimir"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "An&ular"
      Height          =   615
      Left            =   4680
      Picture         =   "frm_INV_CompraDirecta.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Anular"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompra 
      Caption         =   "&Compra"
      Height          =   615
      Left            =   120
      Picture         =   "frm_INV_CompraDirecta.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nueva Compra"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   6000
      Picture         =   "frm_INV_CompraDirecta.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   6480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFchIni 
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   17104899
      CurrentDate     =   38874
   End
   Begin MSComCtl2.DTPicker dtpFchFin 
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   17104899
      CurrentDate     =   38874
   End
   Begin vbp_Ventas.ctlGrilla grdDetalle 
      Height          =   4995
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8811
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlDataCombo ctlcboLocal 
      Height          =   315
      Left            =   1200
      TabIndex        =   15
      Top             =   480
      Width           =   2895
      _ExtentX        =   7011
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Local :"
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   16
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de O.Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Busqueda de Compra Directa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta :"
      Height          =   195
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde :"
      Height          =   195
      Index           =   1
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   555
   End
End
Attribute VB_Name = "frm_INV_CompraDirecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iLineTop As Integer
Private intLine As Integer

Private Sub cmdAnular_Click()

    Dim lstrError As String
    Dim lstrSecDocumento As String
    Dim lstrUsuario As String
    
    If grdDetalle.ApproxCount < 1 Then
        Exit Sub
    End If

    lstrSecDocumento = grdDetalle.DataSource("SEC_DOCUMENTO").Value
    lstrUsuario = objUsuario.Codigo

    If MsgBox("¿Desea Anular Registro de Compra?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Confirmación") = vbNo Then
        Exit Sub
    End If
    
    Dim objCompra As New clsCompra
    
    lstrError = objCompra.Anula(objUsuario.CodigoEmpresa, _
                                objUsuario.CodigoLocal, _
                                lstrSecDocumento, _
                                lstrUsuario, _
                                "0", "1")
    
    If lstrError <> "" Then
        MsgBox lstrError, vbCritical, "Aviso"
        Exit Sub
    End If
    
    MsgBox "REGISTRO ANULADO", vbExclamation, "Aviso"
    
    cmdBuscar_Click

End Sub

Private Sub cmdBuscar_Click()
    
On Error GoTo Control

    Dim lstrLocal As String
    Dim lstrFchIni As String
    Dim lstrFchFin As String
    
    lstrLocal = Trim(ctlcboLocal.BoundText)
    lstrFchIni = IIf(IsNull(dtpFchIni.Value), "", Format(dtpFchIni.Value, "dd/mm/yyyy"))
    lstrFchFin = IIf(IsNull(dtpFchFin.Value), "", Format(dtpFchFin.Value, "dd/mm/yyyy"))
    
    If lstrLocal = "" Or lstrFchIni = "" Or lstrFchFin = "" Then
        MsgBox "Son requeridos Local y Fechas", vbExclamation, "Aviso"
        Return
    End If
    
    Dim objCompra As New clsCompra
    Set grdDetalle.DataSource = objCompra.Lista(lstrLocal, lstrFchIni, lstrFchFin)
    Set objCompra = Nothing
    grdDetalle.SetFocus
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cmdCompra_Click()

    frm_INV_CompraDirectaReg.Mostrar "Compra"

End Sub

Private Sub cmdImprimir_Click()
On Error GoTo Control

    Dim lstrProveedor As String
    Dim lstrTipDocumento As String
    Dim lstrNumDocumento As String

    If grdDetalle.ApproxCount < 1 Then
        Exit Sub
    End If

    lstrProveedor = Trim(grdDetalle.DataSource("RUC").Value)
    lstrTipDocumento = Trim(grdDetalle.DataSource("TIP_DOC").Value)
    lstrNumDocumento = Trim(grdDetalle.DataSource("NUM_DOC").Value)
    
    If lstrProveedor = "" Or lstrTipDocumento = "" Or lstrNumDocumento = "" Then
        Exit Sub
    End If

    If MsgBox("¿Desea imprimir el Parte de Recepción?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Confirmación") = vbYes Then
        ImprimeParte lstrProveedor, _
                     lstrTipDocumento, _
                     lstrNumDocumento
    End If

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdRecepcion_Click()

    frm_INV_CompraDirectaReg.Mostrar "Recepcion"

End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        dtpFchIni.SetFocus
    End If
    If KeyCode = vbKeyF2 Then
        grdDetalle.SetFocus
    End If

End Sub

Private Sub Form_Load()

    SetteaFormulario Me
    
    iLineTop = 70

    Dim objLocal As New clsLocal
    Set ctlcboLocal.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
    ctlcboLocal.ListField = "LOCAL_DEX"
    ctlcboLocal.BoundColumn = "COD_LOCAL"
    ctlcboLocal.BoundText = objUsuario.CodigoLocal
    ctlcboLocal.Enabled = False
    Set objLocal = Nothing

    dtpFchIni.Value = CDate(Format(Now, "dd/mm/yyyy"))
    dtpFchFin.Value = CDate(Format(Now, "dd/mm/yyyy"))
    
    'Seteamos la grilla
    SetGrid

End Sub

Private Sub SetGrid()

    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("DES_TIPO", "NUM_OC", "DOCUMENTO", "FECHA", "PROVEEDOR", "NOMBRE", "COD_REGISTRO")
    arrCaption = Array("Tipo", "O.Compra", "Documento", "Fecha", "Proveedor", "Usuario", "Reg.Compra")
    arrAncho = Array(1400, 1100, 1500, 1000, 4000, 3000, 1200)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter)
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDetalle.Columns(0).FetchStyle = True
    grdDetalle.Columns(1).FetchStyle = True
    grdDetalle.Columns(2).FetchStyle = True
    grdDetalle.Columns(3).FetchStyle = True
    grdDetalle.Columns(4).FetchStyle = True
    grdDetalle.Columns(5).FetchStyle = True
    grdDetalle.Col = 0
    
End Sub

'********* RECEPCION
Private Sub spCabecera_Parte_Recep(ByRef rrstRep As oraDynaset)

    Dim blnReimpreso As Boolean
    
    blnReimpreso = False
    
    Printer.Print String(137, "-")
    Printer.FontName = "Draft 10cpi"
    Printer.FontBold = True
    Printer.Print fstr_Centrar("**--ORDEN DE COMPRA Nº " & CStr(rrstRep("NUM_ORDEN_COMPRA").Value) & "--**", 81)
    Printer.Print
    Printer.FontName = "Draft 12cpi"
    Printer.Print "PARTE DE RECEPCCION Nº " & CStr(rrstRep("SEC_DOCUMENTO").Value) & Space(28) & _
                  "FCH.RECEP : " & Format(rrstRep("FCH_RECEPCION").Value, "DD/MM/YYYY HH:MM:SS AM/PM")
    Printer.Print
    Printer.Print "PROVEEDOR : " & CStr(rrstRep("RUC_PROVEEDOR").Value) & " " & _
                 left(CStr(rrstRep("DES_PROVEEDOR").Value) & Space(60), 60)
    Printer.Print "TIPO DOC. : " & CStr(rrstRep("TIP_DOCUMENTO").Value) & Space(3) & _
                  left(rrstRep("DES_DOCUMENTO").Value & Space(24), 24) & Space(4) & _
                  "NUM. DOC. : " & CStr(rrstRep("NUM_DOCUMENTO").Value) & Space(4) & _
                  "FCH. DOC. : " & Format(rrstRep("FCH_EMISION").Value, "DD/MM/YYYY")
    Printer.Print "USUARIO   : " & IIf(IsNull(rrstRep("NOM_USUARIO").Value), "", rrstRep("NOM_USUARIO").Value)
    Printer.FontBold = False
    Printer.FontName = "Draft 17cpi"
    Printer.Print
    Printer.Print IIf(blnReimpreso, "RE-IMPRESION", "")
    Printer.Print String(137, "-")
    Printer.Print right(Space(3) & "#", 3) & Space(2) & _
                  left("CODIGO" & Space(5), 5) & Space(2) & _
                  right(Space(10) & "CTD.VERIF.", 10) & Space(2) & _
                  left("DESCRIPCION" & Space(60), 60) & Space(2) & _
                  left("UBICACION" & Space(10), 10)
    Printer.Print String(137, "=")
    intLine = intLine + 10
End Sub

Private Sub ImprimeParte(ByVal vstrRuc_Proveedor As String, _
                         ByVal vstrTip_Documento As String, _
                         ByVal vstrNum_Documento As String)
On Error GoTo Control

    Dim intCopia As Integer
    Dim rstRep As oraDynaset
    Dim rstRep2 As oraDynaset
    Dim rstRep3 As oraDynaset
    Dim p As Printer
    
    intCopia = 1
    
    For Each p In Printers
       If UCase(p.Port) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p

    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
        Dim objImpresion As New clsImpresion
     
        If Not objImpresion.cargaParametros("003") Then
            Exit Sub
        End If
    Else
        'Printer.Width = 1440 * 8.5
        'Printer.Height = 1440 * 8
        Printer.PaperSize = vbPRPSLetter
    End If
    
''    blnReimpreso = vblnRePrint
''    intFrmCargado = 0

''Inicia_Impresion:
''
''    If pfstr_Leer_Cadena_INI("Impresion", "ParteRep", istrArchivo) <> "1" Then Exit Sub
''
''    intCopia = Val(pfstr_Leer_Cadena_INI("Impresion", "NroCopParteRep", istrArchivo))
''
''    If Not fbln_SetImpresora(istrDevImpParteRep, "ImpParteRep") Then
'''        If MsgBox("No Tiene Impresora Seleccionada" & Chr(13) & "¿Desea Seleccionar Alguna?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'''            intFrmCargado = intFrmCargado + 1
'''            If intFrmCargado <> 1 Then Exit Sub
'''            frm_A0045.Show vbModal
'''            intFrmCargado = 0
'''            GoTo Inicia_Impresion
'''        End If
''        Exit Sub
''    End If
        
''    If intCopia = 0 Then Exit Sub

    Dim objCompra As New clsCompra
    
    Set rstRep = objCompra.ParteRecepcion(vstrRuc_Proveedor, vstrNum_Documento, vstrTip_Documento)
    Set rstRep2 = objCompra.ParteRecepcionDetalle(vstrRuc_Proveedor, vstrNum_Documento, vstrTip_Documento)
    Set rstRep3 = objCompra.ParteRecepcionDevolucion(vstrRuc_Proveedor, vstrNum_Documento, vstrTip_Documento)
    
    Set objCompra = Nothing
    
    If rstRep.EOF Then
        MsgBox "Error en ubicar el documento en la Base de Datos", vbExclamation, "Impresión"
        Exit Sub
    End If

    If rstRep2.EOF Then
        MsgBox "Error en ubicar el documento en la Base de Datos", vbExclamation, "Impresión"
        Exit Sub
    End If
    
Copia:
    rstRep.MoveFirst
    rstRep2.MoveFirst
    rstRep3.MoveFirst

On Error GoTo ErrorImpresora
    
    Printer.FontName = "Draft 17cpi"
    intLine = 0
    While Not rstRep2.EOF
        If Printer.Page <> 1 Then
            Printer.NewPage
            intLine = 0
        End If
        intLine = intLine + 1
        Call spCabecera_Parte_Recep(rstRep)

        Dim SUMA_codigos As Double
        Dim SUMA_cant_u As Double
        Dim NroItem%
        NroItem = 1
        SUMA_codigos = 0
        SUMA_cant_u = 0
        While intLine <= iLineTop And Not rstRep2.EOF
            Printer.Print right(Space(3) & CStr(NroItem), 3) & Space(2) & _
                          left(rstRep2("COD_PRODUCTO").Value & Space(5), 5) & Space(2) & _
                          right(Space(10) & rstRep2("CTD_PRODUCTO_ING").Value, 10) & Space(2) & _
                          left(rstRep2("DES_PRODUCTO").Value & Space(60), 60) & Space(2) & _
                          left(rstRep2("PIC").Value & Space(10), 10)

            SUMA_codigos = SUMA_codigos + Val(rstRep2.Fields("COD_PRODUCTO").Value)
            SUMA_cant_u = SUMA_cant_u + Val(rstRep2.Fields("CTD_PRODUCTO_ING").Value)
            
            rstRep2.MoveNext
            intLine = intLine + 1
            NroItem = NroItem + 1
        Wend
    Wend

    Printer.Print String(137, "=")
    Printer.Print
    Printer.Print "Suma de Còdigos = " & SUMA_codigos
    Printer.Print "Suma de Cantidad = " & SUMA_cant_u
    
    Dim cabe As Boolean
    cabe = True
    
    If Not rstRep3.RecordCount = 0 Then
        While Not rstRep3.EOF And intLine <= iLineTop
            If Not cabe Then
                Call spCabecera_Parte_Recep(rstRep)
                cabe = True
            End If
            Printer.Print fstr_Centrar("GENERO RECLAMO CON LA GUIA Nº    " & _
                          CStr(rstRep3.Fields("NUM_GUIA").Value), 137)
                          
            rstRep3.MoveNext
            intLine = intLine + 1
            While Not rstRep3.EOF And intLine <= iLineTop
                'Printer.Print fstr_Centrar(Space(21) & CStr(rstRep3.Fields("NUM_GUIA")), 137)
                Printer.Print pfstr_Alineado(Space(21) & CStr(rstRep3.Fields("NUM_GUIA")), 137, centro, "")
                rstRep3.MoveNext
                intLine = intLine + 1
            Wend
            If Not rstRep3.EOF Then
                cabe = False
                Printer.NewPage
                intLine = 1
            End If

        Wend
    End If

    Printer.EndDoc

   On Error GoTo 0

   GoTo final

ErrorImpresora:

    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
    End If
    
final:
    If intCopia > 1 Then intCopia = intCopia - 1: GoTo Copia
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    Exit Sub
End Sub

Private Sub grdDetalle_DblClick()

    Dim lstrTipo As String
    Dim lstrCaso As String
    Dim lstrProveedor As String
    Dim lstrTipDocumento As String
    Dim lstrNumDocumento As String

    If grdDetalle.ApproxCount < 1 Then
        Exit Sub
    End If

    lstrTipo = Trim(grdDetalle.DataSource("COD_TIPO").Value)
    lstrProveedor = Trim(grdDetalle.DataSource("RUC").Value)
    lstrTipDocumento = Trim(grdDetalle.DataSource("TIP_DOC").Value)
    lstrNumDocumento = Trim(grdDetalle.DataSource("NUM_DOC").Value)

    If lstrProveedor = "" Or lstrTipDocumento = "" Or lstrNumDocumento = "" Then
        Exit Sub
    End If
    
    If lstrTipo = "7" Or lstrTipo = "8" Then
        lstrCaso = "COMPRA"
    Else
        lstrCaso = "RECEPCION"
    End If

    frm_INV_CompraDirectaReg.Mostrar lstrCaso, _
                                     True, _
                                     lstrProveedor, _
                                     lstrTipDocumento, _
                                     lstrNumDocumento

End Sub
