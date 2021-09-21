VERSION 5.00
Begin VB.Form frmColaImpresion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresiones Pendientes"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrillaArray grdDocumentos 
      Height          =   5295
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9340
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5415
      Begin vbp_Ventas.ctlDataCombo cboCodTipoDoc 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   260
         Width           =   1230
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1058
   End
End
Attribute VB_Name = "frmColaImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objImpresion As New clsImpresion

Private Sub cboCodTipoDoc_Click(Area As Integer)
    Call Consulta
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case boton
        Case 6
            Call Imprimir
        Case 13
            Call Eliminar
        Case 15
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Dim oDoc As New clsDocumento
    
    With cboCodTipoDoc
        Set .RowSource = oDoc.ListaTipo(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
        .ListField = "DESCRIPCION"
        .BoundColumn = "CODIGO"
        .BoundText = "*"
    End With
    
    SeteaToolBar
    SeteaGrila
    Consulta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objImpresion = Nothing
End Sub

Sub Consulta()
    Dim objResult As oraDynaset
    Dim arrTempC As New XArrayDB
        
    Set objResult = objImpresion.Lista(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, _
                                       objUsuario.NombrePC, IIf(cboCodTipoDoc.BoundText <> "*", cboCodTipoDoc.BoundText, vbNullString), "1")

    If objResult.RecordCount > 0 Then
        arrTempC.LoadRows objResult.GetRows
    Else
        grdDocumentos.Limpiar
        arrTempC.ReDim 0, -1, 0, -1
    End If
    
    grdDocumentos.Array1 = arrTempC
    grdDocumentos.Rebind
    
    Set objResult = Nothing
End Sub

Sub SeteaGrila()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant
    
    arrCampos = Array("MARCA", "CIA", "COD_LOCAL", "COD_TIPO_DOCUMENTO", _
                      "NUM_DOCUMENTO", "NUM_COPIAS", "NUM_IMPRESAS", "COD_MAQUINA_ORIGEN", _
                      "COD_USUARIO_ORIGEN", "COD_MAQUINA_DESTINO", "COD_FORMATO", "COD_MODALIDAD_VENTA")
                      
    arrCaption = Array("S", "", "Local", "TD", _
                       "N° Documento", "", "", "Maquina", _
                       "", "", "COD_FORMATO", "Modalidad")
                       
    arrAncho = Array(400, 0, 650, 650, _
                     1700, 0, 0, 1700, _
                     0, 0, 0, 900)
                     
    arrAlineacion = Array(dbgCenter, dbgCenter, vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, vbAlignLeft, dbgCenter)
                          
    arrFoco = Array(True, False, False, False, _
                    False, False, False, False, _
                    False, False, False, False)
    
    grdDocumentos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    
    grdDocumentos.Columns(0).ValueItems.Presentation = dbgCheckBox
    grdDocumentos.Columns(0).Visible = True
    grdDocumentos.Columns(1).Visible = False
    grdDocumentos.Columns(2).Visible = True
    grdDocumentos.Columns(3).Visible = True
    grdDocumentos.Columns(4).Visible = True
    grdDocumentos.Columns(5).Visible = False
    grdDocumentos.Columns(6).Visible = False
    grdDocumentos.Columns(7).Visible = True
    grdDocumentos.Columns(8).Visible = False
    grdDocumentos.Columns(9).Visible = False
    grdDocumentos.Columns(10).Visible = False
    grdDocumentos.AllowUpdate = True

End Sub

Private Sub SeteaToolBar()
    Dim i As Byte
    
    With ctlToolBar1
        For i = 1 To .Buttons.Count
            .Buttons(i).Visible = False
        Next
    
        .Buttons(6).Visible = True
        .Buttons(13).Visible = True
        .Buttons(15).Visible = True
    End With
End Sub

Private Sub Imprimir() 'grdDocumentos_DblClick()
    Dim i As Integer
    Dim k As Integer
    Dim X As Integer
    Dim arrTempI As New XArrayDB
    Dim objImpresion As New clsDocumento
    Dim objImprime As New clsImpresion
    Dim strMensaje As String
    
    On Error GoTo handle

    If grdDocumentos.ApproxCount <= 0 Then Exit Sub
    grdDocumentos.MoveNext
    grdDocumentos.MovePrevious
    arrTempI.ReDim 0, -1, 0, -1
    k = 0
    
    With grdDocumentos.Array1
        For i = .LowerBound(1) To .UpperBound(1)
            If Abs(Val(.Value(i, 0))) = 1 Then
                arrTempI.ReDim 0, k, 0, .UpperBound(2)
                For X = .LowerBound(2) To .UpperBound(2)
                    arrTempI(k, X) = .Value(i, X)
                Next X
                k = k + 1
            End If
        Next i
    End With

    If arrTempI.Count(1) <> 1 Then
        MsgBox "Debe seleccionar solamente un documento para imprimir.", vbCritical + vbOKOnly, "Imprimir"
        Exit Sub
    End If
    
    Call ImprimeCupon("" & arrTempI.Value(0, 3), "" & arrTempI.Value(0, 4))
    
    'objImpresion.ImprimirDocumento grdDocumentos.Columns("COD_TIPO_DOCUMENTO").Value, grdDocumentos.Columns("NUM_DOCUMENTO").Value, grdDocumentos.Columns("COD_FORMATO").Value
    objImpresion.ImprimirDocumento "" & arrTempI.Value(0, 3), "" & arrTempI.Value(0, 4), "" & arrTempI.Value(0, 10), "" & arrTempI.Value(0, 11)
    
    'strMensaje = objImprime.EliminaCola(objUsuario.CodigoEmpresa, grdDocumentos.Columns("COD_LOCAL").Value, grdDocumentos.Columns("COD_TIPO_DOCUMENTO").Value, grdDocumentos.Columns("NUM_DOCUMENTO").Value, objUsuario.NombrePC)
    strMensaje = objImprime.EliminaCola(objUsuario.CodigoEmpresa, "" & arrTempI.Value(0, 2), "" & arrTempI.Value(0, 3), "" & arrTempI.Value(0, 4), objUsuario.NombrePC)
    
    If Not strMensaje = "" Then
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    ''objDocumento.ImprimirGuiaTransferencia strNumDocumento
    Consulta
    Set objImpresion = Nothing
    Set objImprime = Nothing
    Exit Sub
handle:
    Set objImpresion = Nothing
    Set objImprime = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Eliminar() 'grdDocumentos_DblClick()
    Dim i As Integer
    Dim k As Integer
    Dim X As Integer
    Dim arrTempE As New XArrayDB
    Dim objImprime As New clsImpresion
    Dim strMensaje As String
    
    On Error GoTo handle

    If grdDocumentos.ApproxCount <= 0 Then Exit Sub
    grdDocumentos.MoveNext
    grdDocumentos.MovePrevious
    arrTempE.ReDim 0, -1, 0, -1
    k = 0
    
    With grdDocumentos.Array1
        For i = .LowerBound(1) To .UpperBound(1)
            If Abs(Val(.Value(i, 0))) = 1 Then
                arrTempE.ReDim 0, k, 0, .UpperBound(2)
                For X = .LowerBound(2) To .UpperBound(2)
                    arrTempE(k, X) = .Value(i, X)
                Next X
                k = k + 1
            End If
        Next i
    End With

    If arrTempE.Count(1) = 0 Then
        MsgBox "No ha seleccionado documentos para eliminar.", vbCritical + vbOKOnly, "Imprimir"
        Exit Sub
    End If
    
    With arrTempE
        For i = .LowerBound(1) To .UpperBound(1)
            strMensaje = objImprime.EliminaCola(objUsuario.CodigoEmpresa, "" & arrTempE.Value(i, 2), "" & arrTempE.Value(i, 3), "" & arrTempE.Value(i, 4), objUsuario.NombrePC)
            If Not strMensaje = "" Then
                MsgBox strMensaje, vbCritical, App.ProductName
            End If
        Next i
    End With
    
    Call Consulta
    Set objImprime = Nothing
    Exit Sub
handle:
    Set objImprime = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub ImprimeCupon(tpDocumento As String, numDocumento As String)
    Dim objDocumento As New clsDocumento
    Dim strNombreCuponera As String
    Dim Impresoras As Printer
    Dim bolValor As Boolean
    Dim strSecVentas As String
    Dim rsCupones As oraDynaset
    Dim Indicador As Boolean
    
    Indicador = False
        
    strNombreCuponera = objVenta.NombreCuponera(objUsuario.NombrePC)
    strSecVentas = objVenta.fnDevuelveSecuencia(objUsuario.CodigoEmpresa, tpDocumento, numDocumento)
    
    
    bolValor = False
    For Each Impresoras In Printers
        If UCase(strNombreCuponera) = UCase(Impresoras.Devicename) Then
           Set Printer = Impresoras
           bolValor = True
: Exit For
        End If
    Next
         
    
    
    If bolValor Then
       Set rsCupones = objDocumento.ListaCupones(strSecVentas)
       Dim cantidad As Integer
       cantidad = rsCupones.RecordCount
       
       If cantidad > 0 Then
            MsgBox "Se procederá a generar Cupón de Descuento." & Chr(13) & "Prepare su impresora.", vbInformation, App.ProductName
            mdiPrincipal.fnImprimeCupon "" & rsCupones("COD_DOCUMENTO_PAGO"), "" & rsCupones("NUM_DOCUMENTO_PAGO")
       End If
    Else
       MsgBox "No tiene la impresora de cupones Instalada o tiene diferente nombre", vbCritical, App.ProductName
       Exit Sub
    End If
End Sub

Private Sub grdDocumentos_DblClickRegistro(ByVal DatoColumna0 As String)
    grdDocumentos.Columns(0).Value = -1
    'Call ImprimeCupon
    Call Imprimir
End Sub
