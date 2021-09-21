VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_ADM_PedEspecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Especiales"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Estado"
      Height          =   735
      Left            =   8160
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
      Begin vbp_Ventas.ctlDataCombo cboEstado 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         MatchEntry      =   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Emitidas"
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   3855
      Begin MSComCtl2.DTPicker dtpFchIni 
         Height          =   315
         Left            =   480
         TabIndex        =   11
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   66977795
         CurrentDate     =   37950
      End
      Begin MSComCtl2.DTPicker dtpFchFin 
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   66977795
         CurrentDate     =   37950
      End
      Begin VB.Label Label7 
         Caption         =   "&Al"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "De&l"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "&Producto"
      Height          =   735
      Left            =   6120
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
      Begin vbp_Ventas.ctlTextBox txtCod_Producto 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Tipo            =   8
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Número de Pedido"
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin vbp_Ventas.ctlTextBox txtNumero 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10575
      Begin vbp_Ventas.ctlGrilla grdDetalle 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   3360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4895
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlGrilla grdCabecera 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4895
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle del Pedido:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblPedEspecial 
         Caption         =   "2008050001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1058
      ModoBotones     =   8
      EnabledEfecto   =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_PedEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrParam As Variant
Private strCodLocal As String
Public FLAG As Integer

Public Property Get CodLocal() As String
    CodLocal = strCodLocal
End Property

Public Property Let CodLocal(ByVal vstrNewValue As String)
    strCodLocal = vstrNewValue
End Property

Private Sub cmdBusProducto_Click()
    frm_ADM_BuscaProd.Show vbModal
    txtCod_Producto.Text = frm_ADM_BuscaProd.strCodProducto
End Sub

Private Sub sub_MM(Optional ByVal vstrNumero As String = "")
Dim strNumero As String

    If grdCabecera.ApproxCount > 0 Then

        If vstrNumero = "" Then
            strNumero = grdCabecera.Columns("NUM_PEDESPECIAL").Value
        Else
            strNumero = vstrNumero
        End If

        sub_Buscar arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4), arrParam(5)

        grdCabecera.DataSource.FindFirst " NUM_PEDESPECIAL = '" & strNumero & "'"

        If grdCabecera.DataSource.NoMatch Then
            grdCabecera.MoveFirst
        End If

    End If

End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    On Error GoTo CtrlErr
    Select Case boton
        Case Nuevo
            FLAG = 1
            sub_Nuevo
        Case Modificar
            If grdCabecera.Columns("COD_ESTADO").Value = "EMI" Then
               sub_Actualizar grdCabecera.Columns("NUM_PEDESPECIAL").Value
            Else
                MsgBox "Para editar, el Estado debe ser EMITIDO", vbExclamation, "Atención"
            End If
        Case Buscar
            spBuscar
            
        Case tb_Actualizar
            sub_MM
        Case Imprimir
            If grdCabecera.ApproxCount > 0 Then
                Me.grdCabecera.MostrarImprimir
            Else
                MsgBox "No se encontraron Pedidos para Imprimir.", vbCritical + vbInformation, "Aviso"
            End If
        Case tb_Excel
            If grdCabecera.ApproxCount > 0 Then
                Me.grdCabecera.MostrarExcel
            Else
                MsgBox "No se encontraron Pedidos para Exportar.", vbCritical + vbInformation, "Aviso"
            End If
        Case tb_email
            If grdCabecera.ApproxCount > 0 Then
                Me.grdCabecera.MostrarEmail
            Else
                MsgBox "No se encontraron Pedidos para Adjuntar.", vbCritical + vbInformation, "Aviso"
            End If
        Case Grabar
            
        Case Cancelar
            
        Case Eliminar
            If grdCabecera.Columns("COD_ESTADO").Value = "EMI" Then
               sub_Anular
            Else
                MsgBox "Para anular, el Estado debe ser EMITIDO", vbExclamation, "Atención"
            End If
        Case salir
            Unload Me
    End Select
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub spBuscar()
If txtNumero.Text = "" And dtpFchIni.Value > dtpFchFin.Value Then
    MsgBox "La Fecha ""DE"" no puede ser mayor a la fecha ""AL"" ", vbExclamation, "Atención"
                dtpFchIni.SetFocus
Else
    sub_Buscar CodLocal, dtpFchIni.Value, dtpFchFin.Value, _
                            txtNumero.Text, cboEstado.BoundText, txtCod_Producto.Text
End If
End Sub

Private Sub dtpFchFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpFchIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If grdCabecera.Columns("COD_ESTADO").Value = "EMI" Then
           sub_Actualizar grdCabecera.Columns("NUM_PEDESPECIAL").Value
        Else
           MsgBox "Para editar, el Estado debe ser EMITIDO", vbExclamation, "Atención"
        End If
    End If
End Sub

Public Sub Form_Load()
    
    'Cargando el DataComboBox de estado
    Dim objEstado As New clsEstadoSMM
    Set cboEstado.RowSource = objEstado.Lista
    Set objEstado = Nothing
    cboEstado.BoundColumn = "COD"
    cboEstado.ListField = "DES"
    cboEstado.BoundText = "*"
    
    SeteaGrilla
    
    dtpFchFin.Value = gclsOracle.Fecha_Servidor
    dtpFchIni.Value = dtpFchFin.Value - 30
    
    lblPedEspecial.Caption = ""
    
    'ctlToolBar1.Buttons(3).Visible = False
    ctlToolBar1.Buttons(4).Visible = False
    ctlToolBar1.Buttons(5).Visible = False
    ctlToolBar1.Buttons(6).Visible = False
    ctlToolBar1.Buttons(7).Visible = False

    
    sub_Buscar CodLocal, dtpFchIni.Value, dtpFchFin.Value, _
                        txtNumero.Text, cboEstado.BoundText, txtCod_Producto.Text, False
        
End Sub





Private Sub grdCabecera_DblClick()
'If grdCabecera.Columns("COD_ESTADO").Value = "EMI" Then
'  sub_Actualizar grdCabecera.Columns("NUM_PEDESPECIAL").Value
'Else
'  MsgBox "Para editar, el Estado debe ser EMITIDO", vbExclamation, "Atención"
'End If
End Sub

Private Sub grdCabecera_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
'    Select Case Col
'        Case grdCabecera.Columns("COD_ESTADO").ColIndex
'            Select Case grdCabecera.Columns("COD_ESTADO").CellValue(Bookmark)
'                Case "EMI"
'                    CellStyle.BackColor = RGB(50, 175, 50)
'                    CellStyle.ForeColor = vbWhite
'                Case "ATE"
'                    CellStyle.BackColor = RGB(50, 50, 175)
'                    CellStyle.ForeColor = vbWhite
'                Case "ANU"
'                    CellStyle.BackColor = RGB(175, 50, 50)
'                    CellStyle.ForeColor = vbWhite
'            End Select
'
'    End Select
    
End Sub

Private Sub grdCabecera_RegistroSeleccionado(ByVal DatoColumna0 As String)
    Dim objSPVM As New clsPedEspecial

    On Error GoTo CtrlErr

    grdDetalle.Limpiar
''''    Dim objEstadistica  As New clsEstadistica
''''    objEstadistica.Limpia
''''    Set objEstadistica = Nothing

    If grdCabecera.ApproxCount > 0 Then
        Set grdDetalle.DataSource = objSPVM.ListaDet(CodLocal, grdCabecera.Columns("NUM_PEDESPECIAL").Value)
        Set objSPVM = Nothing

        If arrParam(5) <> "" Then
            grdDetalle.DataSource.FindFirst " COD_PRODUCTO = '" & arrParam(5) & "'"
            If grdDetalle.DataSource.NoMatch Then
                grdDetalle.MoveFirst
            End If
        End If

    End If

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error Carga Detalle"
End Sub


Private Sub grdDetalle_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    
    On Error GoTo CtrlErr
        
    Select Case grdDetalle.Columns(Col).DataField
        Case "DES_PRODUCTO"
            If grdDetalle.Columns("FLG_SELECCIONADO").CellValue(Bookmark) = "1" Then
                Select Case Condition
                    Case CellStyleConstants.dbgNormalCell
                        CellStyle.ForeColor = vbRed
                        CellStyle.Font.Bold = True
                    Case CellStyleConstants.dbgMarqueeRow, CellStyleConstants.dbgMarqueeRow + CellStyleConstants.dbgCurrentCell
                        CellStyle.ForeColor = vbYellow
                        CellStyle.Font.Bold = True
                End Select
            End If
            
        Case "CTD_MAXIMO_APROBADO", "MOTIVO", "FCH_INICIAL_VIGENCIA", "FCH_FINAL_VIGENCIA"
            If grdDetalle.Columns("FLG_PROCESADO").CellValue(Bookmark) = "1" Then
                Select Case Condition
                    Case CellStyleConstants.dbgNormalCell
                        CellStyle.ForeColor = vbBlue
                        CellStyle.Font.Bold = True
                    Case CellStyleConstants.dbgMarqueeRow, CellStyleConstants.dbgMarqueeRow + CellStyleConstants.dbgCurrentCell
                        CellStyle.ForeColor = vbCyan
                        CellStyle.Font.Bold = True
                End Select
            End If
        End Select
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub txtCod_Producto_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
    
    If KeyAscii = vbKeyReturn Then
        
        If Len(Trim(txtCod_Producto.Text)) < 3 Then
            MsgBox "use por lo menos 3 caracteres", vbExclamation, "Aviso"
            Exit Sub
        End If
        
        Dim frm As New frm_ADM_ProductoDatos
        frm.Dato = Trim(txtCod_Producto.Text)
        frm.Show vbModal
        
        If frm.Salida(1) <> "" Then
            txtCod_Producto.Text = frm.Salida(1)
        End If
                
                    
        Set frm = Nothing
        
    End If

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub sub_Nuevo()
If FLAG = 1 Then
Dim strNumero As String
Dim frm As New frm_ADM_PedEspecialNuevo
    
    frm.solicitud.CodLocal = CodLocal
    frm.v_accion = 1
    frm.Show vbModal
    strNumero = frm.solicitud.Numero
    
    If strNumero <> "" Then
        If grdCabecera.ApproxCount > 0 Then
            sub_MM strNumero
        Else
            sub_Buscar CodLocal, dtpFchIni.Value, dtpFchFin.Value, _
                        txtNumero.Text, cboEstado.BoundText, txtCod_Producto.Text
                        
            grdCabecera.DataSource.FindFirst " NUM_PEDESPECIAL = '" & strNumero & "'"

            If grdCabecera.DataSource.NoMatch Then
                grdCabecera.MoveFirst
            End If
        End If
    End If
    frm.Caption = "Nuevo Pedido Especial"
    Set frm = Nothing
End If
End Sub

Private Sub sub_Buscar(ByVal vstrCodLocal As String, _
                       ByVal vfecDesde As Date, _
                       ByVal vfecHasta As Date, _
                       ByVal vstrNumero As String, _
                       ByVal vstrEstado As String, _
                       ByVal vstrCodProducto As String, _
                       Optional ByVal vblnMsgShow As Boolean = True)
    
    Dim objSPVM As New clsPedEspecial
        
    
    Set grdCabecera.DataSource = objSPVM.Lista(Format(vfecDesde, "DD/MM/YYYY"), _
                                              Format(vfecHasta, "DD/MM/YYYY"), _
                                              vstrCodLocal, _
                                              vstrEstado, _
                                              vstrNumero, _
                                              vstrCodProducto)
    Set objSPVM = Nothing
    
    arrParam = Array(vstrCodLocal, _
                     vfecDesde, _
                     vfecHasta, _
                     vstrNumero, _
                     vstrEstado, _
                     vstrCodProducto)
    
    If grdCabecera.ApproxCount < 1 Then
        grdDetalle.Limpiar
        If vblnMsgShow Then
            MsgBox "Búsqueda sin Resultados", vbInformation, "Aviso"
        End If
    End If
    
End Sub

Private Sub sub_Actualizar(Optional ByVal vstrNumero As String = "")
Dim strNumero As String
Dim objPedido As New clsPedEspecial
Dim frm As New frm_ADM_PedEspecialNuevo
frm.solicitud.CodLocal = CodLocal
strNumero = frm.solicitud.Numero
    
    If grdCabecera.ApproxCount > 0 Then
        If vstrNumero = "" Then
            strNumero = grdCabecera.Columns("NUM_PEDESPECIAL").Value
        Else
            strNumero = vstrNumero
        End If
        
        Dim ODyn As oraDynaset
        Set ODyn = objPedido.ListaDet(CodLocal, strNumero)
        frm.grdDetalle.Limpiar
        frm.SeteaGrilla
        frm.AdicionaDetalle ODyn
        frm.v_accion = 2
        frm.v_numPedido = strNumero
    Else
        
    End If
    frm.Caption = "Editar Pedido Especial"
    frm.Show vbModal
    Set frm = Nothing
End Sub

Private Sub sub_Anular()
Dim objSPVM As New clsPedEspecial
Dim strNumero As String
Dim strError As String

    If grdCabecera.ApproxCount > 0 Then
        If grdCabecera.Columns("COD_ESTADO") = "EMI" Then
            strNumero = grdCabecera.Columns("NUM_PEDESPECIAL").Value
            If MsgBox("¿Desea Anular el pedido ->" & strNumero & "<- ?", vbQuestion + vbYesNo + vbDefaultButton2, "Anular") = vbYes Then
                objSPVM.Anular CodLocal, strNumero, objUsuario.codigo
                MsgBox "La Solicitud se anuló con exito", vbInformation, "Anular"
                
                sub_Buscar CodLocal, dtpFchIni.Value, dtpFchFin.Value, _
                            txtNumero.Text, cboEstado.BoundText, txtCod_Producto.Text
                            
            End If
        Else
            MsgBox "La Solicitud No esta en estado Emitida", vbExclamation, "Anular"
        End If
    End If
    
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim i As Byte
  
    '---------------------------------------------------------------
    '-- Cabecera
    '---------------------------------------------------------------
    
    arrCampos = Array("NUM_PEDESPECIAL", "FCH_EMISION", "NUM_ITEMS", "USUARIO_EMISOR", _
                      "COD_ESTADO", "USUARIO_REVISA", "FCH_REVISION", _
                      "USUARIO_ANULA", "FCH_ANULACION")
                      
    arrCaption = Array("Número", "F.Emi.", "Items", "Emite", _
                       "Estado", "Revisa", "F.Rev.", _
                       "Anula", "F.Anu.")
    arrAncho = Array(1500, 1200, 600, 1800, _
                     600, 1800, 1200, _
                     1800, 1200)
    
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgLeft, _
                          dbgCenter, dbgLeft, dbgCenter, _
                          dbgLeft, dbgCenter)
    
    grdCabecera.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdCabecera.Columns("COD_ESTADO").FetchStyle = True
    grdCabecera.Columns("NUM_PEDESPECIAL").FetchStyle = True
    grdCabecera.Columns("FCH_EMISION").NumberFormat = "DD/MM/YY"
    grdCabecera.Columns("FCH_REVISION").NumberFormat = "DD/MM/YY"
    grdCabecera.Columns("FCH_ANULACION").NumberFormat = "DD/MM/YY"
    
    '---------------------------------------------------------------
    '-- Detalle
    '---------------------------------------------------------------
    arrCampos = Array("COD_PRODUCTO", _
                      "DES_PRODUCTO", _
                      "LABORATORIO", _
                      "LINEA", _
                      "COD_ESTADO", _
                      "FLG_SELECCIONADO", _
                      "STOCK", _
                      "CTD_SOLICITADA")
    
    arrCaption = Array("Código", _
                       "Descripción", _
                       "Laboratorio", _
                        "Línea", _
                        "Est.", _
                        "Sel.", _
                        "Stock", _
                        "Cant. Sol.")
    
    arrAncho = Array(650, _
                     3800, _
                     1400, _
                     1400, _
                     800, _
                     800, _
                     800, _
                     800)
    
    arrAlineacion = Array(dbgCenter, _
                          dbgLeft, _
                          dbgLeft, _
                          dbgLeft, _
                          dbgCenter, _
                          dbgCenter, _
                          dbgCenter, _
                          dbgCenter)
    
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDetalle.HeadLines = 2
    grdDetalle.RowHeight = 0
    grdDetalle.RowHeight = grdDetalle.RowHeight * 1.8

    
    grdDetalle.Columns("DES_PRODUCTO").FetchStyle = True
    grdDetalle.Columns("FLG_SELECCIONADO").Visible = False
    
End Sub




