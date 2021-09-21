VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_ADM_SMM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Modificación de Máximos"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   ControlBox      =   0   'False
   Icon            =   "frm_ADM_SMM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Emitidas"
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   3135
      Begin MSComCtl2.DTPicker dtpFchIni 
         Height          =   315
         Left            =   525
         TabIndex        =   6
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   55312387
         CurrentDate     =   37950
      End
      Begin MSComCtl2.DTPicker dtpFchFin 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   55312387
         CurrentDate     =   37950
      End
      Begin VB.Label Label7 
         Caption         =   "&Al"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "De&l"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "&Estado"
      Height          =   735
      Left            =   7200
      TabIndex        =   9
      Top             =   600
      Width           =   2175
      Begin vbp_Ventas.ctlDataCombo cboEstado 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         MatchEntry      =   1
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "&Producto"
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   2055
      Begin vbp_Ventas.ctlTextBox txtCod_Producto 
         Height          =   375
         Left            =   120
         TabIndex        =   3
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
      Caption         =   "&Número de Solicitud"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      Begin vbp_Ventas.ctlTextBox txtNumero 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   11
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
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1058
      ModoBotones     =   8
      EnabledEfecto   =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   0
      TabIndex        =   14
      Top             =   1200
      Width           =   10575
      Begin vbp_Ventas.ctlGrilla grdDetalle 
         Height          =   2775
         Left            =   120
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4895
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle de la Solicitud :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label lblSMM 
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
         Left            =   1920
         TabIndex        =   15
         Top             =   3120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_ADM_SMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrParam As Variant
Private strCodLocal As String

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

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    On Error GoTo CtrlErr
    Select Case boton
        Case Nuevo
            sub_Nuevo
        Case Modificar
        
        Case Buscar
            If txtNumero.Text = "" And dtpFchIni.Value > dtpFchFin.Value Then
                MsgBox "La Fecha ""DE"" no puede ser mayor a la fecha ""AL"" ", vbExclamation, "Atención"
                dtpFchIni.SetFocus
            Else
                sub_Buscar CodLocal, dtpFchIni.Value, dtpFchFin.Value, _
                            txtNumero.Text, cboEstado.BoundText, txtCod_Producto.Text
            End If

        Case tb_Actualizar
            sub_Actualizar
        Case Imprimir
            
        Case tb_Excel
            
        Case tb_email
            
        Case Grabar
            
        Case Cancelar
            
        Case Eliminar
            sub_Anular
        Case salir
            Unload Me
    End Select
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error"
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

Private Sub Form_Load()
    
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
    
    lblSMM.Caption = ""
    
    ctlToolBar1.Buttons(2).Visible = False
    ctlToolBar1.Buttons(6).Visible = False
    ctlToolBar1.Buttons(7).Visible = False
    
    'para que busque cuando cargue la pantalla
    sub_Buscar CodLocal, dtpFchIni.Value, dtpFchFin.Value, _
                        txtNumero.Text, cboEstado.BoundText, txtCod_Producto.Text, False


        
        
End Sub

Private Sub grdCabecera_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    Select Case Col
        Case grdCabecera.Columns("COD_ESTADO").ColIndex
            Select Case grdCabecera.Columns("COD_ESTADO").CellValue(Bookmark)
                Case "EMI"
                    CellStyle.BackColor = RGB(50, 175, 50)
                    CellStyle.ForeColor = vbWhite
                Case "ATE"
                    CellStyle.BackColor = RGB(50, 50, 175)
                    CellStyle.ForeColor = vbWhite
                Case "ANU"
                    CellStyle.BackColor = RGB(175, 50, 50)
                    CellStyle.ForeColor = vbWhite
            End Select
            
    End Select
    
End Sub

Private Sub grdCabecera_RegistroSeleccionado(ByVal DatoColumna0 As String)
    Dim objSMM As New clsSMM

    On Error GoTo CtrlErr

    grdDetalle.Limpiar
    Dim objEstadistica  As New clsEstadistica
    objEstadistica.Limpia
    Set objEstadistica = Nothing

    If grdCabecera.ApproxCount > 0 Then
        Set grdDetalle.DataSource = objSMM.ListaDet(CodLocal, grdCabecera.Columns("NUM_SMM").Value)
        Set objSMM = Nothing

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

Private Sub grdDetalle_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo CtrlErr
    
    Select Case grdDetalle.Columns(ColIndex).DataField
        Case "VENTAS"
            
        Dim frm As New frm_ADM_Graf_Ventas
        Dim objDatos As New clsEstadistica
        Dim odynDatos As oraDynaset
        Dim varDaTos(0 To 1, 0 To 9) As Variant
    
            
        Set odynDatos = objDatos.Lista(strCodLocal, grdDetalle.Columns("COD_PRODUCTO").Value, left(grdCabecera.Columns("NUM_SMM").Value, 6))
        
        odynDatos.MoveFirst
        
        While Not odynDatos.EOF
                    
            varDaTos(1, odynDatos("ORDEN").Value) = odynDatos("CTD_VENTA").Value
            varDaTos(0, odynDatos("ORDEN").Value) = odynDatos("PERIODO").Value
            odynDatos.MoveNext
            
        Wend
        
        varDaTos(1, 0) = "Cantidad de Venta"
        frm.Titulo = grdDetalle.Columns(0).Value & ": " & grdDetalle.Columns(1).Value
        frm.Datos = varDaTos
        frm.Mostrar
            
        Set frm = Nothing
            
            
    End Select
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"

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
Dim strNumero As String
Dim frm As New frm_ADM_SMMNuevo
    
    frm.Solicitud.CodLocal = CodLocal
    frm.Show vbModal
    strNumero = frm.Solicitud.Numero
    
    If strNumero <> "" Then
        If grdCabecera.ApproxCount > 0 Then
            sub_Actualizar strNumero
        Else
            sub_Buscar CodLocal, dtpFchIni.Value, dtpFchFin.Value, _
                        txtNumero.Text, cboEstado.BoundText, txtCod_Producto.Text
                        
            grdCabecera.DataSource.FindFirst " NUM_SMM = '" & strNumero & "'"

            If grdCabecera.DataSource.NoMatch Then
                grdCabecera.MoveFirst
            End If
        End If
    End If
    
    Set frm = Nothing
End Sub

Private Sub sub_Buscar(ByVal vstrCodLocal As String, _
                       ByVal vfecDesde As Date, _
                       ByVal vfecHasta As Date, _
                       ByVal vstrNumero As String, _
                       ByVal vstrEstado As String, _
                       ByVal vstrCodProducto As String, _
                       Optional ByVal vblnMsgShow As Boolean = True)
    
    Dim objSMM As New clsSMM
        
    
    Set grdCabecera.DataSource = objSMM.Lista(Format(vfecDesde, "DD/MM/YYYY"), _
                                              Format(vfecHasta, "DD/MM/YYYY"), _
                                              vstrCodLocal, _
                                              vstrEstado, _
                                              vstrNumero, _
                                              vstrCodProducto)
    Set objSMM = Nothing
    
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

    If grdCabecera.ApproxCount > 0 Then

        If vstrNumero = "" Then
            strNumero = grdCabecera.Columns("NUM_SMM").Value
        Else
            strNumero = vstrNumero
        End If

        sub_Buscar arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4), arrParam(5)

        grdCabecera.DataSource.FindFirst " NUM_SMM = '" & strNumero & "'"

        If grdCabecera.DataSource.NoMatch Then
            grdCabecera.MoveFirst
        End If

    End If

End Sub

Private Sub sub_Anular()
Dim objSMM As New clsSMM
Dim strNumero As String
Dim strError As String

    If grdCabecera.ApproxCount > 0 Then
        If grdCabecera.Columns("COD_ESTADO") = "EMI" Then
            strNumero = grdCabecera.Columns("NUM_SMM").Value
            If MsgBox("¿Desea Anular la Solicitud ->" & strNumero & "<- ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                objSMM.Anular CodLocal, strNumero, objUsuario.Codigo
                MsgBox "La Solicitud se anuló con exito", vbInformation, "Anular"
                sub_Actualizar
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
  
    '---------------------------------------------------------------
    '-- Cabecera
    '---------------------------------------------------------------
    
    arrCampos = Array("NUM_SMM", "FCH_EMISION", "USUARIO_EMISOR", _
                      "COD_ESTADO", "USUARIO_REVISA", "FCH_REVISION", _
                      "USUARIO_ANULA", "FCH_ANULACION")
                      
    arrCaption = Array("Número", "F.Emi.", "Emite", _
                       "Estado", "Revisa", "F.Rev.", _
                       "Anula", "F.Anu.")
    arrAncho = Array(1200, 1200, 1800, _
                     600, 1800, 1200, _
                     1800, 1200)
    
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, _
                          dbgCenter, dbgLeft, dbgCenter, _
                          dbgLeft, dbgCenter)
    
    grdCabecera.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdCabecera.Columns("COD_ESTADO").FetchStyle = True
    grdCabecera.Columns("NUM_SMM").FetchStyle = True
    grdCabecera.Columns("FCH_EMISION").NumberFormat = "DD/MM/YY"
    grdCabecera.Columns("FCH_REVISION").NumberFormat = "DD/MM/YY"
    grdCabecera.Columns("FCH_ANULACION").NumberFormat = "DD/MM/YY"
    
    '---------------------------------------------------------------
    '-- Detalle
    '---------------------------------------------------------------
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "LABORATORIO", _
                      "LINEA", "COD_ESTADO", "FLG_SELECCIONADO", _
                      "MOTIVO", "CTD_MAXIMO_APROBADO", "FCH_INICIAL_VIGENCIA", _
                      "FCH_FINAL_VIGENCIA", "VENTAS", "CTD_DIAS_QUIEBRE", _
                      "FLG_PROCESADO", "CTD_MAXIMO_REGISTRO", "CTD_MAXIMO_SOLICITADO", _
                      "OBS")
    
    arrCaption = Array("Código", "Descripción", "Laboratorio", _
                        "Línea", "Est.", "Sel.", _
                        "Motivo", "Máx Aprob.", "Desde", _
                        "Hasta", "Vnt.", "Quieb", _
                        "Aprob.", "Máx. Regis.", "Máx. Solic.", _
                        "Observación del Local")
    
    arrAncho = Array(650, 2800, 1200, _
                     1200, 500, 500, _
                      1200, 600, 1000, _
                      1000, 400, 500, _
                     500, 500, 500, _
                     1600)
    
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, _
                          dbgLeft, dbgCenter, dbgCenter, _
                           dbgLeft, dbgCenter, dbgCenter, _
                           dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, _
                          dbgLeft)
    
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDetalle.HeadLines = 2
    grdDetalle.RowHeight = 0
    grdDetalle.RowHeight = grdDetalle.RowHeight * 1.8
    
    grdDetalle.Columns("VENTAS").ButtonText = True
    grdDetalle.Columns("VENTAS").ButtonAlways = True
    
    grdDetalle.Columns("DES_PRODUCTO").FetchStyle = True
    grdDetalle.Columns("FLG_SELECCIONADO").Visible = False
    
    
    grdDetalle.Columns("MOTIVO").FetchStyle = True
    grdDetalle.Columns("CTD_MAXIMO_APROBADO").FetchStyle = True
    grdDetalle.Columns("FCH_INICIAL_VIGENCIA").FetchStyle = True
    grdDetalle.Columns("FCH_FINAL_VIGENCIA").FetchStyle = True
    
    grdDetalle.Columns("FLG_PROCESADO").Visible = False
    
    grdDetalle.Columns("FCH_INICIAL_VIGENCIA").NumberFormat = "DD/MM/YY"
    grdDetalle.Columns("FCH_FINAL_VIGENCIA").NumberFormat = "DD/MM/YY"
    
    
End Sub

