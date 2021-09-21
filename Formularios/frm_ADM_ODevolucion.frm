VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_ODevolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Devolución"
   ClientHeight    =   7935
   ClientLeft      =   180
   ClientTop       =   675
   ClientWidth     =   12555
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleMode       =   0  'User
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   60
      TabIndex        =   12
      Top             =   1560
      Width           =   12375
      Begin vbp_Ventas.ctlGrilla grdDetalle 
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4895
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlGrilla grdCabecera 
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4895
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Items:"
         Height          =   195
         Left            =   8880
         TabIndex        =   18
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label lblItems 
         Caption         =   "1000"
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
         Left            =   9840
         TabIndex        =   17
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblNroODev 
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
         Left            =   1680
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de la Orden :"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Número de Orden"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   2055
      Begin vbp_Ventas.ctlTextBox txtNumODev 
         Height          =   375
         Left            =   120
         TabIndex        =   10
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
   Begin VB.Frame Frame7 
      Caption         =   "&Producto"
      Height          =   735
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   2055
      Begin vbp_Ventas.ctlTextBox txtCod_Producto 
         Height          =   375
         Left            =   120
         TabIndex        =   8
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
   Begin VB.Frame Frame4 
      Caption         =   "&Estado"
      Height          =   735
      Left            =   7320
      TabIndex        =   5
      Top             =   720
      Width           =   2175
      Begin vbp_Ventas.ctlDataCombo cboEstado 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         MatchEntry      =   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vencimiento"
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   3135
      Begin MSComCtl2.DTPicker dtpFechaIni 
         Height          =   315
         Left            =   525
         TabIndex        =   1
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   62062595
         CurrentDate     =   37950
      End
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   62062595
         CurrentDate     =   37950
      End
      Begin VB.Label Label8 
         Caption         =   "De&l"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "&Al"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
   End
   Begin vbp_Ventas.ctlToolBar toolMain 
      Height          =   600
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1058
      ModoBotones     =   9
      EnabledEfecto   =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_ODevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objODev As clsOrdenDevolucion
Private oradynDetalle As oraDynaset

Private Sub BuscarOrdenes()
    Dim strFchInicio As String
    Dim strFchFinal As String
    
    On Error GoTo Control
    
    Set grdCabecera.DataSource = objODev.ListaOD(txtNumODev.Text, _
                                                vbNullString, _
                                                vbNullString, _
                                                dtpFechaIni.Value, _
                                                dtpFechaFin.Value, _
                                                cboEstado.BoundText, _
                                                vbNullString, _
                                                vbNullString, _
                                                objUsuario.CodigoLocal)
    grdCabecera.col = 0
    grdCabecera.SetFocus

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    On Error GoTo Control

    Set objODev = New clsOrdenDevolucion
    
    Call SetGrid
    
    Set cboEstado.RowSource = objODev.ListaEstados("")
    cboEstado.BoundColumn = "COD"
    cboEstado.ListField = "DES"
    cboEstado.BoundText = "*"

    'Me.Caption = "Administrador"
    Me.dtpFechaIni.Value = DateSerial(Year(Now), Month(Now), 1)
    Me.dtpFechaFin.Value = DateSerial(Year(Now), Month(Now) + 1, 0)
    
    With toolMain
        .Buttons(1).Visible = False
        .Buttons(7).Visible = False
        .Buttons(8).Visible = False
        .Buttons(12).Visible = False
        .Buttons(13).Visible = False
        .Buttons(2).Caption = "Atender"
    End With
    
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub AtenderOrden()
    On Error GoTo Control

    If grdCabecera.ApproxCount > 0 Then
        If grdCabecera.Columns("COD_ESTADO_REL").Value <> "EMI" And grdCabecera.Columns("COD_ESTADO_REL").Value <> "PAR" Then
            MsgBox "Solo se pueden atender ordenes en estado EMITIDO o PARCIALMENTE ATENDIDO.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If

        If Not IsDate(grdCabecera.Columns("FCH_VIGENCIA").Value) Then
            MsgBox "La fecha de vencimiento no es valida.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        If CDate(grdCabecera.Columns("FCH_VIGENCIA").Value) < CDate(Date) Then
            MsgBox "La orden de devolucion se encuentra vencida.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        
        With frm_ADM_ODevolucion_Atc
            .pstrNroOrdenDev = grdCabecera.Columns("NUM_ORDENDEV").Value
            .pstrTipoDevolucion = grdCabecera.Columns("COD_TIPODEV").Value
            .lblNroOrden.Caption = grdCabecera.Columns("NUM_ORDENDEV").Value
            .lblFchEmision.Caption = grdCabecera.Columns("FCH_ENVIO").Value
            .lblFchVigencia.Caption = grdCabecera.Columns("FCH_VIGENCIA").Value
            '.lblTipoDev.Caption = grdCabecera.Columns("COD_TIPODEV").Value & " - " & grdCabecera.Columns("DES_TIPODEV").Value
            .lblTipoDev.Caption = grdCabecera.Columns("DES_TIPODEV").Value
            '.lblMotivoDev.Caption = grdCabecera.Columns("COD_MOTIVODEV").Value & " - " & grdCabecera.Columns("DES_MOTIVODEV").Value
            .lblMotivoDev.Caption = grdCabecera.Columns("DES_MOTIVODEV").Value
            .Show vbModal, Me
            Call BuscarOrdenes
        End With
    End If

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub SetGrid()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim i As Integer

    On Error GoTo Control
    
    arrCampos = Array("NUM_ORDENDEV", "COD_ESTADO_REL", "FCH_ENVIO", "FCH_VIGENCIA", "COD_TIPODEV", "DES_TIPODEV", "COD_MOTIVODEV", "DES_MOTIVODEV", "COD_USUARIO", "NOMBRE", "FCH_ATENCION_LOCAL")
    arrCaption = Array("Nro. Orden", "Estado", "F.Emisión", "F.Vigencia", "", "Tipo Dev.", "", "Motivo Dev.", "", "Usuario", "")
    arrAncho = Array(1200, 900, 1100, 1100, 500, 3000, 500, 3000, 500, 3000, 1000)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgLeft, dbgCenter, dbgLeft, dbgLeft)
    
    With grdCabecera
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        For i = 0 To .Columns.Count - 1
            .Columns(i).AllowSizing = False
            .Columns(i).WrapText = False
        Next i
        .Columns("COD_ESTADO_REL").FetchStyle = True
        .Columns("COD_TIPODEV").Visible = False
        .Columns("COD_MOTIVODEV").Visible = False
        .Columns("COD_USUARIO").Visible = False
        .Columns("FCH_ATENCION_LOCAL").Visible = False
        .col = 0
        .RowHeight = 1.5 * .RowHeight
    End With

    arrCampos = Array("ITEM", "COD_PRODUCTO", "DES_PRODUCTO", "FLG_TOTAL_STK", _
                      "CTD_PRODUCTO", "CTD_PRODUCTO_DEV", "FLG_TOTAL_STK_FRAC", "CTD_PRODUCTO_FRAC", _
                      "CTD_PRODUCTO_FRAC_DEV", "CTD_STOCK", "FLG_FRACCIONA", "CTDU", _
                      "CTDF", "NRO_LOTE", "FCH_VENCIMIENTO")
    arrCaption = Array("Item", "Código", "Descripción", "Todo Stock Unidades", _
                       "Unidades Solicitadas", "Unidades Devevueltas", "Todo Stock Fracciones", "Fracciones Solicitadas", _
                       "Fracciones Devueltas", "Stock", "Fraccion", "CTDU", _
                       "CTDF", "Nro. Lote", "Fch. Vencimiento")
    arrAncho = Array(400, 800, 4000, 1000, _
                     1000, 1000, 1000, 1000, _
                     1000, 1000, 100, 100, _
                     100, 1500, 1500)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgLeft, dbgCenter)
    
    With grdDetalle
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        .HeadLines = 2
        For i = 0 To .Columns.Count - 1
            .Columns(i).AllowSizing = False
            .Columns(i).WrapText = False
            .Columns(i).Visible = False
        Next i
        .Columns("ITEM").Visible = True
        .Columns("COD_PRODUCTO").Visible = True
        .Columns("DES_PRODUCTO").Visible = True
        .Columns("CTD_STOCK").Visible = True
        .Columns("CTD_PRODUCTO_DEV").Visible = True
        .Columns("CTD_PRODUCTO_FRAC_DEV").Visible = True
        .Columns("FLG_TOTAL_STK").NumberFormat = "Yes/No"
        .Columns("FLG_TOTAL_STK_FRAC").NumberFormat = "Yes/No"
        .Columns("FLG_TOTAL_STK").FetchStyle = True
        .Columns("FLG_TOTAL_STK_FRAC").FetchStyle = True
        .col = 0
        .RowHeight = 1.5 * .RowHeight
    End With
    
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objODev = Nothing
    grdCabecera.Limpiar
    grdDetalle.Limpiar
    Set oradynDetalle = Nothing
End Sub

Private Sub grdCabecera_DblClick()
    Call toolMain_Click(Modificar, 0)
End Sub

Private Sub grdCabecera_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    On Error GoTo Control
    
    Select Case col
        Case grdCabecera.Columns("COD_ESTADO_REL").ColIndex
            Select Case grdCabecera.Columns("COD_ESTADO_REL").CellValue(Bookmark)
                Case "SOL"
                CellStyle.BackColor = RGB(215, 215, 255)
                CellStyle.ForeColor = vbBlack
            Case "PEN"
                CellStyle.BackColor = RGB(255, 255, 202)
                CellStyle.ForeColor = vbBlack
            Case "EMI"
                CellStyle.BackColor = RGB(50, 175, 50)
                CellStyle.ForeColor = vbWhite
            Case "ATE"
                CellStyle.BackColor = RGB(50, 50, 175)
                CellStyle.ForeColor = vbWhite
            Case "ANU"
                CellStyle.BackColor = RGB(175, 50, 50)
                CellStyle.ForeColor = vbWhite
            Case "PAR"
                CellStyle.BackColor = RGB(197, 137, 137)
                CellStyle.ForeColor = vbWhite
            Case "VEN"
                CellStyle.BackColor = RGB(255, 255, 255)
                CellStyle.ForeColor = RGB(255, 0, 0)
            End Select
    End Select

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub grdCabecera_RegistroSeleccionado(ByVal DatoColumna0 As String)
    On Error GoTo Control
    
    If grdCabecera.ApproxCount <= 0 Then lblNroODev.Caption = "": lblItems.Caption = "": Exit Sub
    lblNroODev.Caption = DatoColumna0
    Set oradynDetalle = objODev.ListaDetalleOD(grdCabecera.Columns("NUM_ORDENDEV").Value, objUsuario.CodigoLocal)
    
    Set grdDetalle.DataSource = oradynDetalle

    lblItems.Caption = oradynDetalle.RecordCount
    
    With grdDetalle
        .Columns("FLG_TOTAL_STK").Visible = True
        .Columns("CTD_PRODUCTO").Visible = True
        .Columns("FLG_TOTAL_STK_FRAC").Visible = True
        .Columns("CTD_PRODUCTO_FRAC").Visible = True
    End With
    
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub grdDetalle_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    On Error GoTo Control
    
    Select Case col
        Case grdDetalle.Columns("FLG_TOTAL_STK").ColIndex
            grdDetalle.Columns("CTD_PRODUCTO").Visible = Not CBool(grdDetalle.Columns("FLG_TOTAL_STK").Value)
            grdDetalle.Columns("FLG_TOTAL_STK").Visible = CBool(grdDetalle.Columns("FLG_TOTAL_STK").Value)
        Case grdDetalle.Columns("FLG_TOTAL_STK_FRAC").ColIndex
            grdDetalle.Columns("CTD_PRODUCTO_FRAC").Visible = Not CBool(grdDetalle.Columns("FLG_TOTAL_STK_FRAC").Value)
            grdDetalle.Columns("FLG_TOTAL_STK_FRAC").Visible = CBool(grdDetalle.Columns("FLG_TOTAL_STK_FRAC").Value)
    End Select

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub toolMain_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case boton
        Case Modificar
            Call AtenderOrden
        Case Buscar
            Call BuscarOrdenes
        Case salir
            Unload Me
        Case Else
            MsgBox "Opción no disponible por el momento.", vbCritical + vbOKOnly, "Error"
    End Select
End Sub
