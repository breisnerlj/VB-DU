VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_AdmGuias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Guías"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10545
   ControlBox      =   0   'False
   Icon            =   "frm_ADM_AdmGuias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1058
      ModoBotones     =   9
      EnabledEfecto   =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   10575
      Begin VB.Frame Frame6 
         Caption         =   "&Número de Guía"
         Height          =   735
         Left            =   1080
         TabIndex        =   0
         Top             =   120
         Width           =   2055
         Begin vbp_Ventas.ctlTextBox txtNum_Guia 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            Tipo            =   7
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
         Left            =   3120
         TabIndex        =   2
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton cmdBusProducto 
            Caption         =   "..."
            Height          =   375
            Left            =   1635
            TabIndex        =   4
            Top             =   240
            Width           =   315
         End
         Begin vbp_Ventas.ctlTextBox txtCod_Producto 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Tipo            =   3
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
         Left            =   8280
         TabIndex        =   9
         Top             =   120
         Width           =   2175
         Begin VB.ComboBox cboEstado 
            Height          =   315
            ItemData        =   "frm_ADM_AdmGuias.frx":0442
            Left            =   120
            List            =   "frm_ADM_AdmGuias.frx":0452
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   300
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Emitidas"
         Height          =   735
         Left            =   5160
         TabIndex        =   13
         Top             =   120
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
            Format          =   59047939
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
            Format          =   59047939
            CurrentDate     =   37950
         End
         Begin VB.Label Label8 
            Caption         =   "De&l"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "&Al"
            Height          =   255
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.OptionButton optDestino 
         Caption         =   "&Destino"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optOrigen 
         Caption         =   "&Origen"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6855
      Left            =   0
      TabIndex        =   18
      Top             =   1200
      Width           =   10575
      Begin vbp_Ventas.ctlGrilla grdDetalle 
         Height          =   3135
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5530
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlGrilla grdCabecera 
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5318
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label lblGuia 
         Caption         =   "999-9999999"
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
         TabIndex        =   20
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle de la Guía :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_ADM_AdmGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrParam As Variant
Private strCodLocal As String

Property Let CodLocal(ByVal vstrNewValue As String)
    strCodLocal = vstrNewValue
End Property

Private Sub cmdBusProducto_Click()
    frm_ADM_BuscaProd.Show vbModal
    txtCod_Producto.Text = frm_ADM_BuscaProd.strCodProducto
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case boton
        Case Nuevo
            sub_Nuevo
        Case Modificar
            sub_Recepcionar
        Case Buscar
            Dim strEstado As String
            Select Case cboEstado.ListIndex
                Case 0: strEstado = ""
                Case Else: strEstado = Mid(cboEstado.Text, 1, 3)
            End Select
            
            sub_Buscar optOrigen, dtpFchIni.Value, dtpFchFin.Value, _
                        txtNum_Guia.Text, strEstado, txtCod_Producto.Text
            
        Case tb_Actualizar
            sub_Actualizar
        Case Imprimir
            sub_Imprimir
        Case tb_Excel
            
        Case tb_email
            
        Case Grabar
            
        Case Cancelar
            
        Case eliminar
            sub_Anular
        Case salir
            Unload Me
    End Select
        
End Sub

Private Sub Form_Load()
    
    
    'SetteaFormulario Me, False
    'Caption = Caption & " - Local <<" & lstrCodLocal & ">>"
    
    'top = 0
    'left = 0
    Move 0, 0
    
    'Color = &H80000005
    'ColorFondo = &H80000005
    
    SeteaGrilla
    
    dtpFchIni.Value = gclsOracle.Fecha_Servidor - 30
    dtpFchFin.Value = gclsOracle.Fecha_Servidor
    lblGuia.Caption = ""
    
    ctlToolBar1.Buttons(1).Visible = False
    ctlToolBar1.Buttons(7).Visible = False
    
    cboEstado.Clear
    cboEstado.AddItem "[ TODOS ]", 0
    cboEstado.AddItem "TRANSITO", 1
    cboEstado.AddItem "RECEPCIONADA", 2
    cboEstado.AddItem "ANULADA", 3
    cboEstado.ListIndex = 1
    
End Sub

Private Sub grdCabecera_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    Select Case col
        Case grdCabecera.Columns("EST_GUIA").ColIndex
            Select Case grdCabecera.Columns("EST_GUIA").CellValue(Bookmark)
                Case "TRA"
                    CellStyle.BackColor = RGB(50, 175, 50)
                    CellStyle.ForeColor = vbWhite
                Case "REC"
                    CellStyle.BackColor = RGB(50, 50, 175)
                    CellStyle.ForeColor = vbWhite
                Case "ANU"
                    CellStyle.BackColor = RGB(175, 50, 50)
                    CellStyle.ForeColor = vbWhite
            End Select
            
        Case grdCabecera.Columns("ORIGEN").ColIndex
            If grdCabecera.Columns("ORIGEN").CellValue(Bookmark) <> strCodLocal Then
                CellStyle.ForeColor = vbYellow
                CellStyle.BackColor = vbBlack
                CellStyle.Font.Bold = True
            End If
            
        Case grdCabecera.Columns("DESTINO").ColIndex
            If grdCabecera.Columns("DESTINO").CellValue(Bookmark) <> strCodLocal Then
                CellStyle.ForeColor = vbYellow
                CellStyle.BackColor = vbBlack
                CellStyle.Font.Bold = True
            End If
        
        Case grdCabecera.Columns("NUM_GUIA").ColIndex
            CellStyle.Font.Bold = True
            If grdCabecera.Columns("TARJETA").CellValue(Bookmark) = "1" Then
                CellStyle.ForeColor = vbYellow
                CellStyle.BackColor = vbBlack
                CellStyle.Font.Bold = True
            End If
        Case grdCabecera.Columns("COD_BANDEJA").ColIndex
            
            If grdCabecera.Columns("TARJETA").CellValue(Bookmark) = "1" Then
                CellStyle.ForeColor = vbYellow
                CellStyle.BackColor = vbBlack
                CellStyle.Font.Bold = True
            End If

    End Select
    
End Sub

Private Sub grdCabecera_RegistroSeleccionado(ByVal DatoColumna0 As String)
    Dim objGuia As New clsGuia
    
    On Error GoTo CtrlErr
    
    grdDetalle.Limpiar
    
    If grdCabecera.ApproxCount > 0 Then
        Set grdDetalle.DataSource = objGuia.ListaDet(grdCabecera.Columns("NUM_GUIA").Value)
        Set objGuia = Nothing
        
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

Private Sub sub_Nuevo()

'    If strNumGuia <> "" Then
'        If grdCabecera.ApproxCount > 0 Then
'            sub_Actualizar strNumGuia
'        Else
'            sub_Buscar optOrigen.Value, dtpDesde.Value, dtpHasta.Value
'            grdCabecera.DataSource.FindFirst " NUM_GUIA = '" & strNumGuia & "'"
'
'            If grdCabecera.DataSource.NoMatch Then
'                grdCabecera.MoveFirst
'            End If
'        End If
'    End If

End Sub

Private Sub sub_Recepcionar()
Dim strNumGuia As String
Dim objGuia As New clsGuia
Dim strError As String

    On Error GoTo CtrlErr
    If grdCabecera.ApproxCount > 0 And optDestino.Value Then
        If grdCabecera.Columns("EST_GUIA") = "TRA" Then
            strNumGuia = grdCabecera.Columns("NUM_GUIA").Value

            If MsgBox("¿Desea Recepcionar la Guía Nro ->" & strNumGuia & "<- ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                strError = objGuia.Recepciona(objUsuario.CodigoEmpresa, strCodLocal, strNumGuia, objUsuario.codigo)
                If strError <> "" Then Err.Raise 1, "", strError
                sub_Actualizar strNumGuia
            End If
'            If MsgBox("¿Desea Recepcionar la Guía Nro ->" & strNumGuia & "<- ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'                With frm_ADM_AdmGuiasDet
'                    .pstrNumGuia = strNumGuia
'                    .pstrCodLocal = grdCabecera.Columns("DESTINO").Value
'                    .lblFchEmision = grdCabecera.Columns("FCH_EMISION").Value
'                    .lblOrigen = grdCabecera.Columns("ORIGEN").Value
'                    .lblDestino = grdCabecera.Columns("DESTINO").Value
'                    .lblDesObs = grdCabecera.Columns("DES_OBSERVACIONES").Value
'                    .Show vbModal
'                End With
'                sub_Actualizar strNumGuia
'            End If
        Else
            MsgBox "La Guía NO esta en transito", vbExclamation, "Recepcionar"
        End If
    End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error Recepcionar"
End Sub

Private Sub sub_Buscar(ByVal vblnOrigen As Boolean, _
                       ByVal vfecDesde As Date, _
                       ByVal vfecHasta As Date, _
                       ByVal vstrNumguia As String, _
                       ByVal vstrEstGuia As String, _
                       ByVal vstrCodProducto As String)
    
    Dim objGuia As New clsGuia
    
    On Error GoTo CtrlErr
    
    Set grdCabecera.DataSource = objGuia.ListaCab(vstrNumguia, _
                                                  IIf(vblnOrigen, strCodLocal, ""), _
                                                  IIf(vblnOrigen, "", strCodLocal), _
                                                  vstrEstGuia, _
                                                  Format(vfecDesde, "DD/MM/YYYY"), _
                                                  Format(vfecHasta, "DD/MM/YYYY"), , , , _
                                                  vstrCodProducto)
    Set objGuia = Nothing
    
    arrParam = Array(vblnOrigen, _
                     vfecDesde, _
                     vfecHasta, _
                     vstrNumguia, _
                     vstrEstGuia, _
                     vstrCodProducto)
    
    If grdCabecera.ApproxCount < 1 Then
        grdDetalle.Limpiar
        MsgBox "Búsqueda sin Resultados", vbInformation, "Aviso"
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error Buscar"
End Sub

Private Sub sub_Actualizar(Optional ByVal vstrNumguia As String = "")
Dim strNumGuia As String
    
On Error GoTo CtrlErr
    
    If grdCabecera.ApproxCount > 0 Then
        
        If vstrNumguia = "" Then
            strNumGuia = grdCabecera.Columns("NUM_GUIA").Value
        Else
            strNumGuia = vstrNumguia
        End If
        
        sub_Buscar arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4), arrParam(5)
        
        grdCabecera.DataSource.FindFirst " NUM_GUIA = '" & strNumGuia & "'"
        
        If grdCabecera.DataSource.NoMatch Then
            grdCabecera.MoveFirst
        End If
                        
    End If
    
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error Actualizar"
End Sub

Private Sub sub_Imprimir()
Dim objDocumento As New clsDocumento
On Error GoTo CtrlErr
    If MsgBox("Coloque papel en impresora " & Chr(13) & "¿Continua con la impresión?", vbQuestion + vbYesNo + vbDefaultButton2, "Impresión de Guía de Remisión") = vbNo Then
        Exit Sub
    End If
    If grdDetalle.ApproxCount > 0 Then
        Select Case grdCabecera.Columns("DESTINO").Value
'               Case "PRV"
'                    Dim objGuia As New clsGuia
'                    objGuia.spImprime_Guia_Dev "", "", grdCabecera.Columns("NUM_GUIA").Value
'                    Set objGuia = Nothing
               Case "CLI"
                    'objDocumento.Guia objUsuario.CodigoEmpresa, "GRL", grdCabecera.Columns("NUM_GUIA").Value
                    objDocumento.GuiaNew objUsuario.CodigoEmpresa, "GRL", grdCabecera.Columns("NUM_GUIA").Value
                    Set objDocumento = Nothing
               Case Else
                    objDocumento.ImprimirGuiaTransferencia grdCabecera.Columns("NUM_GUIA").Value
                    Set objDocumento = Nothing
        End Select
    End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

''    Dim lstrNumGuia As String
''
''    If grdCabecera.ApproxCount < 1 Then
''        MsgBox "Debe seleccionar una Guia", vbInformation
''        Exit Sub
''    End If
''
''    lstrNumGuia = grdCabecera.Columns("NUM_GUIA").Value
''    Dim objGuia As New clsDocumento
''    objGuia.ImprimirGuiaTransferencia lstrNumGuia
''    Set objGuia = Nothing
End Sub

Private Sub sub_Anular()
Dim objGuia As New clsGuia
Dim strNumGuia As String
Dim strError As String
On Error GoTo CtrlErr
    If grdCabecera.ApproxCount > 0 Then
        If grdCabecera.Columns("EST_GUIA") = "TRA" Then
            strNumGuia = grdCabecera.Columns("NUM_GUIA").Value
            If MsgBox("¿Desea Anular la Guía ->" & strNumGuia & "<- ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                strError = objGuia.Anula(objUsuario.CodigoEmpresa, strCodLocal, strNumGuia, objUsuario.codigo)
                If strError <> "" Then Err.Raise 1, "", strError
                sub_Actualizar strNumGuia
            End If
        Else
            MsgBox "La Guía NO esta en transito", vbExclamation, "Anular"
        End If
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Anular"
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    '---------------------------------------------------------------
    '-- Cabecera
    '---------------------------------------------------------------
    
    arrCampos = Array("ORIGEN", "DESTINO", "NUM_GUIA", _
                      "EST_GUIA", "FCH_EMISION", "FCH_RECEPCION", _
                      "FCH_ANULACION", "COD_BANDEJA", "DES_OBSERVACIONES", _
                      "EMISOR", "RECEPTOR", "ANULADOR", "TARJETA")
                      
    arrCaption = Array("De", "A", "Número", _
                        "Estado", "F.Emisión", "F.Recep.", _
                        "F.Anu.", "Bandeja", "Observaciones", "Emisor", _
                        "Receptor", "Anulador", "Tar")
    arrAncho = Array(500, 500, 1200, _
                    450, 900, 900, _
                    900, 1200, 1600, 1800, _
                    1800, 1800, 0)
    
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft)
    
    grdCabecera.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdCabecera.Columns("EST_GUIA").FetchStyle = True
    grdCabecera.Columns("ORIGEN").FetchStyle = True
    grdCabecera.Columns("DESTINO").FetchStyle = True
    grdCabecera.Columns("NUM_GUIA").FetchStyle = True
    grdCabecera.Columns("COD_BANDEJA").FetchStyle = True
    grdCabecera.Columns("FCH_EMISION").NumberFormat = "DD/MM/YY"
    grdCabecera.Columns("FCH_RECEPCION").NumberFormat = "DD/MM/YY"
    grdCabecera.Columns("FCH_ANULACION").NumberFormat = "DD/MM/YY"
    grdCabecera.Columns("TARJETA").Visible = False
    
    
    '---------------------------------------------------------------
    '-- Detalle
    '---------------------------------------------------------------
    arrCampos = Array("ORDEN", "COD_PRODUCTO", "DESCRIPCION", _
                      "LABORATORIO", "FRACCIONO", "CTD_FRACCIONO", _
                      "CTD_PRODUCTO", "CTD_PRODUCTO_FRAC")
    
    arrCaption = Array("#", "Código", "Descripción", _
                        "Laboratorio", "Frac.", "C.F.", _
                        "U", "F")
    
    arrAncho = Array(500, 800, 4000, _
                    1600, 600, 800, _
                    900, 900)
    
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, _
                          dbgLeft, dbgCenter, dbgRight, _
                          dbgRight, dbgRight)
    
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub


