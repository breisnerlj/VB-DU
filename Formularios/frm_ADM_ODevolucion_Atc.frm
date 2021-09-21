VERSION 5.00
Begin VB.Form frm_ADM_ODevolucion_Atc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atención de Ordenes de Devolución"
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
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   8535
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   6720
         TabIndex        =   25
         Top             =   280
         Width           =   1335
      End
      Begin vbp_Ventas.ctlTextBox txtBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   300
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         Tipo            =   8
         TABAuto         =   0   'False
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto (F1):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   720
         TabIndex        =   24
         Top             =   380
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   8760
      TabIndex        =   19
      Top             =   2160
      Width           =   3735
      Begin VB.CommandButton cmdPropagar 
         Caption         =   "Propagar"
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   12375
      Begin VB.CommandButton cmdRecuperar 
         Caption         =   "Recuperar"
         Height          =   375
         Left            =   10800
         TabIndex        =   18
         Top             =   280
         Width           =   1095
      End
      Begin vbp_Ventas.ctlDataCombo cboLaboratorio 
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   300
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboLinea 
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Top             =   300
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   5760
         TabIndex        =   17
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Laboratorio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   16
         Top             =   360
         Width           =   1035
      End
   End
   Begin vbp_Ventas.ctlGrillaArray grdDetalle 
      Height          =   4695
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8281
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   9975
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Orden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Emisión:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   7080
         TabIndex        =   10
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Vigencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7080
         TabIndex        =   9
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Dev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo Dev.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label lblNroOrden 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Orden:"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lblFchEmision 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Orden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   8520
         TabIndex        =   5
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblFchVigencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Orden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   8520
         TabIndex        =   4
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblTipoDev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Orden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblMotivoDev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Orden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   945
      End
   End
   Begin vbp_Ventas.ctlToolBar toolForm 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1058
      ModoBotones     =   6
   End
End
Attribute VB_Name = "frm_ADM_ODevolucion_Atc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrNroOrdenDev As String
Public pstrTipoDevolucion As String
Public pstrMotivoDevolucion As String
Private bolVerUnid As Boolean
Private bolVerFrac As Boolean
Private objODev As clsOrdenDevolucion
Private objDocumento As clsDocumento
Private objLaboratorio As clsLaboratorio
Private xarrDetalle As New XArrayDB
Private strCodTipo As String

Private Enum sCol
    NRO_ITEM
    COD_PRODUCTO
    DES_PRODUCTO
    FLG_FRACCIONA
    CTD_STOCK
    FLG_TOTAL_STK
    CTD_PRODUCTO
    FLG_TOTAL_STK_FRAC
    CTD_PRODUCTO_FRAC
    CTD_PRODUCTO_DEV
    CTD_PRODUCTO_FRAC_DEV
    CTDU
    CTDF
    NRO_LOTE
    FCH_VENCIMIENTO
End Enum

Private Sub GrabarOrdenDevolucion()
    Dim i As Integer
    Dim vstrCadCodProducto As String
    Dim vstrCadCtdProductoFrac As String
    Dim vstrCadCtdProducto As String
    Dim vstrCadImpProducto As String
    Dim vstrCadNroLote As String
    Dim vstrCadFchVencimiento As String
    Dim vstrGrabaGuia As String
    Dim vstrCambiaEstado As String
    Dim strNumDocumento As String
    Dim strMensaje As String
    Dim arrDocs As Variant
    Dim vstrObs As String
    
    On Error GoTo Control
    
    strMensaje = "Asegurese de verificar las cantidades, los números de lote y las fechas de vencimiento" & vbCrLf & _
                 "de los productos a devolver antes de generar la Guía." & vbCrLf & _
                 "¿Desea continuar generando la Guía?"
    
    If MsgBox(strMensaje, vbQuestion + vbYesNo, "Atender Orden") = vbNo Then
        Exit Sub
    End If
    
    Set objDocumento = New clsDocumento
    
    vstrCadCodProducto = vbNullString
    vstrCadCtdProductoFrac = vbNullString
    vstrCadCtdProducto = vbNullString
    vstrCadImpProducto = vbNullString
    vstrCadNroLote = vbNullString
    vstrCadFchVencimiento = vbNullString
    vstrObs = Trim(lblNroOrden.Caption) & " " & Trim(lblTipoDev.Caption) & " " & Mid(Trim(lblMotivoDev.Caption), 8)

    For i = xarrDetalle.LowerBound(1) To xarrDetalle.UpperBound(1)
        If IsEmpty(xarrDetalle.Value(i, sCol.CTD_PRODUCTO_DEV)) And IsEmpty(xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC_DEV)) Then
           grdDetalle.Bookmark = i
           MsgBox "No se han ingresado cantidades a devolver.", vbCritical, "Error"
           Exit Sub
        End If

        If Val(Trim("" & xarrDetalle.Value(i, sCol.CTD_PRODUCTO_DEV))) = 0 And Val(Trim("" & xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC_DEV))) = 0 Then
        Else
            If CBool(xarrDetalle.Value(i, sCol.FLG_FRACCIONA)) = False And Val(xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC_DEV)) > 0 Then
               grdDetalle.Bookmark = i
               MsgBox "El producto actual no fracciona.", vbCritical, "Error"
               grdDetalle.col = sCol.CTD_PRODUCTO_FRAC_DEV
               Exit Sub
            End If
            
            If strCodTipo = pstrTipoDevolucion And IsEmpty(xarrDetalle.Value(i, sCol.FCH_VENCIMIENTO)) Then
               grdDetalle.Bookmark = i
               MsgBox "No se ha ingresado la fecha de vencimiento.", vbCritical, "Error"
               Exit Sub
            End If
            
            vstrCadCodProducto = vstrCadCodProducto & _
                                    Trim(xarrDetalle.Value(i, sCol.COD_PRODUCTO)) & "|"
            vstrCadCtdProducto = vstrCadCtdProducto & _
                                        Val(Trim("" & xarrDetalle.Value(i, sCol.CTD_PRODUCTO_DEV))) & "|"
            vstrCadCtdProductoFrac = vstrCadCtdProductoFrac & _
                                    Val(Trim("" & xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC_DEV))) & "|"
            vstrCadImpProducto = vstrCadImpProducto & "0|"
            vstrCadNroLote = vstrCadNroLote & _
                                Trim(xarrDetalle.Value(i, sCol.NRO_LOTE)) & "|"
        
            vstrCadFchVencimiento = vstrCadFchVencimiento & _
                                        Format(Trim(xarrDetalle.Value(i, sCol.FCH_VENCIMIENTO)), "DD/MM/YYYY") & "|"
                                        
                                                    
            If vstrCadFchVencimiento <> "" And Len(vstrCadNroLote) = 1 Then
                vstrCadNroLote = "<SIN_LOTE>"
            End If
                                                    
        End If
    Next i
    
    If Len(Trim(vstrCadCodProducto)) > 0 Then
        'Generar la guia
        vstrGrabaGuia = objDocumento.GeneraGuiaLocal(strNumDocumento, _
                                                     objUsuario.CodigoEmpresa, _
                                                     objUsuario.CodigoLocal, _
                                                     "CJE", _
                                                     objUsuario.NombrePC, _
                                                     objUsuario.TipoDocGuia, _
                                                     objUsuario.MotivoGeneraGuiaLocal, _
                                                     objUsuario.Codigo, _
                                                     vstrObs, _
                                                     vstrCadCodProducto, _
                                                     vstrCadCtdProducto, _
                                                     vstrCadCtdProductoFrac, _
                                                     "", _
                                                     "", _
                                                     vstrCadNroLote, _
                                                     vstrCadFchVencimiento, _
                                                     "ODE", _
                                                     pstrNroOrdenDev, "", "", "")

        If Len(Trim(vstrGrabaGuia)) = 0 Then

            arrDocs = Split(strNumDocumento, "|")

            vstrCambiaEstado = objODev.AtenderOrden(pstrNroOrdenDev, _
                                                    objUsuario.CodigoLocal, _
                                                    vstrCadCodProducto, _
                                                    vstrCadCtdProducto, _
                                                    vstrCadCtdProductoFrac)
            
            If Len(Trim(vstrCambiaEstado)) <> 0 Then
                MsgBox vstrCambiaEstado, vbCritical, "Error"
                Exit Sub
            End If
            
            strMensaje = "Se generó los siguientes documentos:" & vbCrLf & Join(arrDocs, vbCrLf)
            MsgBox strMensaje, vbOKOnly + vbInformation, App.ProductName
            
            For i = LBound(arrDocs) To UBound(arrDocs)
                If Trim(arrDocs(i)) = vbNullString Then Exit For
                If MsgBox("Guía de Remisión Nro." & arrDocs(i) & vbCrLf & vbCrLf & _
                          "Coloque papel en impresora " & vbCrLf & _
                          "¿Continua con la impresión?", vbQuestion + vbYesNo + vbDefaultButton2, "Impresión de Guía de Remisión") = vbYes Then
                    objDocumento.ImprimirGuiaTransferencia arrDocs(i)
                End If
            Next i
            
            Unload Me
        Else
            MsgBox vstrGrabaGuia, vbCritical, "Error"
        End If
    
    Else
    
        MsgBox "Al menos un producto debe tener las cantidades mayores a cero", vbCritical, "Error"
    
    End If
    
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub cboLaboratorio_Change()
    Dim objLinea As New clsLinea
    
    On Error GoTo Control
    
    With cboLinea
        Set .RowSource = objLinea.Lista(cboLaboratorio.BoundText)
        .ListField = "DES"
        .BoundColumn = "COD"
        .BoundText = "*"
    End With
    
    Set objLinea = Nothing

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdBuscar_Click()
    Dim rs As oraDynaset
    Dim resultIndex As Long
    Dim objProducto As clsProducto
    
    On Error GoTo Control
    
    If Len(Trim(txtBuscar.Text)) < 3 Then
        MsgBox "Ingresar como minimo 3 digitos", vbInformation + vbOKOnly, App.ProductName
        txtBuscar.SetFocus
        Exit Sub
    End If

    Me.MousePointer = vbHourglass
    
    Set objProducto = New clsProducto
'    Set rs = objProducto.ListaBusqueda(Trim(txtBuscar.Text))
    Set rs = objProducto.Lista(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"), Trim(txtBuscar.Text), "", "", objUsuario.CodigoLocal)
    
    Me.MousePointer = vbDefault
    
    If rs.RecordCount = 0 Then
        MsgBox "Producto no encontrado.", vbInformation + vbOKOnly, App.ProductName
        txtBuscar.SetFocus
        GoTo Final
    ElseIf rs.RecordCount > 1 Then
        MsgBox "El criterio de busqueda devuelve más de un producto como resultado." & vbNewLine & _
               "Sea más específico.", vbInformation + vbOKOnly, App.ProductName
        txtBuscar.SetFocus
        GoTo Final
    Else
        resultIndex = xarrDetalle.Find(xarrDetalle.LowerBound(1), _
                                       CInt(sCol.COD_PRODUCTO), _
                                       CStr(rs.Fields("COD_PRODUCTO").Value), XORDER_ASCEND, XCOMP_EQ, XTYPE_NUMBER)
        If resultIndex = xarrDetalle.LowerBound(1) - 1 Then
            MsgBox "El producto indicado no se encuentra en la lista." & vbNewLine & _
                   rs.Fields(0) & " - " & rs.Fields(1), vbInformation + vbOKOnly, App.ProductName
            txtBuscar.SetFocus
            GoTo Final
        Else
            grdDetalle.Bookmark = resultIndex
            'grdDetalle.col = CInt(Text2.Text)
            grdDetalle.SetFocus
            GoTo Final
        End If
    End If

Final:
    Set rs = Nothing
    Set objProducto = Nothing
    Exit Sub

Control:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    GoTo Final
End Sub

Private Sub cmdLimpiar_Click()
    On Error GoTo Control
    
    Call LimpiarDetalle

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdPropagar_Click()
    On Error GoTo Control
    
    Call PropagarCeros
    
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdRecuperar_Click()
    On Error GoTo Control
    
    Call LlenarGrid(pstrNroOrdenDev, _
                    objUsuario.CodigoLocal, _
                    cboLaboratorio.BoundText, _
                    cboLinea.BoundText)

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Activate()
    grdDetalle.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objDocumento = Nothing
    Set objODev = Nothing
    Set objLaboratorio = Nothing
End Sub

Private Sub grdDetalle_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Select Case ColIndex
        Case sCol.CTD_PRODUCTO_DEV, sCol.CTD_PRODUCTO_FRAC_DEV, sCol.NRO_LOTE, sCol.FCH_VENCIMIENTO
            Cancel = False
        Case Else
            Cancel = True
    End Select
End Sub

Private Sub toolForm_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    On Error GoTo Control
    
    Select Case Index
        Case 1
            Call GrabarOrdenDevolucion
        Case 2, 3
            If MsgBox("Esto descartará los cambios realizados y cerrará el formulario." & vbCrLf & _
               "¿Está seguro que desea continuar?", vbYesNo + vbQuestion, "Cancelar") = vbYes Then
                Unload Me
            End If
    End Select

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call toolForm_Click(salir, 3)
        Case vbKeyF1
            txtBuscar.SetFocus
    End Select
End Sub

Private Sub LlenarGrid(vNroOrden As String, _
                       vCodLocal As String, _
                       Optional vCodLaboratorio As String = "", _
                       Optional vCodLinea As String = "")
    Dim rsDetalle As oraDynaset
    Dim i As Integer, intsep As Integer
    Dim strStock As String
    Dim dblStkUnid As Double
    Dim dblStkFrac As Double

    On Error GoTo Control
    
    vCodLaboratorio = Replace(vCodLaboratorio, "*", "")
    vCodLinea = Replace(vCodLinea, "*", "")
    
    Set rsDetalle = objODev.ListaDetalleOD(vNroOrden, vCodLocal, vCodLaboratorio, vCodLinea)
    
    If rsDetalle.RecordCount = 0 Then
        Call LimpiarDetalle
        Exit Sub
    End If
    
    xarrDetalle.LoadRows rsDetalle.GetRows
    xarrDetalle.AppendColumns 2
     
    For i = xarrDetalle.LowerBound(1) To xarrDetalle.UpperBound(1)
   
        strStock = vbNullString
        
        If i = 0 Then
            bolVerUnid = CBool(Val("" & xarrDetalle.Value(i, sCol.FLG_TOTAL_STK)))
            bolVerFrac = CBool(Val("" & xarrDetalle.Value(i, sCol.FLG_TOTAL_STK_FRAC)))
        End If
        
        'Obtener la posicion del separador de fracciones del stock
        strStock = xarrDetalle.Value(i, sCol.CTD_STOCK)
        intsep = InStr(1, strStock, "F")
        dblStkUnid = IIf(intsep = 0, strStock, Val(Trim(Mid(strStock, 1, intsep))))
        dblStkFrac = IIf(intsep = 0, 0, Val(Trim(Mid(strStock, intsep + 1, Len(strStock)))))

        If CBool(xarrDetalle.Value(i, sCol.FLG_TOTAL_STK)) = True Then
            'Obtener la cantidad de unidades del stock
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO) = dblStkUnid
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO_DEV) = dblStkUnid
        Else
            'Obtener la cantidad de unidades de la orden si el stock mayor a 0, sino, igual a 0
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO) = xarrDetalle.Value(i, sCol.CTDU)
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO_DEV) = _
                        IIf(xarrDetalle.Value(i, sCol.CTDU) < dblStkUnid, xarrDetalle.Value(i, sCol.CTDU), dblStkUnid)
        End If
    
        If CBool(xarrDetalle.Value(i, sCol.FLG_TOTAL_STK_FRAC)) = True Then
            'Obtener la cantidad de fracciones del stock
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC) = dblStkFrac
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC_DEV) = dblStkFrac
        Else
            'Obtener la cantidad de fracciones de la orden
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC) = xarrDetalle.Value(i, sCol.CTDF)
            xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC_DEV) = _
                        IIf(xarrDetalle.Value(i, sCol.CTDF) < dblStkFrac, xarrDetalle.Value(i, sCol.CTDF), dblStkFrac)
        End If
        
    Next i
    
    grdDetalle.Array1 = xarrDetalle
    grdDetalle.Rebind
    
    Set rsDetalle = Nothing
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub LimpiarDetalle()
    Dim i As Integer
    
    On Error GoTo Control
    
    If grdDetalle.ApproxCount < 1 Then Exit Sub
    
    If MsgBox("¿Seguro(a) de Eliminar Items que no tienen cantidades a devolver?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    grdDetalle.MoveFirst
    While Not grdDetalle.EOF
        If Val(Trim("" & grdDetalle.Columns(sCol.CTD_PRODUCTO_DEV).Value)) = 0 And Val(Trim("" & grdDetalle.Columns(sCol.CTD_PRODUCTO_FRAC_DEV).Value)) = 0 Then
            grdDetalle.Delete
        Else
            grdDetalle.MoveNext
        End If
    Wend
    grdDetalle.col = 0
    grdDetalle.MoveFirst
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub PropagarCeros()
    Dim i As Integer
    Dim strVal As String
    Dim strCol As String
    
    On Error GoTo Control
    
    If grdDetalle.EditActive = True Then Exit Sub
    If grdDetalle.ApproxCount < 1 Then Exit Sub
    
    strCol = grdDetalle.col
    strVal = grdDetalle.Columns(grdDetalle.col).Value
    
    For i = xarrDetalle.LowerBound(1) To xarrDetalle.UpperBound(1)
        xarrDetalle.Value(i, strCol) = strVal
    Next i
    grdDetalle.Rebind
    grdDetalle.col = strCol

    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Form_Load()
    
    On Error GoTo Control
    
    '-- Lista Laboratorios  --'
    Set objLaboratorio = New clsLaboratorio
    With cboLaboratorio
        Set .RowSource = objLaboratorio.Lista
        .ListField = "DES"
        .BoundColumn = "COD"
        .BoundText = "*"
    End With

    strCodTipo = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_TIPODEV_CJE_VENC")

    Set objODev = New clsOrdenDevolucion
    
    Call LlenarGrid(pstrNroOrdenDev, objUsuario.CodigoLocal)
    Call SetGrid

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
    Dim s As TrueDBGrid70.Split
    Dim c As Column
    
    On Error GoTo Control
    
    arrCampos = Array("NRO_ITEM", "COD_PRODUCTO", "DES_PRODUCTO", "FLG_FRACCIONA", _
                    "CTD_STOCK", "FLG_TOTAL_STK", "CTD_PRODUCTO", "FLG_TOTAL_STK_FRAC", _
                    "CTD_PRODUCTO_FRAC", "CTD_PRODUCTO_DEV", "CTD_PRODUCTO_FRAC_DEV", "CTDU", _
                      "CTDF", "NRO_LOTE", "FCH_VENCIMIENTO")
    arrCaption = Array("Item", "Código", "Descripción", "Fracciona", _
                       "Stock", "Todo Unidades", "Unidades Solicitadas", "Todo Fracciones", _
                       "Fracciones Solicitadas", "Unidades A Devolver", "Fracciones A Devolver", "CTDU", _
                       "CTDF", "Nro. Lote", "Fch. Vencimiento")
    arrAncho = Array(400, 700, 4000, 1000, _
                     700, 900, 1000, 900, _
                     1000, 1000, 1000, 1000, _
                     1000, 1500, 1500)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgLeft, dbgCenter)
    
    With grdDetalle
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        .HeadLines = 2
        .EditorStyle.BackColor = vbWhite
        .EditorStyle.ForeColor = RGB(180, 0, 180)
        .EditorStyle.Font.Bold = True
        .AllowUpdate = True
        .RowHeight = 1.5 * .RowHeight
        .MarqueeStyle = 2
        .col = 0

        For i = 0 To .Columns.Count - 1
            .Columns(i).AllowSizing = False
            .Columns(i).WrapText = False
            .Columns(i).Visible = False
        Next i
        
        .Columns(sCol.FLG_TOTAL_STK).FetchStyle = True
        .Columns(sCol.FLG_TOTAL_STK_FRAC).FetchStyle = True
        
        .Columns(sCol.FLG_TOTAL_STK).NumberFormat = "Yes/No"
        .Columns(sCol.FLG_TOTAL_STK_FRAC).NumberFormat = "Yes/No"
        .Columns(sCol.FCH_VENCIMIENTO).NumberFormat = "MM/YYYY"

        'Columnas editables
        .Columns(sCol.CTD_PRODUCTO_DEV).BackColor = vbInfoBackground
        .Columns(sCol.CTD_PRODUCTO_DEV).DataWidth = 4
        .Columns(sCol.CTD_PRODUCTO_FRAC_DEV).BackColor = vbInfoBackground
        .Columns(sCol.CTD_PRODUCTO_FRAC_DEV).DataWidth = 4
        .Columns(sCol.NRO_LOTE).BackColor = vbInfoBackground
        .Columns(sCol.NRO_LOTE).DataWidth = 20
        .Columns(sCol.FCH_VENCIMIENTO).BackColor = vbInfoBackground
        .Columns(sCol.FCH_VENCIMIENTO).DataWidth = 10
    End With

    Set s = grdDetalle.Splits.Add(1)

    With grdDetalle.Splits(0)
        .RecordSelectors = False
        .SizeMode = dbgNumberOfColumns
        .Size = 6
        .AllowSizing = False
        .Columns(sCol.NRO_ITEM).Visible = True
        .Columns(sCol.COD_PRODUCTO).Visible = True
        .Columns(sCol.DES_PRODUCTO).Visible = True
        .Columns(sCol.CTD_STOCK).Visible = True
        .Columns(sCol.FLG_TOTAL_STK).Visible = bolVerUnid
        .Columns(sCol.CTD_PRODUCTO).Visible = Not bolVerUnid
        .Columns(sCol.FLG_TOTAL_STK_FRAC).Visible = bolVerFrac
        .Columns(sCol.CTD_PRODUCTO_FRAC).Visible = Not bolVerFrac
    End With

    With grdDetalle.Splits(1)
        .RecordSelectors = False
        .SizeMode = dbgScalable
        .AllowSizing = False
        .Columns(sCol.CTD_PRODUCTO_DEV).Visible = True
        .Columns(sCol.CTD_PRODUCTO_FRAC_DEV).Visible = True
        .Columns(sCol.NRO_LOTE).Visible = True
        .Columns(sCol.FCH_VENCIMIENTO).Visible = True
    End With

    grdDetalle.Rebind
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub grdDetalle_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Select Case ColIndex
        Case sCol.CTD_PRODUCTO_DEV
            If Not IsNumeric(Trim(grdDetalle.Columns(ColIndex).Value)) And _
                    Trim(grdDetalle.Columns(ColIndex).Value) <> "" Then
                MsgBox "El valor no es valido", vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
            
            If Val(xarrDetalle(grdDetalle.Bookmark, sCol.CTD_STOCK)) < Val(Trim(grdDetalle.Columns(ColIndex).Value)) Then
                MsgBox "La unidades a devolver no pueden ser mayores al stock de unidades disponible.", _
                                vbCritical, Caption
                Cancel = True
                Exit Sub
            End If
            
            If pstrTipoDevolucion <> strCodTipo Then
                If Val(Trim(grdDetalle.Columns(ColIndex).Value)) > Val(Trim(grdDetalle.Columns(sCol.CTD_PRODUCTO).Value)) Then
                    MsgBox "La unidades a devolver no pueden ser mayores a las unidades solicitadas.", _
                            vbCritical, "Error"
                    Cancel = True
                    Exit Sub
                End If
           End If
           
        Case sCol.CTD_PRODUCTO_FRAC_DEV
        
            If Not IsNumeric(Trim(grdDetalle.Columns(ColIndex).Value)) And _
                    Trim(grdDetalle.Columns(ColIndex).Value) <> "" Then
                MsgBox "El valor no es valido", vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
            
            If CBool(Val(Trim(grdDetalle.Columns(sCol.FLG_FRACCIONA).Value))) = False And Val(Trim(grdDetalle.Columns(sCol.CTD_PRODUCTO_FRAC_DEV).Value)) > 0 Then
                MsgBox "El producto seleccionado no fracciona.", _
                        vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
            
            If pstrTipoDevolucion <> strCodTipo Then
                If Val(Trim(grdDetalle.Columns(ColIndex).Value)) > Val(Trim(grdDetalle.Columns(sCol.CTD_PRODUCTO_FRAC).Value)) Then
                    If CBool(Val(Trim(grdDetalle.Columns(sCol.FLG_TOTAL_STK_FRAC).Value))) = True Or pstrTipoDevolucion <> strCodTipo Then
                        MsgBox "La fracciones a devolver no pueden ser mayores al stock de fracciones disponible.", _
                                vbCritical, "Error"
                        Cancel = True
                        Exit Sub
                    Else
                        MsgBox "La fracciones a devolver no pueden ser mayores a las fracciones solicitadas.", _
                                vbCritical, "Error"
                        Cancel = True
                        Exit Sub
                    End If
                End If
                
             Else
             
               If Val(xarrDetalle(grdDetalle.Bookmark, sCol.CTD_STOCK)) < Val(Trim(grdDetalle.Columns(ColIndex).Value)) Then
                    MsgBox "La fracciones a devolver no pueden ser mayores a las fracciones solicitadas.", _
                                vbCritical, "Error"
                    Cancel = True
                    Exit Sub
               End If
                
            End If
            
        Case sCol.FCH_VENCIMIENTO
            Dim FchVenc As String
                                                          
            If Not IsDate(Trim(grdDetalle.Columns(ColIndex).Value)) And Trim(grdDetalle.Columns(ColIndex).Value) <> "" Then
                MsgBox "El valor no es valido", vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
        
            If pstrTipoDevolucion = strCodTipo And Trim(grdDetalle.Columns(ColIndex).Value) <> "" Then
                
                FchVenc = objODev.FechaVencimiento(grdDetalle.Columns(sCol.COD_PRODUCTO).Value, _
                                                   Format(lblFchVigencia.Caption, "YYYY"), _
                                                   Format(lblFchVigencia.Caption, "MM"))
            
                If FchVenc <> vbNullString Then
                    If CDate(Format(Trim(grdDetalle.Columns(ColIndex).Value), "MM/YYYY")) <> CDate(Format(FchVenc, "MM/YYYY")) Then
                        MsgBox "La fecha ingresada no coincide con el cronograma.", vbCritical, "Error"
                        Cancel = True
                        Exit Sub
                    End If
                End If
            
            End If
        
        
    End Select
End Sub

Private Sub grdDetalle_AfterColUpdate(ByVal ColIndex As Integer)
    grdDetalle.MovePrevious
    grdDetalle.MoveNext
End Sub

Private Sub grdDetalle_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    On Error GoTo Control
    
    Select Case col
        Case sCol.CTD_PRODUCTO_DEV, sCol.CTD_PRODUCTO_FRAC_DEV, sCol.NRO_LOTE, sCol.FCH_VENCIMIENTO
            Select Case Condition
                Case CellStyleConstants.dbgMarqueeRow
                    CellStyle.Font.Bold = True
                Case CellStyleConstants.dbgMarqueeRow + CellStyleConstants.dbgCurrentCell
                    CellStyle.BackColor = vbCyan  'RGB(60, 255, 255)  'vbInfoBackground
                    CellStyle.Font.Bold = True
            End Select
    End Select

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar_Click
    End If
End Sub

