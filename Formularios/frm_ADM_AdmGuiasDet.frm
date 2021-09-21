VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_ADM_AdmGuiasDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmación de Recepción"
   ClientHeight    =   8070
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Agregar Producto:"
      Height          =   1575
      Left            =   0
      TabIndex        =   19
      Top             =   2040
      Width           =   10455
      Begin MSMask.MaskEdBox txtFecVen 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   7
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin vbp_Ventas.ctlTextBox txtFrac 
         Height          =   255
         Left            =   8040
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Tipo            =   3
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
      Begin vbp_Ventas.ctlTextBox txtUnid 
         Height          =   255
         Left            =   8040
         TabIndex        =   4
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Tipo            =   3
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
      Begin vbp_Ventas.ctlTextBox txtNroLote 
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Tipo            =   2
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
      Begin vbp_Ventas.ctlTextBox txtCodProducto 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
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
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Agregar"
         Height          =   615
         Left            =   9480
         Picture         =   "frm_ADM_AdmGuiasDet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   360
         TabIndex        =   25
         Top             =   405
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Lote:"
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
         Left            =   360
         TabIndex        =   24
         Top             =   765
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fch.Venc:"
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
         Left            =   360
         TabIndex        =   23
         Top             =   1125
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid. Recibidas:"
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
         Left            =   6600
         TabIndex        =   22
         Top             =   765
         Width           =   1425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frac. Recibidas"
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
         Left            =   6600
         TabIndex        =   21
         Top             =   1125
         Width           =   1350
      End
      Begin VB.Label lblDesProducto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-----"
         Height          =   195
         Left            =   3000
         TabIndex        =   20
         Top             =   405
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   10515
      Begin VB.Label lblDesObs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         TabIndex        =   18
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lblFchEmision 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/01/2000"
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
         TabIndex        =   17
         Top             =   600
         Width           =   930
      End
      Begin VB.Label lblDestino 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "005"
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
         Left            =   6960
         TabIndex        =   16
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblOrigen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDI"
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
         Left            =   6960
         TabIndex        =   15
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblNroGuia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-9999999"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
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
         TabIndex        =   13
         Top             =   960
         Width           =   1335
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
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Local Destino:"
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
         Left            =   5520
         TabIndex        =   11
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Local Origen:"
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
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Guía:"
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
         TabIndex        =   9
         Top             =   240
         Width           =   915
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1058
      ModoBotones     =   6
      EnabledEfecto   =   0   'False
   End
   Begin vbp_Ventas.ctlGrillaArray grdDetalle 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   7646
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
      MultiSelect     =   0
   End
End
Attribute VB_Name = "frm_ADM_AdmGuiasDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrNumGuia As String
Public pstrCodLocal As String
Private xarrDetalle As New XArrayDB
Private objGuia As New clsGuia
Private rsProducto As oraDynaset

Private Enum sCol
    ORDEN
    NUM_GUIA
    NUM_ITEM
    NUM_ITEM_LOTE
    COD_PRODUCTO
    DES_PRODUCTO
    DES_LABORATORIO
    NUM_LOTE
    FCH_VENCIMIENTO
    FLG_FRACCION
    Ctd_Fraccion
    CTD_PRODUCTO
    CTD_PRODUCTO_FRAC
    CTDU
    CTDF
    ISNEW
    ROWID
End Enum

Private Sub CmdAdd_Click()
    Dim Index As Integer
    Dim sRowid As String
    
    On Error GoTo Control
    
    If Trim(txtCodProducto.Text) = vbNullString Then
        MsgBox "Ingrese código de producto", vbCritical, "Error Agregar Detalle"
        txtCodProducto.SetFocus
        Exit Sub
    End If
    If Trim(txtNroLote.Text) = vbNullString Then
        MsgBox "Ingrese número de lote", vbCritical, "Error Agregar Detalle"
        txtNroLote.SetFocus
        Exit Sub
    End If
    If Trim(txtFecVen.Text) = vbNullString Or Not IsDate(txtFecVen.FormattedText) Then
        MsgBox "Ingrese fecha de vencimiento", vbCritical, "Error Agregar Detalle"
        txtFecVen.SetFocus
        Exit Sub
    End If
    If (txtUnid.Text + txtFrac.Text) = 0 Then
        MsgBox "Ingrese cantidad recepcionada", vbCritical, "Error Agregar Detalle"
        txtUnid.SetFocus
        Exit Sub
    End If
    
    sRowid = txtCodProducto.Text & txtNroLote.Text & Format(txtFecVen.FormattedText, "mmyyyy")
    
    Index = xarrDetalle.Find(0, sCol.ROWID, sRowid)
    If Index >= 0 Then
        grdDetalle.Bookmark = Index
        grdDetalle.col = 0
        If MsgBox("El producto ya existe en el detalle. ¿Desea remplazarlo?", vbQuestion + vbYesNo, "Agregar Producto") = vbYes Then
            grdDetalle.Columns(sCol.CTDU).Value = txtUnid.Text
            grdDetalle.Columns(sCol.CTDU).Value = txtFrac.Text
            Limpiar_Form_Add
        End If
        grdDetalle.SetFocus
    Else
        xarrDetalle.InsertRows xarrDetalle.UpperBound(1) + 1, 1
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.ORDEN) = xarrDetalle.UpperBound(1) + 1
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.COD_PRODUCTO) = rsProducto("COD_PRODUCTO").Value
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.DES_PRODUCTO) = rsProducto("DES_PRODUCTO").Value
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.FLG_FRACCION) = rsProducto("FLG_FRACCIONAMIENTO").Value
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.Ctd_Fraccion) = rsProducto("CTD_FRACCIONAMIENTO").Value
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.NUM_LOTE) = txtNroLote.Text
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.FCH_VENCIMIENTO) = txtFecVen.FormattedText
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.CTD_PRODUCTO) = 0
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.CTD_PRODUCTO_FRAC) = 0
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.CTDU) = txtUnid.Text
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.CTDF) = txtFrac.Text
        xarrDetalle(xarrDetalle.UpperBound(1), sCol.ISNEW) = 1
        grdDetalle.Rebind
        
        Limpiar_Form_Add
    End If
    
    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, "Error Agregar Detalle"
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case boton
        Case Nuevo

        Case Modificar

        Case Buscar

        Case tb_Actualizar

        Case Imprimir

        Case tb_Excel

        Case tb_email

        Case Grabar
            Confirmar_Recepcion
        Case Cancelar

        Case Eliminar

        Case salir
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo Control
    lblNroGuia.Caption = pstrNumGuia
    LlenarGrid pstrNumGuia
    SetGrid
    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, "Form Load"
End Sub

Private Sub SetGrid()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim i As Integer
    Dim s As TrueDBGrid70.Split
    
    On Error GoTo Control
    
    '---------------------------------------------------------------
    '-- Detalle
    '---------------------------------------------------------------
    arrCampos = Array("ORDEN", "NUM_GUIA", "NUM_ITEM", "NUM_ITEM_LOTE", _
                      "COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "NUM_LOTE", _
                      "FCH_VENCIMIENTO", "FLG_FRACCION", "CTD_FRACCION", "CTD_PRODUCTO", _
                      "CTD_PRODUCTO_FRAC", "CTDU", "CTDF", "ISNEW", "ROWID")
    arrCaption = Array("Item", "Guia", "Item1", "Item2", _
                       "Código", "Descripción", "Laboratorio", "Nro. Lote", _
                       "Fch. Vencimiento", "Frac.", "CF.", "Unid. Envi.", _
                       "Frac. Envi.", "Unid. Recib.", "Frac. Recib.", "IsNew", _
                       "RowId")
    arrAncho = Array(400, 400, 400, 400, _
                     700, 3000, 4000, 1000, _
                     1000, 700, 700, 650, _
                     650, 650, 650, 400, 4000)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgLeft, dbgLeft, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    
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

        .Columns(sCol.FLG_FRACCION).NumberFormat = "Yes/No"
        .Columns(sCol.FCH_VENCIMIENTO).NumberFormat = "MM/YYYY"

        'Columnas editables
        .Columns(sCol.CTDU).BackColor = vbInfoBackground
        .Columns(sCol.CTDU).DataWidth = 4
        .Columns(sCol.CTDF).BackColor = vbInfoBackground
        .Columns(sCol.CTDF).DataWidth = 4
    End With

    Set s = grdDetalle.Splits.Add(1)

    With grdDetalle.Splits(0)
        .RecordSelectors = False
        .SizeMode = dbgScalable 'dbgNumberOfColumns
        .Size = 2
        .AllowSizing = False
        .MarqueeStyle = dbgHighlightRow
        .Columns(sCol.ORDEN).Visible = True
        .Columns(sCol.COD_PRODUCTO).Visible = True
        .Columns(sCol.DES_PRODUCTO).Visible = True
        .Columns(sCol.NUM_LOTE).Visible = True
        .Columns(sCol.FCH_VENCIMIENTO).Visible = True
        .Columns(sCol.FLG_FRACCION).Visible = True
        .Columns(sCol.Ctd_Fraccion).Visible = True
    End With

    With grdDetalle.Splits(1)
        .RecordSelectors = False
        .SizeMode = dbgNumberOfColumns 'dbgScalable
        .Size = 4
        .AllowSizing = False
        .Columns(sCol.CTD_PRODUCTO).Visible = True
        .Columns(sCol.CTD_PRODUCTO_FRAC).Visible = True
        .Columns(sCol.CTDF).Visible = True
        .Columns(sCol.CTDU).Visible = True
    End With

    grdDetalle.Rebind
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub LlenarGrid(vNroGuia As String)
    Dim rsDetalle As oraDynaset
    Dim i As Integer, intsep As Integer

    On Error GoTo Control
    
    Set rsDetalle = objGuia.ListaDetLote(vNroGuia)
    
    If rsDetalle.RecordCount = 0 Then
'        Call LimpiarDetalle
        Exit Sub
    End If
    
    xarrDetalle.LoadRows rsDetalle.GetRows
    xarrDetalle.AppendColumns 4

    For i = xarrDetalle.LowerBound(1) To xarrDetalle.UpperBound(1)
        xarrDetalle.Value(i, sCol.CTDU) = xarrDetalle.Value(i, sCol.CTD_PRODUCTO)
        xarrDetalle.Value(i, sCol.CTDF) = xarrDetalle.Value(i, sCol.CTD_PRODUCTO_FRAC)
        xarrDetalle.Value(i, sCol.ISNEW) = 0
        xarrDetalle.Value(i, sCol.ROWID) = xarrDetalle.Value(i, sCol.COD_PRODUCTO) & _
                                           xarrDetalle.Value(i, sCol.NUM_LOTE) & _
                                           Format(xarrDetalle.Value(i, sCol.FCH_VENCIMIENTO), "mmyyyy")
    Next i
    
    grdDetalle.Array1 = xarrDetalle
    grdDetalle.Rebind
    
    Set rsDetalle = Nothing
    
    Exit Sub
Control:
    xarrDetalle.AppendColumns 21
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Buscar_Producto(vCodProducto As String)
    Dim objProducto As New clsProducto

    On Error GoTo Control
    
    If Trim(vCodProducto) = vbNullString Then
        Limpiar_Form_Add
        Exit Sub
    End If

    Set rsProducto = objProducto.fnBuscaProducto(vCodProducto)
    If rsProducto.RecordCount > 0 Then
        If rsProducto(0) <> -1 Then
            Ubicar_Producto rsProducto("COD_PRODUCTO").Value
            txtCodProducto.Text = rsProducto("COD_PRODUCTO").Value
            lblDesProducto.Caption = rsProducto("DES_PRODUCTO").Value
            txtUnid.Text = 0
            txtFrac.Text = 0
            txtNroLote.SetFocus
        Else
            MsgBox "Producto no encontrado"
        End If
    End If

    Set objProducto = Nothing
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Ubicar_Producto(vCodProducto As String)
    Dim Index As Integer
    
    On Error GoTo Control

    Index = xarrDetalle.Find(0, sCol.COD_PRODUCTO, vCodProducto)
    If Index >= 0 And grdDetalle.ApproxCount > 0 Then
        grdDetalle.Bookmark = Index
        grdDetalle.col = 0
        grdDetalle.SetFocus
    End If
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xarrDetalle = Nothing
    Set objGuia = Nothing
    Set rsProducto = Nothing
End Sub

Private Sub grdDetalle_AfterColUpdate(ByVal ColIndex As Integer)
    grdDetalle.MovePrevious
    grdDetalle.MoveNext
End Sub

Private Sub grdDetalle_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Select Case ColIndex
        Case sCol.CTDU
            Cancel = False
        Case sCol.CTDF
            Cancel = IIf(grdDetalle.Columns(sCol.FLG_FRACCION).Value = "1", False, True)
        Case sCol.NUM_LOTE
            Cancel = IIf(grdDetalle.Columns(sCol.ISNEW).Value = "1", False, True)
        Case sCol.FCH_VENCIMIENTO
            Cancel = IIf(grdDetalle.Columns(sCol.ISNEW).Value = "1", False, True)
        Case Else
            Cancel = True
    End Select
End Sub

Private Sub grdDetalle_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Select Case ColIndex
        Case sCol.CTDU, sCol.CTDF
            If Not IsNumeric(Trim(grdDetalle.Columns(ColIndex).Value)) And _
                    Trim(grdDetalle.Columns(ColIndex).Value) <> "" Then
                MsgBox "El valor no es valido", vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
        Case sCol.NUM_LOTE
            If Trim(grdDetalle.Columns(ColIndex).Value) = vbNullString Then
                MsgBox "Ingrese Número de Lote", vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
        Case sCol.FCH_VENCIMIENTO
            If Trim(grdDetalle.Columns(ColIndex).Value) = vbNullString Or _
                Not IsDate(Trim(grdDetalle.Columns(ColIndex).Value)) Then
                MsgBox "Ingrese Fecha de vencimiento", vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
    End Select
End Sub

Private Sub grdDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If grdDetalle.Columns(sCol.ISNEW).Value = 1 Then
                If MsgBox("¿Esta seguro que desea eliminar este item del detalle?", vbQuestion + vbYesNo, "Eliminar Item") = vbYes Then
                    grdDetalle.Delete
                End If
            Else
                MsgBox "No puede eliminar este item." + vbCrLf + _
                "Si no recibió el producto colocar 0 en unidades/fraciones recibidas", _
                vbOKOnly + vbInformation, "Eliminar Item"
            End If
    End Select
End Sub

Private Sub Limpiar_Form_Add()
    Set rsProducto = Nothing
    txtCodProducto.Text = vbNullString
    lblDesProducto.Caption = vbNullString
    txtNroLote.Text = vbNullString
    txtFecVen.Mask = vbNullString
    txtFecVen.Text = vbNullString
    txtFecVen.Mask = "##/####"
    txtUnid.Text = vbNullString
    txtFrac.Text = vbNullString
End Sub

'Private Sub LimpiarDetalle()
'    Dim i As Integer
'
'    On Error GoTo Control
'
'    If grdDetalle.ApproxCount < 1 Then Exit Sub
'
'    If MsgBox("¿Seguro(a) de Eliminar Items que no tienen cantidades a devolver?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'        Exit Sub
'    End If
'
'    grdDetalle.MoveFirst
'    While Not grdDetalle.EOF
'        If Val(Trim("" & grdDetalle.Columns(sCol.CTD_PRODUCTO_DEV).Value)) = 0 And Val(Trim("" & grdDetalle.Columns(sCol.CTD_PRODUCTO_FRAC_DEV).Value)) = 0 Then
'            grdDetalle.Delete
'        Else
'            grdDetalle.MoveNext
'        End If
'    Wend
'    grdDetalle.col = 0
'    grdDetalle.MoveFirst
'
'    Exit Sub
'Control:
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Sub

Private Sub txtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Set rsProducto = Nothing
        Buscar_Producto Trim(txtCodProducto.Text)
    End If
End Sub

Private Sub txtFecVen_Validate(Cancel As Boolean)
    If IsDate(txtFecVen.FormattedText) = False Then
       'Opcional: podemos mostrar un mensaje
       MsgBox " La Fecha no es válida ", vbCritical, " Error al ingresar la fecha "
       Cancel = True
    End If
End Sub

Private Sub Confirmar_Recepcion()
    On Error GoTo Control
    If MsgBox("¿Desea Confirmar la Recepción la Guía Nro ->" & pstrNumGuia & "<- ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        'strError = objGuia.Recepciona(objUsuario.CodigoEmpresa, strCodLocal, strNumGuia, objUsuario.Codigo)
        If objGuia.RecepcionaNew(objUsuario.CodigoEmpresa, pstrCodLocal, pstrNumGuia, objUsuario.Codigo, xarrDetalle) = True Then
            MsgBox "Se confirmó la recepción", vbInformation + vbOKOnly
            Unload Me
        End If
        'If strError <> "" Then Err.Raise 1, "", strError
        'sub_Actualizar strNumGuia
    End If
    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, "Form Load"
End Sub
