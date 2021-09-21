VERSION 5.00
Begin VB.Form frm_VTA_CobroRespoProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producto a "
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7230
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5880
      Picture         =   "frm_VTA_CobroRespoProd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4200
      Picture         =   "frm_VTA_CobroRespoProd.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrillaArray grdProductosCobro 
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingresar los datos del documento a favor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   5160
         Picture         =   "frm_VTA_CobroRespoProd.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin vbp_Ventas.ctlTextBox txtNumDocumento 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         Tipo            =   7
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
      Begin vbp_Ventas.ctlDataCombo dbcTipoDoc 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número de Documento"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1020
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento"
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
         Left            =   510
         TabIndex        =   5
         Top             =   540
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frm_VTA_CobroRespoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento
Dim DescCol As TrueDBGrid70.Column


Public Sub datos(ByVal Caption As String)



Me.Caption = Caption

Me.Show vbModal

End Sub

Private Sub cmdAceptar_Click()
Dim bolMal As Boolean
    If grdProductosCobro.ApproxCount = 0 Then Exit Sub
    grdProductosCobro.Update
    
            bolMal = False
            Dim i As Integer
            Dim intCheck As Integer
            intCheck = 0
            For i = 0 To objVenta.ProductoCobro.UpperBound(1)
                If objVenta.ProductoCobro(i, 0) = "-1" Then intCheck = intCheck + 1
            Next
            
            If intCheck = 0 Then
                MsgBox "Seleccionar algun producto", vbInformation, App.ProductName
                Exit Sub
            End If
            
            If intCheck > 1 Then
                MsgBox "Se debe seleccionar 1 solo producto", vbInformation, App.ProductName
                Exit Sub
            End If
            
            For i = 0 To objVenta.ProductoCobro.UpperBound(1)
                If objVenta.ProductoCobro(i, 0) = "-1" Then
                    If Val(objVenta.ProductoCobro(i, 10)) > Val(objVenta.ProductoCobro(i, 7)) Then
                        MsgBox "El importe es mayor que el monto del documento", vbInformation, App.ProductName
                        bolMal = True
                        Exit For
                    Else
                        frm_VTA_CobroXResponsabilidad.txtImporteAFavor.Text = Format(objVenta.ProductoCobro(i, 10), "0.00")
                        
                        Dim intCant As Integer
                        Dim flgFracc As String
                        
                        intCant = 0
                        flgFracc = "0"
                        
                        
                        If Val(grdProductosCobro.Columns(8).Value) > 0 And Val(grdProductosCobro.Columns(9).Value) = 0 Then
                            intCant = Val(grdProductosCobro.Columns(8).Value)
                            flgFracc = "0"
                        
                        End If
                        
                        If Val(grdProductosCobro.Columns(8).Value) = 0 And Val(grdProductosCobro.Columns(9).Value) > 0 Then
                            intCant = Val(grdProductosCobro.Columns(9).Value)
                            flgFracc = "1"
                        End If
                        
                        If Val(grdProductosCobro.Columns(8).Value) > 0 And Val(grdProductosCobro.Columns(9).Value) > 0 Then
                            intCant = (Val(grdProductosCobro.Columns(8).Value) * Val(grdProductosCobro.Columns(6).Value)) + Val(grdProductosCobro.Columns(9).Value)
                            flgFracc = "1"
                        End If
                        
                        
                        
                        
                        
                        frmPedido.grdPedido.Limpiar
                        frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(grdProductosCobro.Columns(1).Value, _
                                                    grdProductosCobro.Columns(2).Value, _
                                                    intCant, _
                                                    flgFracc, _
                                                    -Val(grdProductosCobro.Columns(10).Value), _
                                                    objVenta.CodigoTipoVenta, _
                                                    0, , , , , , "", _
                                                    0)
                        frmPedido.Cal_Montos
                        frmPedido.grdPedido.Rebind
                        
                        
                        
                    End If
                    
                End If
            Next
    
    If Not bolMal Then Unload Me
    
    
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo CtrlErr
    
    
    objVenta.CargarProductoCobro dbcTipoDoc.BoundText, Replace(txtNumDocumento.Text, "-", "")
    
    If objVenta.ProductoCobro.UpperBound(1) = -1 Then MsgBox "No se encontro el Documento " & dbcTipoDoc.BoundText & " - " & txtNumDocumento.Text, vbInformation + vbOKOnly, App.ProductName
    
    
    grdProductosCobro.Array1 = objVenta.ProductoCobro
    
    grdProductosCobro.Col = 6
    grdProductosCobro.SetFocus
    
    grdProductosCobro.Rebind
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    

    Set dbcTipoDoc.RowSource = objDocumento.ListaTipo(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
    dbcTipoDoc.ListField = "DESCRIPCION"
    dbcTipoDoc.BoundColumn = "CODIGO"
    
    
    Call SetteaGrd
    
    Set DescCol = grdProductosCobro.Columns(11)
End Sub


Private Sub SetteaGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim arrFoco As Variant


arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "")
arrCaption = Array("Sel", "Código", "Descripción", "Cant. Und.", "Cant. Fracc.", "Prec. Unit.", "Cant. Fracc.", "Monto Sub Total", "Unidades", "Fraccion", "Importe", "des")
arrAncho = Array(800, 800, 2500, 200, 500, 700, 800, 500, 800, 800, 800, 800)
arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgCenter, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight)
arrFoco = Array(True, False, False, False, False, False, False, False, True, True, False, False)
grdProductosCobro.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
grdProductosCobro.HeadLines = 2
grdProductosCobro.AllowUpdate = True
grdProductosCobro.Columns(0).ValueItems.Presentation = dbgCheckBox
grdProductosCobro.Columns(0).ValueItems.CycleOnClick = True
grdProductosCobro.Columns(3).Visible = False
grdProductosCobro.Columns(4).Visible = False
grdProductosCobro.Columns(5).Visible = False
grdProductosCobro.Columns(6).Visible = False
grdProductosCobro.Columns(7).Visible = False
grdProductosCobro.Columns(11).Visible = False
grdProductosCobro.Columns(8).EditMask = "####"
grdProductosCobro.Columns(9).EditMask = "####"
grdProductosCobro.CellTips = 1
grdProductosCobro.CellTipsWidth = 5000
grdProductosCobro.MarqueeStyle = dbgHighlightCell



'    arrCampos = Array("", "", "", "", "")
'    arrCaption = Array("Tipo", "Descripción", "Modalidad", "Descripción", "Activo")
'    arrAncho = Array(800, 2000, 800, 2500, 900)
'    arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgLeft, dbgCenter)
'    arrFoco = Array(False, False, False, False, True)
'
'
'    grdModalidadVenta.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
'    grdModalidadVenta.AllowUpdate = True
'    grdModalidadVenta.Columns(4).ValueItems.Presentation = dbgCheckBox
'    grdModalidadVenta.Columns(4).ValueItems.CycleOnClick = True
'    grdModalidadVenta.Columns(0).Merge = True
'    grdModalidadVenta.Columns(1).Merge = True
'    grdModalidadVenta.MarqueeStyle = dbgHighlightCell





End Sub





Private Sub grdProductosCobro_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 8 Or ColIndex = 9 Then
            grdProductosCobro.Columns(10).Value = Round(((Val(grdProductosCobro.Columns(8).Value) * Val(grdProductosCobro.Columns(6).Value)) + Val(grdProductosCobro.Columns(9).Value)) * (Val(grdProductosCobro.Columns(5).Value) / Val(grdProductosCobro.Columns(6).Value)), 2)
            grdProductosCobro.Update
            grdProductosCobro.Rebind
    End If
End Sub

Private Sub grdProductosCobro_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueDBGrid70.StyleDisp)
    If RowIndex > -1 Then
        CellTip = DescCol.CellText(grdProductosCobro.RowBookmark(RowIndex))
    End If
End Sub

