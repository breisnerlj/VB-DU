VERSION 5.00
Begin VB.Form frm_ADM_ProdDevolucion 
   Caption         =   "Lectura Producto(s) Politica de Canje"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdF11 
      Caption         =   "[F11] Finalizar"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "[F2] Eliminar Producto"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "[Esc] Salir"
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdF1 
      Caption         =   "[F1] Buscar Producto"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   6120
      Width           =   1935
   End
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9551
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
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
      Caption         =   "Ingrese Codigo de Barras o Producto:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frm_ADM_ProdDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objEntrega As New clsEntrega
Public strIdEntrega As String
Public strTipo As String
Dim glMomento As String
Dim xarrDetalle As XArrayDB

Private Sub cmdEsc_Click()
    Unload Me
End Sub

Private Sub cmdF1_Click()
    frm_ADM_Devolucion_LstProd.recibe strIdEntrega, glMomento
    frm_ADM_Devolucion_LstProd.Show vbModal
End Sub

Private Sub cmdF11_Click()
If Me.ctlGrilla1.ApproxCount > 0 Then
    frm_ADM_ProdDevGuia.strTipo = strTipo
    frm_ADM_ProdDevGuia.strIdEntrega = strIdEntrega
    
    Dim strCadCodProducto, strCadCtdProducto, strCadNumLote, strCadFchVenc, strCadCtdFProducto, strCadCodFacSap As String
    
    Dim i As Integer
    i = 0
    Me.ctlGrilla1.MoveFirst
    
    While i < Me.ctlGrilla1.ApproxCount
        strCadCodProducto = strCadCodProducto & Me.ctlGrilla1.Columns("COD_PRODUCTO").Value & "|"
        strCadCtdProducto = strCadCtdProducto & Me.ctlGrilla1.Columns("CANTIDAD").Value & "|"
        strCadNumLote = strCadNumLote & UCase(Me.ctlGrilla1.Columns("LOTE").Value) & "|"
        strCadFchVenc = strCadFchVenc & Replace(Me.ctlGrilla1.Columns("FCH_VENC").Value, "__/__/____", "") & "|"
        strCadCtdFProducto = strCadCtdFProducto & "0|"
        strCadCodFacSap = strCadCodFacSap & Me.ctlGrilla1.Columns("COD_FAC_SAP").Value & "|"
        Me.ctlGrilla1.MoveNext
        i = i + 1
    Wend
    frm_ADM_ProdDevGuia.strCadCodProducto = strCadCodProducto
    frm_ADM_ProdDevGuia.strCadCtdProducto = strCadCtdProducto
    frm_ADM_ProdDevGuia.strCadNumLote = strCadNumLote
    frm_ADM_ProdDevGuia.strCadFchVenc = strCadFchVenc
    frm_ADM_ProdDevGuia.strCadCtdFProducto = strCadCtdFProducto
    'Agregado 27/11/2012
    frm_ADM_ProdDevGuia.strCadCodFacSap = strCadCodFacSap
    frm_ADM_ProdDevGuia.Show vbModal
Else
    MsgBox "No existen productos disponibles", vbCritical + vbInformation, "Aviso"
End If

End Sub

Private Sub cmdF2_Click()
Dim msbo As Variant
msgbo = MsgBox("¿Seguro que desea eliminar el producto?", vbYesNo + vbInformation, App.ProductName)
If msgbo = vbYes Then
    objEntrega.EliminaDetDev strIdEntrega, Me.ctlGrilla1.Columns("COD_PRODUCTO").Value
    recibe
End If
End Sub


Private Sub ctlTextBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
     Dim rs As oraDynaset
     Dim strCodido As String
     Dim strDesripcion As String
     Dim objProducto As New clsProducto
     strCodido = ""
    Set rs = objProducto.ListaBusqueda(Trim(Me.ctlTextBox1.Text))
    If Not rs.EOF Then
        strCodido = rs("COD").Value
        strDesripcion = rs("DES").Value
    Else
        MsgBox "El codigo de producto no existe", vbCritical, App.ProductName
        Exit Sub
    End If
    
    
    Dim odyn As oraDynaset
    Set odyn = objEntrega.ListaProdDev(strIdEntrega, strCodido)

    
'    ODyn.MoveFirst
'    While Not ODyn.EOF
'        If ODyn("COD_PRODUCTO") = Me.ctlTextBox1.Text Then
'            objentrega.GrabaDetDev strIdEntrega, "" & ODyn("COD_PRODUCTO"), "1", "", "", glMomento
'        End If
'        ODyn.MoveNext
'    Wend
'    Me.ctlGrilla1.Limpiar
'    Set Me.ctlGrilla1.DataSource = objentrega.ListaDetDev(strIdEntrega, glMomento)
'    ctlTextBox1.Text = ""
'    If Me.ctlGrilla1.ApproxCount > 0 Then
        frm_ADM_ProdDevCant.idEntrega = strIdEntrega
        frm_ADM_ProdDevCant.codProducto = strCodido '"" & ODyn("COD_PRODUCTO") 'Me.ctlGrilla1.Columns("COD_PRODUCTO").Value
        frm_ADM_ProdDevCant.lblProducto = strDesripcion '"" & ODyn("DES_PRODUCTO") ' Me.ctlGrilla1.Columns("DES_PRODUCTO").Value
        frm_ADM_ProdDevCant.moment = glMomento
        'frm_ADM_ProdDevCant.txtCantidad = ""
        frm_ADM_ProdDevCant.Show vbModal
'    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    cmdF1_Click
End If
If KeyCode = vbKeyF2 Then
    cmdF2_Click
End If
If KeyCode = vbKeyF11 Then
    cmdF11_Click
End If
If KeyCode = vbKeyEscape Then
    cmdEsc_Click
End If
If KeyCode = 107 Then
    frm_ADM_ProdDevCant.lblProducto.Caption = Me.ctlGrilla1.Columns("DES_PRODUCTO").Value
    frm_ADM_ProdDevCant.idEntrega = strIdEntrega
    frm_ADM_ProdDevCant.codProducto = Me.ctlGrilla1.Columns("COD_PRODUCTO").Value
    frm_ADM_ProdDevCant.Show vbModal
End If
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "CANTIDAD", "ITEM", "FLG_BLOQUEADO", "LOTE", "fch_venc", "cod_fac_sap")
    arrCaption = Array("Codigo", "Producto", "Laboratorio", "Cantidad", "item", "bloqueado", "Lote", "Fec. Venc.", "Factura")
    arrAncho = Array(800, 3100, 3100, 500, 500, 500, 1000, 1000, 800)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgRight)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    Me.ctlGrilla1.Columns("item").Visible = False
    Me.ctlGrilla1.Columns("bloqueado").Visible = False
End Sub

Private Sub Form_Load()
    SeteaGrilla
    Me.ctlGrilla1.Limpiar
    Dim odyn As oraDynaset
    Set odyn = objEntrega.momentoDetDev(strIdEntrega)
    Dim momento As String
    momento = "" & odyn("MOMENTO")
    glMomento = momento
    Set Me.ctlGrilla1.DataSource = objEntrega.ListaDetDev(strIdEntrega, momento)
End Sub

Public Sub recibe()
    Me.ctlGrilla1.Limpiar
    SeteaGrilla
    Set Me.ctlGrilla1.DataSource = objEntrega.ListaDetDev(strIdEntrega, glMomento)
End Sub

Public Sub load()
Me.Show vbModal
Me.ctlTextBox1.SetFocus
End Sub

