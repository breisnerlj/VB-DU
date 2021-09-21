VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_ProdDevCant 
   Caption         =   "Ingreso de Cantidad"
   ClientHeight    =   2715
   ClientLeft      =   6390
   ClientTop       =   6990
   ClientWidth     =   4890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4890
   Begin VB.CommandButton Command2 
      Caption         =   "[Esc] Cerrar"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "[F11] Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
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
   Begin vbp_Ventas.ctlTextBox txtLote 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
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
   Begin MSComCtl2.DTPicker dtpFch 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   59899907
      CurrentDate     =   40813
   End
   Begin vbp_Ventas.ctlDataCombo CboLote 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
   End
   Begin VB.Label lblProducto 
      Caption         =   "AAA"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Fec. Vcmto. :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Lote :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Producto :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm_ADM_ProdDevCant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objEntrega As New clsEntrega
Public idEntrega As String
Public codProducto As String
Public moment As String
Dim objProducto As New clsProducto



Private Sub CboLote_Change()
    If CboLote.Text = "NO TIENE N LOTE" Then
        txtLote.Visible = True
        txtLote.Enabled = True
        txtLote.Text = ""
        txtLote.SetFocus
        CboLote.Visible = False
    Else
        txtLote.Visible = True
        txtLote.Enabled = False
        txtLote.Text = CboLote.Text
        If CboLote.Text = "" Then
            txtLote.Enabled = True
            CboLote.Visible = False
        End If
    End If
End Sub

Private Sub Command1_Click()
If validaForm Then
    If Val(Me.txtCantidad.Text) <> 0 Then
        Dim odyn As oraDynaset
        Set odyn = objEntrega.MaxStockProd(idEntrega, codProducto)
        odyn.MoveFirst
        Dim maxStock As String
        maxStock = "" & odyn("maxstock").Value
            If Val(Me.txtCantidad.Text) > Val(maxStock) Then
                MsgBox "Ha excedido el stock maximo para devolver.", vbCritical + vbInformation, "Aviso"
                Me.txtCantidad.Text = ""
                Me.txtCantidad.SetFocus
            Else
                'objentrega.EditaDetDev idEntrega, codProducto, Me.txtCantidad.Text, Me.txtLote.Text, Me.dtpFch.Value ' Me.txtFecVen.Text
                objEntrega.GrabaDetDev idEntrega, codProducto, Me.txtCantidad.Text, Me.txtLote.Text, Me.dtpFch.Value, moment
                frm_ADM_ProdDevolucion.strIdEntrega = idEntrega
                frm_ADM_ProdDevolucion.recibe
                Unload Me
            End If
    Else
        Dim msbo As Variant
        msgbo = MsgBox("¿Seguro que desea eliminar el producto?", vbYesNo + vbInformation, App.ProductName)
        If msgbo = vbYes Then
            objEntrega.EliminaDetDev idEntrega, codProducto
            frm_ADM_ProdDevolucion.strIdEntrega = idEntrega
            frm_ADM_ProdDevolucion.recibe
            Unload Me
        End If
    End If
End If
End Sub

Function buscaCantidad(grilla As ctlGrilla, codProd As String)

End Function

Function validaForm() As Boolean
    validaForm = True
    If Me.txtCantidad.Text = "" Then
        validaForm = False
        MsgBox "Ingrese Cantidad", vbCritical + vbInformation, "Aviso"
        Me.txtCantidad.SetFocus
        Exit Function
    End If
    If Me.txtLote.Text = "" Then
        validaForm = False
        MsgBox "Ingrese Lote", vbCritical + vbInformation, "Aviso"
        Me.txtLote.SetFocus
        Exit Function
    End If
'    If Me.txtFecVen.Text = "" Then
'        validaForm = False
'        MsgBox "Ingrese Fecha de Vencimiento", vbCritical + vbInformation, "Aviso"
'        Me.txtFecVen.SetFocus
'        Exit Function
'    End If
End Function

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF11 Then
        Command1_Click
    End If
    If KeyCode = vbKeyEscape Then
        Command2_Click
    End If
End Sub

Public Sub load()
Me.Show vbModal
Me.txtCantidad.SetFocus
End Sub

Private Sub txtCantidad_GotFocus()
            'If objProducto.DevIndicadorLote(codProducto) = "S" Then
                Set CboLote.RowSource = objProducto.ListaLoteXFacXRc(idEntrega, codProducto, objUsuario.CodigoLocal)  'objUsuario.ListaUsuarioDLV
                CboLote.ListField = "NLOTE"
                CboLote.BoundColumn = "FVENC"
            'Else
            '    CboLote.Visible = False
            '    txtLote.Visible = True
            '    txtLote.Text = "NOTHING"
            '    txtLote.Enabled = False
            'End If
                
End Sub

