VERSION 5.00
Begin VB.Form frm_ADM_ProdLoteSobrFalt 
   Caption         =   "Ingresar Lote"
   ClientHeight    =   1755
   ClientLeft      =   4485
   ClientTop       =   5460
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4890
   Visible         =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtLote 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   600
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
   Begin vbp_Ventas.ctlDataCombo CboLote 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
   End
   Begin VB.Label lblNEntrega 
      Caption         =   "lblNEntrega"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblGuia 
      Caption         =   "lblGuia"
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblCodigo 
      Caption         =   "lblCodigo"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblRegistro 
      Caption         =   "lblRegistro"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Producto :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Lote :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblProducto 
      Caption         =   "AAA"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frm_ADM_ProdLoteSobrFalt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objProducto As New clsProducto
Dim objEntrega As New clsEntrega


Public Sub cargaDatos(ByVal idRegistro As String, ByVal codProducto As String)
    Dim indicador As String
    lblCodigo.Caption = codProducto
    lblRegistro.Caption = idRegistro
'    indicador = objproducto.DevIndicadorLote(lblCodigo.Caption)
'    If indicador = "S" Then
        Set CboLote.RowSource = objProducto.ListaLote(lblCodigo.Caption, objUsuario.CodigoLocal)
        'Set CboLote.RowSource = objproducto.ListaLoteXFacXRc(lblRegistro.Caption, lblCodigo.Caption, objUsuario.CodigoLocal)
        CboLote.ListField = "NLOTE"
        CboLote.BoundColumn = "FVENC"
        Me.Show vbModal
'    End If
End Sub

Private Sub CboLote_Change()
    If CboLote.Text = "NO TIENE N LOTE" Then
        txtLote.Visible = True
        txtLote.Enabled = True
        txtLote.Text = ""
        
        CboLote.Visible = False
    Else
        txtLote.Visible = True
        txtLote.Text = CboLote.Text
        txtLote.Enabled = False
    End If
        
End Sub

Private Sub cmdAceptar_Click()
    If Trim(txtLote.Text) <> "" Then
        If MsgBox("¿Esta seguro de registrar este lote?", vbYesNo + vbCritical, Aviso) = vbYes Then
            Call objEntrega.EditaLoteEnSobrantes(lblRegistro.Caption, txtLote.Text, lblCodigo.Caption, lblGuia.Caption, lblNEntrega.Caption)
            MsgBox "Se grabo el Número de Lote", vbCritical, "AVISO"
            Unload Me
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

