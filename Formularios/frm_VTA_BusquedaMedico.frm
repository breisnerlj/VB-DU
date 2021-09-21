VERSION 5.00
Begin VB.Form frm_VTA_BusquedaMedico 
   Caption         =   "Búsqueda Médico"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   ScaleHeight     =   4395
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7815
      Begin vbp_Ventas.ctlTextBox txtCmp 
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
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
      Begin vbp_Ventas.ctlGrilla grdlListaMedico 
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4471
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Command1"
         Height          =   495
         Left            =   1200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "CMP:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_VTA_BusquedaMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim objMedico As New clsMedico

Public vNumCMP_Ingresado As String


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    
    txtCmp.Text = vNumCMP_Ingresado
    
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    SeteaGrilla
    
    Set grdlListaMedico.DataSource = objMedico.ListaPorCmp(vNumCMP_Ingresado)
    
    'grdlListaMedico.SetFocus
    
    
End Sub


Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant

    arrCampos = Array("NUM_CMP", "NOMBRE", "DES_TIPO_COLEGIO", "COD_MEDICO")
    arrCaption = Array("CMP", "Nombre", "Colegio", "CodigoMedico")
    arrAncho = Array(900, 3000, 3000, 900)
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
    grdlListaMedico.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    Dim i%
    For i = 0 To grdlListaMedico.Columns.Count - 1
        grdlListaMedico.Columns(i).Visible = False
    Next i
    
    grdlListaMedico.Columns("NUM_CMP").Visible = True
    grdlListaMedico.Columns("NOMBRE").Visible = True
    grdlListaMedico.Columns("DES_TIPO_COLEGIO").Visible = True
    
    
End Sub

Private Sub grdlListaMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyReturn Then
        frmPedido.lblMedico.Caption = grdlListaMedico.Columns("NOMBRE").Value
        objMedico.vCodMedico = grdlListaMedico.Columns("COD_MEDICO").Value
        Unload Me
   
   End If
   
   
End Sub



