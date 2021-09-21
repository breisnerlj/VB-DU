VERSION 5.00
Begin VB.Form frm_VTA_BuscaDiagnostico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Diagnostico"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtCriterio 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
   Begin vbp_Ventas.ctlGrilla grdDiagnostico 
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4471
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "frm_VTA_BuscaDiagnostico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objConvenio As New clsConvenio
Public OUTPUT_CODIGO As String
Public OUTPUT_NOMBRE As String

Private Sub Form_Load()
    SeteaGrilla
End Sub

Private Sub cmdBuscar_Click()
    Set grdDiagnostico.DataSource = objConvenio.ListaDignostico(Trim(txtCriterio.Text))
End Sub

Private Sub grdDiagnostico_DblClick()
    If grdDiagnostico.ApproxCount > 0 Then
        OUTPUT_CODIGO = grdDiagnostico.DataSource("COD_DIAGNOSTICO").Value
        OUTPUT_NOMBRE = grdDiagnostico.DataSource("DES_DIAGNOSTICO").Value
        Unload Me
    End If
End Sub

Private Sub grdDiagnostico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then grdDiagnostico_DblClick
End Sub

Sub SeteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_DIAGNOSTICO", "COD_CIE_10", "DES_DIAGNOSTICO")
    arrCaption = Array("Codigo", "CIE 10", "Diagnostico")
    arrAncho = Array(1000, 1000, 2800)
    arrAlineacion = Array(dbgCenter, vbAlignLeft, vbAlignLeft)
    grdDiagnostico.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDiagnostico.Columns("COD_DIAGNOSTICO").Visible = False
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objConvenio = Nothing
End Sub
