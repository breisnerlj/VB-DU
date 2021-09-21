VERSION 5.00
Begin VB.Form frm_ADM_PVM_Historico 
   Caption         =   "Historial del Promedio de Ventas Mensual (PVM)"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "[Esc] Cerrar"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "HISTORICO DE PVM"
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9135
      Begin vbp_Ventas.ctlGrilla ctlGrilla1 
         Height          =   2775
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4895
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.Label lblDesLaboratorio 
         Caption         =   "lblDesLaboratorio"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label lblDesProducto 
         Caption         =   "lblDesProducto"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Laboratorio :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Producto :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_ADM_PVM_Historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPvm As New clsSPVM
Public codProducto As String

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Command2_Click
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    SeteaGrilla
    Set Me.ctlGrilla1.DataSource = objPvm.HistorialPVM(objUsuario.CodigoLocal, codProducto)
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub SeteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant
    arrCampos = Array("CTD_PVM_SOLICITADO", "CTD_PVM_APROBADO", "USU_REG", "FCH_USU_REG", "ORIGEN")
    arrCaption = Array("PVM Solicitado", "PVM Aprobado", "Registrado por", "Fecha Registro", "Lugar")
    arrAncho = Array(1500, 1500, 1500, 1500, 1000)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft)
    arrFoco = Array(False, False, False, False, False)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
End Sub
