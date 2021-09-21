VERSION 5.00
Begin VB.Form frm_VTA_ConsPetitorio 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   5760
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla grdProductos 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8705
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtBuscar 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
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
   Begin vbp_Ventas.ctlDataCombo dbcMotivo 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   855
      Left            =   3000
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Convenio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   180
      Picture         =   "frm_VTA_ConsPetitorio.frx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Busqueda"
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
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Productos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1710
   End
End
Attribute VB_Name = "frm_VTA_ConsPetitorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBuscar_Click()
    Dim objProducto As New clsProducto
    Set grdProductos.DataSource = objProducto.ConsultaPetitorio(Trim(txtBuscar.Text), dbcMotivo.BoundText)
    Set objProducto = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Control

    SetteaFormulario Me
    Dim objConvenio As New clsConvenio
    Set dbcMotivo.RowSource = objConvenio.ListaconPetitorio
    dbcMotivo.ListField = "DES_CONVENIO"
    dbcMotivo.BoundColumn = "COD_CONVENIO"
    SetteaGrd
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub SetteaGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim arrFoco As Variant

    arrCampos = Array("Codigo", "Descripcion", "COD_PETITORIO")
    arrCaption = Array("Codigo", "Descripción", "")
    arrAncho = Array(800, 5500, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter)
    arrFoco = Array(False, False, False)

    grdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    grdProductos.Columns(2).Visible = False
End Sub

