VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ADM_BuscaProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Producto"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin vbp_Ventas.ctlGrilla grdProductos 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7011
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1111
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar"
            Key             =   "Buscar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1440
         TabIndex        =   3
         Top             =   0
         Width           =   5415
         Begin vbp_Ventas.ctlTextBox txtDesProducto 
            Height          =   375
            Left            =   960
            TabIndex        =   0
            Top             =   120
            Width           =   4455
            _ExtentX        =   7858
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   840
         End
      End
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":0000
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":059A
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":0B34
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":10CE
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":1668
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":1C02
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":219C
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":2736
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":2CD0
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":326A
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":3804
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BuscaProd.frx":3D9E
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      Caption         =   "Registros: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4845
      Width           =   1035
   End
End
Attribute VB_Name = "frm_ADM_BuscaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodProducto As String
Dim objProducto As New clsProducto

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    SeteaGrilla
End Sub

Private Sub grdProductos_DblClick()
    If grdProductos.ApproxCount = 0 Then Exit Sub
    strCodProducto = grdProductos.Columns("COD_PRODUCTO").Value
    Unload Me

End Sub

Private Sub grdProductos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then grdProductos_DblClick
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Buscar
        
    Case 2
        Unload Me
End Select

End Sub




Private Sub Buscar()
    Set grdProductos.DataSource = objProducto.Lista(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"), Trim(txtDesProducto.Text), "", "", objUsuario.CodigoLocal)


    
End Sub


Private Sub SeteaGrilla()
On Error GoTo handle
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim columna As TrueDBGrid70.Column

  
                      
                      
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO")
                      
                      
    arrCaption = Array("Código", "Descripción")
    
    arrAncho = Array(1000, 5000)
                     
    arrAlineacion = Array(dbgCenter, dbgLeft)
                          
    grdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    
    
    
    
    For Each columna In grdProductos.Columns
        columna.AllowSizing = False
    
    Next
    

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


Private Sub txtDesProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Buscar
End Sub
