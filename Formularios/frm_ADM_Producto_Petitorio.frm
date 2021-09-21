VERSION 5.00
Begin VB.Form frm_ADM_Producto_Petitorio 
   BorderStyle     =   0  'None
   Caption         =   "Ingreso de producto a petitorio"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   Icon            =   "frm_ADM_Producto_Petitorio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlTextBox txtCodigoGraba 
      Height          =   855
      Left            =   1800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
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
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el convenio"
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Width           =   6555
      Begin vbp_Ventas.ctlDataCombo CboConvenio 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo CboPetitorio 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   1020
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Convenio"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Petitorio"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1050
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_ADM_Producto_Petitorio.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_ADM_Producto_Petitorio.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6555
      Begin vbp_Ventas.ctlTextBox TxtProducto 
         Height          =   345
         Left            =   960
         TabIndex        =   0
         Top             =   300
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
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
         Caption         =   "Producto"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agregar producto a petitorio"
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
      Index           =   4
      Left            =   420
      TabIndex        =   12
      Top             =   60
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_ADM_Producto_Petitorio.frx":0B20
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   11
      Left            =   6090
      TabIndex        =   8
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   12
      Left            =   4380
      TabIndex        =   7
      Top             =   6900
      Width           =   1215
   End
End
Attribute VB_Name = "frm_ADM_Producto_Petitorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPetitorio As New clsPetitorio
Dim CodigoPetitorio As String
Private Sub Form_Load()
    setteaFormulario Me
    CodigoPetitorio = objVenta.ParametroValor("CODPETICOM")
    txtCodigoGraba.Visible = False
    CargaCombos
End Sub

Sub CargaCombos()
    On Error GoTo CtrlErr
    Set CboConvenio.RowSource = objPetitorio.Lista_Convenio(CodigoPetitorio)
    
   CboConvenio.ListField = "DES_CONVENIO"
   CboConvenio.BoundColumn = "COD_CONVENIO"
   
   Exit Sub
CtrlErr:
   MsgBox Err.Description, vbCritical, App.FileDescription
End Sub

Private Sub CboConvenio_Click(Area As Integer)
    On Error GoTo CtrlErr
    If CboConvenio.Text = "" Then Exit Sub
    Set CboPetitorio.RowSource = objPetitorio.Lista_Petitorio_Convenio(CboConvenio.BoundText)
    CboPetitorio.ListField = "DES_PETITORIO"
    CboPetitorio.BoundColumn = "COD_PETITORIO"
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.FileDescription
End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
    TxtProducto.Tipo = AlfaNumerico
    If KeyAscii = 13 Then
        If Len(Trim(TxtProducto.Text)) < 3 Then MsgBox "Ingresar como minimo 3 digitos", vbInformation + vbOKOnly, App.ProductName: TxtProducto.SetFocus: Exit Sub
             
             
             Call frm_ADM_Busqueda_Productos.LoadForm(Trim(TxtProducto.Text), _
                                                      objUsuario.CodigoLocal)
                                                      
        frm_ADM_Busqueda_Productos.Show
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
    Dim a As String
    
    On Error GoTo CtrlErr
    If txtCodigoGraba.Text = "" Then MsgBox "Ingrese el codigo de producto", vbCritical, Caption: Exit Sub
    If CboConvenio.BoundText = "" Then MsgBox "Seleccione el Convenio", vbCritical, Caption: Exit Sub
    'If CboPetitorio.BoundText = "" Then MsgBox "Seleccione el Petitorio", vbCritical, Caption: Exit Sub
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Autor : Arturo Escate
    'Fecha : 10/11/2009
    'Proposito: Esto es para validar si necesita autorizacion previa
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim ObjValidacion As New clsAprobacion
    Dim strNumeroSolicitud As String
    Dim strAccion As String
    Dim strMensaje As String
    Dim strCodigoAutorizacion As String
    Dim srtCodigoAUTH As String
    Dim strStore As String
    srtCodigoAUTH = ""
valida:
'TxtImpTot.text
    Dim strCadCodigoProducto As String
    Dim strSubTotal As String
    Dim strCodigoUsuario As String
    Dim strCantidad As String
    Dim strCantidadfrac As String
    Dim e As Integer
    
    Dim CadenaPrecioUnitProducto As String
    Dim CadenaBaseImponible As String
    Dim CadenaImpuesto As String
    Dim CadenaExonerado As String
Dim strSubTotalUsu  As String
    
    e = 0
        strCadCodigoProducto = ""
        strCantidad = ""
        strCantidadfrac = ""
        CadenaPrecioUnitProducto = ""
        CadenaBaseImponible = ""
        CadenaImpuesto = ""
        CadenaExonerado = ""
        strSubTotal = ""
        strSubTotalUsu = ""
        strCodigoUsuario = ""
        
    
        strCadCodigoProducto = Trim(txtCodigoGraba.Text) & "|"
        strCantidad = "0|"
        strCantidadfrac = "0|"
        CadenaPrecioUnitProducto = "0|"
        CadenaBaseImponible = CadenaBaseImponible & "0|"
        CadenaImpuesto = "0|"
        CadenaExonerado = "0|"
        strSubTotal = "0|"
    
    strStore = ObjValidacion.Solicita("3", strAccion, strMensaje, srtCodigoAUTH, objUsuario.CodigoLocal, objUsuario.CodigoLiquidacion, _
                                      "", "", "", _
                                      "", "", "1", "", "", objUsuario.Codigo, "", _
                                      strCodigoAutorizacion, "", "", "", "", "", "", "", "", "", "", strCadCodigoProducto, strCantidad, _
                                      strCantidadfrac, CadenaPrecioUnitProducto, CadenaBaseImponible, CadenaImpuesto, CadenaExonerado, _
                                      strSubTotal, strCodigoUsuario, strSubTotalUsu, CboConvenio.BoundText, "", Trim(txtCodigoGraba.Text), "", "", "", "", "")
    If Not strStore = "" Then
        MsgBox strStore, vbCritical, App.ProductName
        Exit Sub
    Else
        Select Case strAccion
            Case 0
                    MsgBox strMensaje, vbInformation, App.ProductName
            Case 1
                   MsgBox strMensaje, vbCritical, App.ProductName
                   Exit Sub
            Case 2
                   MsgBox strMensaje, vbInformation, App.ProductName
                   Exit Sub
            Case 3
                If MsgBox(strMensaje & Chr(13) & "¿Desea ingresar el codigo de autorización?", vbYesNo + vbInformation, App.ProductName) = vbYes Then
                    srtCodigoAUTH = frmAprobacion.Carga
                    If Not srtCodigoAUTH = "" Then
                        GoTo valida
                        Exit Sub
                    End If
                   Exit Sub
                Else
                    Exit Sub
                End If
            Case Else
                   MsgBox "no esta implementado", vbInformation, App.ProductName
                   Exit Sub
        End Select
    End If
    ''------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
    
    
    
    
    a = objPetitorio.HabilitaProducto(CboConvenio.BoundText, _
                                        Trim(txtCodigoGraba.Text), _
                                        CodigoPetitorio, _
                                        objUsuario.Codigo)
    If a = "" Then
          MsgBox "Se agrego con exito el producto al petitorio", vbInformation, "Grabar producto en petitorio"
          TxtProducto.Text = ""
          CboConvenio.Text = "": CboPetitorio.Text = ""
          Frame2.Caption = ""
          txtCodigoGraba.Text = ""
          TxtProducto.SetFocus
       Else
          MsgBox a, vbCritical, App.FileDescription
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.FileDescription
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    Set objPetitorio = Nothing
End Sub
