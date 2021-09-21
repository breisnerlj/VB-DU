VERSION 5.00
Begin VB.Form frmLocalesPorZona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales Por Zona"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   Icon            =   "frmLocalesPorZona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
   Begin vbp_Ventas.ctlGrilla grdLocalesxZona 
      Height          =   3945
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   6959
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   510
      Width           =   6525
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   5745
         Picture         =   "frmLocalesPorZona.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   165
         Width           =   600
      End
      Begin VB.OptionButton OptOpciones 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   232
         Width           =   870
      End
      Begin VB.OptionButton OptOpciones 
         Appearance      =   0  'Flat
         Caption         =   "Local"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   1050
         TabIndex        =   3
         Top             =   232
         Value           =   -1  'True
         Width           =   975
      End
      Begin vbp_Ventas.ctlTextBox txtDato 
         Height          =   315
         Left            =   2070
         TabIndex        =   5
         Top             =   180
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   556
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
   End
End
Attribute VB_Name = "frmLocalesPorZona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objZona As New clsZona
'Dim XCodZona As String
'Dim XNomZona As String
'Dim XEntro As Boolean

Private Sub cmdBuscar_Click()
'    If OptOpciones(0).Value = True Then
'        Set grdLocalesxZona.DataSource = objZona.Lista_Locales(objUsuario.CodigoEmpresa, _
'                                                               XCodZona, _
'                                                               txtDato.Text)
'    Else
'        Set grdLocalesxZona.DataSource = objZona.Lista_Locales_Nombres(objUsuario.CodigoEmpresa, _
'                                                                       XCodZona, _
'                                                                       txtDato.Text)
'    End If
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
'    Select Case Index
'        Case 1
'            Call Asignar
'        Case 2
'            Call Buscar
'        Case 3
'            Call ListaGrid
'        Case 4
'            grdLocalesxZona.MostrarImprimir
'        Case 5
'            grdLocalesxZona.MostrarExcel
'        Case 6
'            grdLocalesxZona.MostrarEmail
'        Case 7
'            Unload Me
'    End Select
End Sub

Private Sub Buscar()
'    If XEntro = False Then
'        XEntro = True
'        Me.Height = Me.Height + 500
'        grdLocalesxZona.top = grdLocalesxZona.top + 500
'        txtDato.Text = ""
'    End If
'    txtDato.SetFocus
End Sub

Private Sub Form_Activate()
'    If grdLocalesxZona.ApproxCount = 0 Then Exit Sub
'    grdLocalesxZona.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    'XEntro = False
'    Call ListaGrid
'    ctlToolBar1.Buttons(2).Caption = "Asignar Locales"
'    ctlToolBar1.Buttons(1).Visible = False
End Sub

Private Sub ListaGrid()
'    If XEntro = True Then
'        Me.Height = Me.Height - 500
'        grdLocalesxZona.top = grdLocalesxZona.top - 500
'        XEntro = False
'    End If
'    Call Formato
'    Set grdLocalesxZona.DataSource = objZona.Lista_Locales(objUsuario.CodigoEmpresa, XCodZona)
End Sub

'Public Sub SetCodZona(CodZona As String, NomZona As String)
'    XCodZona = CodZona
'    XNomZona = NomZona
'End Sub

Private Sub Asignar()
'    Call frmGrabaLocalesPorZona.AsignaCodZona(XCodZona, XNomZona)
'    frmGrabaLocalesPorZona.lblzonas.Caption = XNomZona
'    frmGrabaLocalesPorZona.Caption = frmGrabaLocalesPorZona.Caption & "Zona : " & XNomZona
'    frmGrabaLocalesPorZona.Show 1
'    Call ListaGrid
End Sub

Private Sub OptOpciones_Click(Index As Integer)
    'txtDato.Text = "": txtDato.SetFocus
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
   ' If KeyAscii = 13 Then Call cmdBuscar_Click
End Sub

Private Sub Formato()
'  Dim arrCampos As Variant
'  Dim arrCaption As Variant
'  Dim arrAncho As Variant
'  Dim arrAlineacion As Variant
'
'    arrCampos = Array("COD_LOCAL", "DES_LOCAL", "FLG_ACTIVO")
'    arrCaption = Array("Cod. Local", "Descripción", "Activo")
'    arrAncho = Array(900, 3000, 800)
'    arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter)
'    grdLocalesxZona.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

