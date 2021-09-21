VERSION 5.00
Begin VB.Form frmZona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Zonas"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmZona.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
   Begin vbp_Ventas.ctlGrilla grdZona 
      Height          =   3120
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   5503
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6510
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   4950
         Picture         =   "frmZona.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   165
         Width           =   600
      End
      Begin vbp_Ventas.ctlTextBox txtDato 
         Height          =   315
         Left            =   705
         TabIndex        =   4
         Top             =   180
         Width           =   4200
         _ExtentX        =   7408
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmZona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objZona As New clsZona
Dim varBookMark As Variant
Dim XEntro As Boolean

Private Sub cmdBuscar_Click()
    Set grdZona.DataSource = objZona.Lista(txtDato.Text)
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case Index
        Case 1
            Call Nuevo
        Case 2
            Call Editar
        Case 3
            Call Buscar
        Case 4
            Call ListaGrid
        Case 5
            grdZona.MostrarImprimir
        Case 6
            grdZona.MostrarExcel
        Case 7
            grdZona.MostrarEmail
        Case 8
            Unload Me
    End Select
End Sub

Private Sub Buscar()
    If XEntro = False Then
        XEntro = True
        Me.Height = Me.Height + 500
        grdZona.top = grdZona.top + 500
    End If
    txtDato.SetFocus
End Sub

Private Sub Form_Activate()
    If grdZona.ApproxCount = 0 Then Exit Sub
    grdZona.SetFocus
End Sub

Private Sub Formato()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_ZONA", "DES_ZONA", "DES_ABREVIATURA", "FLG_ACTIVO")
    arrCaption = Array("Código", "Nombre Zona", "Abreviatura", "Activo")
    arrAncho = Array(900, 1500, 1200, 800)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgCenter)
    grdZona.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub ListaGrid()
    If XEntro = True Then
        Me.Height = Me.Height - 500
        grdZona.top = grdZona.top - 500
        XEntro = False
    End If
    Call Formato
    Set grdZona.DataSource = objZona.Lista
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    XEntro = False
    Call ListaGrid
End Sub

Private Sub Nuevo()
    If XEntro = True Then
        Me.Height = Me.Height - 500
        grdZona.top = grdZona.top - 500
        XEntro = False
    End If
    Call frmGrabaZona.SetearDatos("", "", "", 0, "Nuevo")
    frmGrabaZona.Show 1
    Call ListaGrid
End Sub

Private Sub Editar()
    If XEntro = True Then
        Me.Height = Me.Height - 500
        grdZona.top = grdZona.top - 500
        XEntro = False
    End If
    If grdZona.ApproxCount = 0 Then Exit Sub
    varBookMark = grdZona.Bookmark
    Call frmGrabaZona.SetearDatos(grdZona.Columns("COD_ZONA").Value, _
                                  grdZona.Columns("DES_ZONA").Value, _
                                  grdZona.Columns("DES_ABREVIATURA").Value, _
                                  grdZona.Columns("FLG_ACTIVO").Value, _
                                  "Editar")
    frmGrabaZona.Show 1
    Call ListaGrid
    grdZona.Bookmark = varBookMark
End Sub

Private Sub grdZona_DblClick()
    Call Editar
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdBuscar_Click
End Sub
