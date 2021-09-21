VERSION 5.00
Begin VB.Form frmMotorizado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Motorizados"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   Icon            =   "frmMotorizado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
   Begin vbp_Ventas.ctlGrilla grdMotorizado 
      Height          =   4890
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8625
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   30
      TabIndex        =   2
      Top             =   555
      Width           =   8565
      Begin vbp_Ventas.ctlDataCombo ctlCboEstado 
         Height          =   315
         Left            =   5760
         TabIndex        =   7
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.OptionButton OptOpciones 
         Appearance      =   0  'Flat
         Caption         =   "Apellidos"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   232
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptOpciones 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   232
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7800
         Picture         =   "frmMotorizado.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   165
         Width           =   600
      End
      Begin vbp_Ventas.ctlTextBox txtDato 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMotorizado As New clsMotorizado
Dim varBookMark As Variant

Private Sub cmdBuscar_Click()
    If OptOpciones(0).Value = True Then
        Set grdMotorizado.DataSource = objMotorizado.Lista(txtDato.Text, ctlCboEstado.BoundText)
    Else
        Set grdMotorizado.DataSource = objMotorizado.ListaApellidos(txtDato.Text, ctlCboEstado.BoundText)
    End If
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    Call ListaGrd
    
    Set ctlCboEstado.RowSource = objMotorizado.LstEstMotorizado_Act_Ina
    ctlCboEstado.BoundColumn = "COD"
    ctlCboEstado.ListField = "DES"
    ctlCboEstado.BoundText = "1"
    
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
            Call ListaGrd
        Case 5
            grdMotorizado.MostrarImprimir
        Case 6
            grdMotorizado.MostrarExcel
        Case 7
            grdMotorizado.MostrarEmail
        Case 8
            Unload Me
    End Select
End Sub

Private Sub Buscar()
'    If XEntro = False Then
'        XEntro = True
'        Me.Height = Me.Height + 500
'        grdMotorizado.top = grdMotorizado.top + 500
'        txtDato.Text = ""
'    End If
    txtDato.SetFocus
End Sub

Private Sub Form_Activate()
    If grdMotorizado.ApproxCount = 0 Then Exit Sub
    grdMotorizado.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_MOTORIZADO", "COD_ESTADO_MOTORIZADO", "DES_ESTADO_MOTORIZADO", "COD_LOCAL", "COD_LOCAL2", "DES_NOMBRES", "DES_APELLIDOS", "DES_NUMERO", "DES_ALIAS", "FLG_ACTIVO")
                      
    arrCaption = Array("Código", "Cod. Estado", "Estado", "Local Asignado", "Local Asignado", "Nombres", "Apellidos", "Num Ref", "Alias", "Activo")
    
    arrAncho = Array(800, 0, 1200, 9, 1200, 2000, 3000, 800, 1200, 800)
    
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgRight, dbgLeft, dbgCenter)
    
    grdMotorizado.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdMotorizado.Columns(1).Visible = False
    grdMotorizado.Columns(1).AllowSizing = False
    grdMotorizado.Columns("COD_LOCAL").Visible = False
    grdMotorizado.Columns("COD_LOCAL").AllowSizing = False
End Sub

Private Sub ListaGrd()
'    If XEntro = True Then
'        Me.Height = Me.Height - 500
'        grdMotorizado.top = grdMotorizado.top - 500
'        XEntro = False
'    End If
    Call SeteaGrilla
    Set grdMotorizado.DataSource = objMotorizado.Lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If MsgBox("Salir del Módulo?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
'        Cancel = 0
'    Else
'        Cancel = 1
'        Form_Activate
'    End If
End Sub

Private Sub Nuevo()
'    If XEntro = True Then
'        Me.Height = Me.Height - 500
'        grdMotorizado.top = grdMotorizado.top - 500
'        XEntro = False
'    End If
    Call frmGrabaMotorizado.datos("", "", "", "", "", "", "", 0, "-(Nuevo)")
    Call ListaGrd
End Sub

Private Sub Editar()
'    If XEntro = True Then
'        Me.Height = Me.Height - 500
'        grdMotorizado.top = grdMotorizado.top - 500
'        XEntro = False
'    End If
    If grdMotorizado.ApproxCount = 0 Then Exit Sub
    varBookMark = grdMotorizado.Bookmark
    Call frmGrabaMotorizado.datos(grdMotorizado.Columns("COD_MOTORIZADO").Value, _
                                  grdMotorizado.Columns("COD_ESTADO_MOTORIZADO").Value, _
                                  grdMotorizado.Columns("COD_LOCAL").Value, _
                                  grdMotorizado.Columns("DES_NOMBRES").Value, _
                                  grdMotorizado.Columns("DES_APELLIDOS").Value, _
                                  grdMotorizado.Columns("DES_NUMERO").Value, _
                                  grdMotorizado.Columns("DES_ALIAS").Value, _
                                  grdMotorizado.Columns("FLG_ACTIVO").Value, _
                                  "-(Editar)")
                            
                            
    Call ListaGrd
    grdMotorizado.Bookmark = varBookMark
End Sub

Private Sub grdMotorizado_DblClick()
    Call Editar
End Sub

Private Sub grdMotorizado_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call Editar
    End Select
End Sub

Private Sub OptOpciones_Click(Index As Integer)
    txtDato.Text = "": txtDato.SetFocus
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdBuscar_Click
End Sub
