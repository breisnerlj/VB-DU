VERSION 5.00
Begin VB.Form frmGrabaLocalesPorZona 
   Caption         =   "Asignación de Locales Por Zona"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   Icon            =   "frmGrabaLocalesPorZona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5078
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6690
      Width           =   1155
   End
   Begin VB.CommandButton CmdAsignar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3038
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6690
      Width           =   1155
   End
   Begin VB.CommandButton CmdRegresaUno 
      Height          =   540
      Left            =   4365
      Picture         =   "frmGrabaLocalesPorZona.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3375
      Width           =   555
   End
   Begin VB.CommandButton CmdMandaUno 
      Height          =   555
      Left            =   4365
      Picture         =   "frmGrabaLocalesPorZona.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2505
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Caption         =   "Locales Asignados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   5130
      TabIndex        =   2
      Top             =   720
      Width           =   4050
      Begin vbp_Ventas.ctlListCombo lstOrigen 
         Height          =   5520
         Index           =   1
         Left            =   45
         TabIndex        =   3
         Top             =   195
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   9737
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Locales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   105
      TabIndex        =   0
      Top             =   720
      Width           =   4050
      Begin vbp_Ventas.ctlListCombo lstOrigen 
         Height          =   5520
         Index           =   0
         Left            =   45
         TabIndex        =   1
         Top             =   240
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   9737
      End
   End
   Begin vbp_Ventas.ctlDataCombo cboZona 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Zona:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   180
      Width           =   420
   End
   Begin VB.Label lblnum2 
      AutoSize        =   -1  'True
      Caption         =   "numfilas"
      Height          =   195
      Left            =   8100
      TabIndex        =   9
      Top             =   6615
      Width           =   570
   End
   Begin VB.Label lblnum1 
      AutoSize        =   -1  'True
      Caption         =   "numfilas"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   6615
      Width           =   570
   End
End
Attribute VB_Name = "frmGrabaLocalesPorZona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLocal As New clsLocal
Dim objZona As New clsZona
Dim XCodZona As String
Dim XNomZona As String
Dim XCad As String

Private Sub Form_Load()
    
    Set cboZona.RowSource = objZona.Lista
    cboZona.ListField = "DES_ZONA"
    cboZona.BoundColumn = "COD_ZONA"
    cboZona.BoundText = ""
    
End Sub

Sub CodigoZona()
    XCodZona = cboZona.BoundText
    XNomZona = cboZona.Text
    
    lstOrigen(1).Clear: lstOrigen(0).Clear
    lstOrigen(0).BoundColumn = "DES_LOCAL"
    lstOrigen(0).ListField = "COD_LOCAL_SAP"
    Set lstOrigen(0).RowSource = objLocal.LstLocalSNZona(objUsuario.CodigoEmpresa)

    lstOrigen(1).BoundColumn = "DES_LOCAL"
    lstOrigen(1).ListField = "COD_LOCAL_SAP"
    'Set lstOrigen(1).RowSource = objZona.Lista_Locales(objUsuario.CodigoEmpresa, XCodZona)
    Set lstOrigen(1).RowSource = objLocal.LstLocalCNZona(objUsuario.CodigoEmpresa, XCodZona)
    lblnum1.Caption = "N° Locales: " & lstOrigen(0).ListCount
    lblnum2.Caption = "N° Locales: " & lstOrigen(1).ListCount
    
    Call objZona.InicializaLista
End Sub

Private Sub cboZona_Click(Area As Integer)
On Error GoTo Control
If Area = 0 Then Exit Sub
    CodigoZona

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub CmdAsignar_Click()
On Error GoTo Asigna
Dim a As String
Dim xItem As ListItem
    'If lstOrigen(1).ListCount <= 0 Then MsgBox "Agrege los locales a asignar", vbExclamation, Caption: Exit Sub
    For Each xItem In lstOrigen(1).ListItems
       XCad = Mid(xItem.Text, 1, 3)
       objZona.AgregaLocal objUsuario.CodigoEmpresa, XCodZona, XCad
    Next xItem
    a = objZona.Graba_Local(gclsOracle.ODataBase, objUsuario.CodigoEmpresa, XCodZona, 1)
                    
    If a = "" Then
        MsgBox "Se asignaron los Locales a la Zona " & XNomZona, vbInformation + vbOKOnly, "Grabar"
        'Unload Me
    Else
        MsgBox a, vbCritical + vbOKOnly, "Atención"
    End If
    '****************************************************'
    '      Cambio del 11/10/2007 Por Cristhian Rueda     '
    ' Despues de grabar la asignación, limpia el arreglo '
    '****************************************************'
    Call objZona.InicializaLista
Exit Sub
Asigna:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub cmdMandaUno_Click()
Inicio:
    Dim xItem As ListItem
    For Each xItem In lstOrigen(0).ListItems
        If xItem.Selected Then
            If Not VerificaenLista(xItem.Text) Then
                lstOrigen(1).AddItem xItem.Text
                lstOrigen(0).RemoveItem xItem.Index
                GoTo Inicio
            Else
                Exit For
            End If
        End If
    Next xItem
    lblnum1.Caption = "N° Locales: " & lstOrigen(0).ListCount
    lblnum2.Caption = "N° Locales: " & lstOrigen(1).ListCount
End Sub

Private Sub cmdRegresaUno_Click()
Inicio:
    Dim xItem As ListItem
    For Each xItem In lstOrigen(1).ListItems
        If xItem.Selected Then
            If VerificaenLista(xItem.Text) Then
                lstOrigen(0).AddItem xItem.Text
                lstOrigen(1).RemoveItem xItem.Index
                GoTo Inicio
            Else
                lstOrigen(1).RemoveItem xItem.Index
                Exit For
            End If
        End If
    Next xItem
    lblnum1.Caption = "N° Locales: " & lstOrigen(0).ListCount
    lblnum2.Caption = "N° Locales: " & lstOrigen(1).ListCount
End Sub

Private Function VerificaenLista(Valor As String) As Boolean
    VerificaenLista = False
    Dim xItem As ListItem
    For Each xItem In lstOrigen(1).ListItems
        If xItem.Text = Valor Then
            VerificaenLista = True
            Exit For
        End If
     Next
End Function

Private Sub LstOrigen_DblClick(Index As Integer)
    Select Case Index
        Case 0
            cmdMandaUno_Click
        Case 1
            cmdRegresaUno_Click
    End Select
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
    Unload frmLocalesPorZona
End Sub
