VERSION 5.00
Begin VB.Form frm_DLV_BuscaDistrito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Determinar Local por Distrito"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3855
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   5655
      Begin VB.CheckBox chkCallao 
         Caption         =   "Callao"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   3480
         Width           =   1095
      End
      Begin vbp_Ventas.ctlDataCombo cboDespacho 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   3000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboLocal 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   2472
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboDistrito 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   1944
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboProvincia 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1416
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboDepartamento 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   888
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboPais 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   3615
         _ExtentX        =   2778
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Local Despacho :"
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
         Left            =   180
         TabIndex        =   14
         Top             =   3060
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Local Precio :"
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
         Left            =   495
         TabIndex        =   13
         Top             =   2532
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Distrito :"
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
         Left            =   960
         TabIndex        =   12
         Top             =   2004
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Provincia :"
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
         Left            =   765
         TabIndex        =   11
         Top             =   1476
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Departamento :"
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
         Left            =   375
         TabIndex        =   10
         Top             =   948
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pais :"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   3285
      Picture         =   "frm_DLV_BuscaDistrito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   1485
      Picture         =   "frm_DLV_BuscaDistrito.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4140
      Width           =   1215
   End
End
Attribute VB_Name = "frm_DLV_BuscaDistrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bolCancelar As Boolean
Dim strDepartamento As String
Dim strProvincia As String
Dim strDistrito As String
Dim strUbigeo As String


Private Sub cboDepartamento_Change()
    Set cboProvincia.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, cboDepartamento.BoundText)
    cboProvincia.ListField = "Descripcion"
    cboProvincia.BoundColumn = "Codigo"
    cboProvincia.BoundText = strProvincia
End Sub


Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
sub_Tecla KeyCode
End Sub

Private Sub cboDespacho_KeyDown(KeyCode As Integer, Shift As Integer)
sub_Tecla KeyCode
End Sub

Private Sub cboDistrito_Change()
Dim rs As oraDynaset
Dim objLocal As New clsLocal

    strUbigeo = cboDepartamento.BoundText & cboProvincia.BoundText & cboDistrito.BoundText
'''    Set cboUrbanizacion.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_URBANIZACION", 0, "2", strUbigeo)
'''    cboUrbanizacion.ListField = "DES_URBANIZACION"
'''    cboUrbanizacion.BoundColumn = "COD_URBANIZACION"

        Set rs = objLocal.ListaLocalPredetDLV(objUsuario.CodigoEmpresa, strUbigeo)
        If Not rs.EOF Then
            cboLocal.BoundText = "" & rs("COD_LOCAL_PRECIO").Value
            cboDespacho.BoundText = "" & rs("COD_LOCAL_REF").Value
        End If
        
        


End Sub


Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
sub_Tecla KeyCode
End Sub

Private Sub cboLocal_KeyDown(KeyCode As Integer, Shift As Integer)
sub_Tecla KeyCode
End Sub

Private Sub cboPais_KeyDown(KeyCode As Integer, Shift As Integer)
sub_Tecla KeyCode
End Sub

Private Sub cboProvincia_Change()
    Set cboDistrito.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, cboDepartamento.BoundText, cboProvincia.BoundText)
    cboDistrito.ListField = "Descripcion"
    cboDistrito.BoundColumn = "Codigo"
    cboDistrito.BoundText = strDistrito
End Sub



Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
sub_Tecla KeyCode
End Sub

Private Sub chkCallao_Click()
    If chkCallao.Value = "1" Then
        cboDepartamento.BoundText = "07"
        cboProvincia.BoundText = "01"
    End If
End Sub

Private Sub cmdAceptar_Click()
    With mdiPrincipal
        .txtDireccion.Text = ""
        With .ctlCliente1
            .ConsultaPrecio = True
            .Limpiar
            .Ubigeo = strUbigeo
            .LocalAsignado = cboLocal.BoundText
            .sCia = "" & gclsOracle.FN_Valor("btlprod.pkg_local.fn_devuelve_cia", cboDespacho.BoundText)
            .LocalDespacho = cboDespacho.BoundText
            .Cargar
        End With
    End With
    
    bolCancelar = False
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    bolCancelar = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strUbigeo As String
    Dim objCliente  As New clsCliente
    Dim objLocal As New clsLocal
    
    Set cboLocal.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
    cboLocal.ListField = "local_dex"
    cboLocal.BoundColumn = "COD_LOCAL"
    Set cboDespacho.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
    cboDespacho.ListField = "local_dex"
    cboDespacho.BoundColumn = "COD_LOCAL"



    Set cboPais.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PAIS", 0)
    cboPais.ListField = "Descripcion"
    cboPais.BoundColumn = "Codigo"
    cboPais.BoundText = "00"


    Set cboDepartamento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0)
    cboDepartamento.ListField = "Descripcion"
    cboDepartamento.BoundColumn = "Codigo"


    strUbigeo = "" & objUsuario.UbigeoLocal
    If Not strUbigeo = "" Then
    On Error GoTo Y
       strDepartamento = Mid(strUbigeo, 1, 2)
       strProvincia = Mid(strUbigeo, 3, 2)
       strDistrito = Mid(strUbigeo, 5, 2)
           
Y:
    End If
    cboDepartamento.BoundText = strDepartamento






End Sub

Private Sub sub_Tecla(ByVal KeyCode As Integer)
If KeyCode = vbKeyEscape Then
    cmdCancelar_Click
End If

End Sub
