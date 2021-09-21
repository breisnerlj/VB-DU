VERSION 5.00
Begin VB.Form frm_VTA_ClienteDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Clientes"
   ClientHeight    =   6165
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7350
   Icon            =   "frm_VTA_ClienteDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Nuevo"
      Height          =   435
      Left            =   2040
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Editar"
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   6000
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   4680
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla GrdBusCliente 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9128
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   6360
      TabIndex        =   8
      Top             =   5760
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   4920
      TabIndex        =   7
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   3840
      TabIndex        =   6
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   2520
      TabIndex        =   5
      Top             =   5760
      Width           =   375
   End
End
Attribute VB_Name = "frm_VTA_ClienteDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim ObjCliente As New clsCliente
Public Pantalla As String
Public vstr1 As String
Public vstr2 As String
Public vstr3 As String
Public vstr4 As String

Public vstrNumDocId As String
Public vstrUbigeo As String

Private Sub Command1_Click()
    ''09/06/2008 - Comentado por JLopez xq la consulta de cliente se hace en Load del formulario
    ''frm_VTA_Cliente.ctlCliente1.Cargar
    ''frm_VTA_Cliente.ctlCliente1.ConsultaCliente GrdBusCliente.Columns(0).Value
    frm_VTA_Cliente.ctlCliente1.XTipoFuncion = "Editar"
    frm_VTA_Cliente.strCodigo = GrdBusCliente.Columns(0).Value
    frm_VTA_Cliente.CargarValores
    frm_VTA_Cliente.Show vbModal
End Sub

Private Sub Command2_Click()
    ''09/06/2008 - Comentado por JLopez xq la carga de la pantalla se hace en Load del formulario
    ''frm_VTA_Cliente.ctlCliente1.Cargar
    frm_VTA_Cliente.ctlCliente1.XTipoFuncion = "Nuevo"
    frm_VTA_Cliente.ctlCliente1.Codigo = ""
    frm_VTA_Cliente.CargarValores
    frm_VTA_Cliente.Show vbModal
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            Command2_Click
        Case vbKeyF2
             Command1_Click
        Case vbKeyF3
            cmdAceptar_Click
    End Select
End Sub

Private Sub Form_Load()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("COD_CLIENTE", "NUM_DOCUMENTO_ID", "DESCRIPCION", "DIRECCION", "UBIGEO")
    arrCaption = Array("Codigo", "Nº Ident", "Datos Comerciales", "Dirección", "Ubigeo")
    arrAncho = Array(1000, 1000, 2200, 2800, 1200)
    arrAlineacion = Array(vbAlignNone, vbAlignNone, vbAlignNone, vbAlignLeft, vbAlignNone, vbAlignNone)
    GrdBusCliente.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub cmdAceptar_Click()
     Call GrdBusCliente_DblClick
End Sub

Private Sub GrdBusCliente_DblClick()
'objVenta.CodigoTipoVenta

objVenta.CodigoCliente = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
Select Case Pantalla
End Select
''''Esta valida por
Select Case Pantalla
    Case 0 'desde la pantalla de documento
''''''        If frm_VTA_Documento.ctlCboTipCliente.BoundText = "1" Then
''''''                frm_VTA_Documento.TxtRuc.Text = "" & GrdBusCliente.Columns("NUM_DOCUMENTO_ID").Value
''''''                frm_VTA_Documento.TxtRazonSocial.Text = "" & GrdBusCliente.Columns("DESCRIPCION").Value
''''''                frm_VTA_Documento.txtDireccionPJ.Text = "" & GrdBusCliente.Columns("DIRECCION").Value
''''''                frm_VTA_Documento.txtCodigoCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
''''''                objVenta.UbigeoEntrega = "" & GrdBusCliente.Columns("UBIGEO").Value
''''''            Else
''''''                frm_VTA_Documento.txtDNI.Text = "" & GrdBusCliente.Columns("NUM_DOCUMENTO_ID").Value
''''''                frm_VTA_Documento.txtNombres.Text = "" & GrdBusCliente.DataSource("DES_NOM_CLIENTE").Value
''''''                frm_VTA_Documento.txtApellidos.Text = "" & GrdBusCliente.DataSource("DES_APE_CLIENTE").Value
''''''                frm_VTA_Documento.txtDireccionPN.Text = "" & GrdBusCliente.Columns("DIRECCION").Value
''''''                frm_VTA_Documento.txtCodigoCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
''''''                objVenta.UbigeoEntrega = "" & GrdBusCliente.DataSource("UBIGEO").Value
''''''         End If

    Case 1 'desde la pantalla de cobranza al credito
            frm_VTA_Cobranza.txtCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
            frm_VTA_Cobranza.LblCliente.Caption = "" & GrdBusCliente.Columns("DESCRIPCION").Value
            frm_VTA_Cobranza.txtNumDocumento.Text = "" & GrdBusCliente.Columns("NUM_DOCUMENTO_ID").Value
            frm_VTA_Cobranza.CargaDatos
            frm_VTA_Cobranza.txtCodigoCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
    Case 2 'desde la pantalla de recetario magistral
         frm_VTA_RecetarioM.txtCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
         frm_VTA_RecetarioM.LblCliente.Caption = "" & GrdBusCliente.Columns("DESCRIPCION").Value
         vstrNumDocId = "" & GrdBusCliente.Columns("NUM_DOCUMENTO_ID").Value
         vstrUbigeo = "" & GrdBusCliente.Columns("UBIGEO").Value
            
End Select
''''''''''    If objVenta.CodigoTipoVenta = Recetario Then
''''''''''         frm_VTA_RecetarioM.TxtCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
''''''''''         frm_VTA_RecetarioM.LblCliente.Caption = "" & GrdBusCliente.Columns("DESCRIPCION").Value
''''''''''    ElseIf penumVentCli = Documento Then
''''''''''         '- Persona Juridico -'
''''''''''         If frm_VTA_Documento.ctlCboTipCliente.BoundText = "1" Then
''''''''''                frm_VTA_Documento.txtRUC.Text = "" & GrdBusCliente.Columns("NUM_DOCUMENTO_ID").Value
''''''''''                frm_VTA_Documento.txtRazonSocial.Text = "" & GrdBusCliente.Columns("DESCRIPCION").Value
''''''''''                frm_VTA_Documento.txtDireccionPJ.Text = "" & GrdBusCliente.Columns("DIRECCION").Value
''''''''''                frm_VTA_Documento.txtCodigoCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
''''''''''
''''''''''            Else
''''''''''                frm_VTA_Documento.txtDNI.Text = "" & GrdBusCliente.Columns("NUM_DOCUMENTO_ID").Value
''''''''''                frm_VTA_Documento.txtNombres.Text = "" & GrdBusCliente.Columns("DESCRIPCION").Value
''''''''''                frm_VTA_Documento.txtApellidos.Text = "" & GrdBusCliente.Columns("DESCRIPCION").Value
''''''''''                frm_VTA_Documento.txtDireccionPN.Text = "" & GrdBusCliente.Columns("DIRECCION").Value
''''''''''                frm_VTA_Documento.txtCodigoCliente.Text = "" & GrdBusCliente.Columns("COD_CLIENTE").Value
''''''''''         End If
''''''''''    Else
''''''''''        frm_VTA_Cobranza.TxtCliente.Text = GrdBusCliente.Columns("COD_CLIENTE").Value
''''''''''        frm_VTA_Cobranza.LblCliente.Caption = GrdBusCliente.Columns("DESCRIPCION").Value
''''''''''    End If
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub GrdBusCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            GrdBusCliente_DblClick
    End Select
End Sub

