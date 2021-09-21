VERSION 5.00
Begin VB.Form frm_VTA_Cliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frm_VTA_Cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   8100
      Picture         =   "frm_VTA_Cliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6660
      Width           =   1095
   End
   Begin vbp_Ventas.ctlCliente ctlCliente1 
      Height          =   6555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11562
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   6780
      Picture         =   "frm_VTA_Cliente.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6660
      Width           =   1095
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
      Left            =   8460
      TabIndex        =   4
      Top             =   7320
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
      Left            =   6720
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodigo As String
Public Telefono As String

Private Sub cmdAceptar_Click()
On Error GoTo Control
    frm_VTA_Previa.bolCancelCliente = False
    frm_VTA_PreviaTomaPedido.bolCancelCliente = False
    If ctlCliente1.Grabar = "" Then
       If Not objVenta.CodigoConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then ''CAMBIAR POR CONSTANTE EN RIMAC
            objVenta.CodigoCliente = ctlCliente1.Codigo
          Else
            objVenta.CodigoCliente = ""
       End If
        objVenta.DesAuxCliTlf = ctlCliente1.Telefono
        mdiPrincipal.ctlCliente1.ConsultaCliente ctlCliente1.Codigo
        'Debug.Print mdiPrincipal.ctlCliente1.LocalDespacho
        Unload Me
    End If

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.number
End Sub

Private Sub cmdCancelar_Click()
    frm_VTA_Previa.bolCancelCliente = True
    If MsgBox("Desea cancelar el ingreso del cliente", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub ctlCliente1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmdCancelar_Click
End Sub

Private Sub Form_Activate()
'se quito codigo (ahora en load)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click
End Sub

Public Sub CargarValores()


On Error GoTo CtrlErr

    
    Select Case ctlCliente1.XTipoFuncion
        Case "Nuevo"
            ctlCliente1.Cargar Telefono
            objVenta.DesAuxCliTlf = Telefono
            If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = "0" Then
                ctlCliente1.DeshabilitaLocalPrecio
            End If
            If Not objUsuario.EsDelivery Then
                ctlCliente1.LocalAsignado = objUsuario.CodigoLocal
                ctlCliente1.LocalDespacho = objUsuario.CodigoLocal
            End If


            'ctlCliente1.SetFocus
        Case "Editar"
            ctlCliente1.Cargar Telefono
            Dim Cia As String
            If objUsuario.CodLocalCallCenter = "1DLV" Then
                Cia = "94"
            Else
                Cia = objUsuario.CodigoEmpresa
            End If
            If Not strCodigo = "" Then ctlCliente1.ConsultaCliente strCodigo, Cia
            objVenta.DesAuxCliTlf = ctlCliente1.Telefono
            If Not objUsuario.EsDelivery Then
                ctlCliente1.LocalAsignado = objUsuario.CodigoLocal
                ctlCliente1.LocalDespacho = objUsuario.CodigoLocal
            End If
            'para solo deshabilitar el local cuando es delivery lima
            If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = "0" Then
                ctlCliente1.DeshabilitaLocalPrecio
            End If
            'ctlCliente1.SetFocus

    End Select

    
    
    
    
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName


    'ctlCliente1.MODO = True
''''    On Error GoTo CtrlErr
''''    ctlCliente1.Cargar Telefono
''''    If Not strCodigo = "" Then ctlCliente1.ConsultaCliente strCodigo
''''    objVenta.DesAuxCliTlf = Telefono
''''    If Not objUsuario.EsDelivery Then
''''        ctlCliente1.LocalAsignado = objUsuario.CodigoLocal
''''      Else
''''        ctlCliente1.ConsultaCliente (objVenta.CodigoCliente)
''''    End If
''''   Exit Sub
''''CtrlErr:
''''     MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_Load()
    Call CargarValores
End Sub
