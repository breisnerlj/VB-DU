VERSION 5.00
Begin VB.Form frm_VTA_Cliente_Prov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   6780
      Picture         =   "frm_VTA_Cliente_Prov.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   8100
      Picture         =   "frm_VTA_Cliente_Prov.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin vbp_Ventas.ctlClienteProv ctlCliente1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13361
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
      TabIndex        =   4
      Top             =   7260
      Width           =   1215
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
      TabIndex        =   3
      Top             =   7260
      Width           =   390
   End
End
Attribute VB_Name = "frm_VTA_Cliente_Prov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodigo As String
Public Telefono As String
Dim t0 As Variant
Dim t1 As Variant

Public Sub CargarValores()
    On Error GoTo CtrlErr

    Select Case ctlCliente1.XTipoFuncion
        Case "Nuevo"
            ctlCliente1.Cargar Telefono, Mid(objUsuario.NombrePC, 4, 3)
            ctlCliente1.LocalAsignado = Mid(objUsuario.NombrePC, 4, 3)
            ctlCliente1.LocalDespacho = Mid(objUsuario.NombrePC, 4, 3)
            ctlCliente1.DeshabilitaLocalPrecio
            
            objVenta.DesAuxCliTlf = Telefono
        Case "Editar"
            ctlCliente1.Cargar Telefono
            If Not strCodigo = "" Then ctlCliente1.ConsultaCliente strCodigo
            objVenta.DesAuxCliTlf = ctlCliente1.Telefono
            If Not objUsuario.EsDelivery Then
                ctlCliente1.LocalAsignado = objUsuario.CodigoLocal
                ctlCliente1.LocalDespacho = objUsuario.CodigoLocal
            End If
            'para solo deshabilitar el local cuando es delivery lima
            If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = "0" Then
                ctlCliente1.DeshabilitaLocalPrecio
            End If
    End Select
   
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo Control
    
    frm_VTA_Previa.bolCancelCliente = False
    
    If ctlCliente1.Grabar = "" Then
       If Not objVenta.CodigoConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then
            objVenta.CodigoCliente = ctlCliente1.Codigo
          Else
            objVenta.CodigoCliente = ""
       End If
        objVenta.DesAuxCliTlf = ctlCliente1.Telefono
        Unload Me
    End If

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
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
    't1 = Format(Now, "hh:mm:ss")
    'MsgBox Format(TimeValue(t1) - TimeValue(t0), "hh:mm:ss")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click
End Sub

Private Sub Form_Load()
    't0 = Format(Now, "hh:mm:ss")
End Sub
