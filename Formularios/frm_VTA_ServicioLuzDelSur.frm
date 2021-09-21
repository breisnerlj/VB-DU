VERSION 5.00
Begin VB.Form frm_VTA_ServicioLuzDelSur 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Pago"
      Height          =   2295
      Left            =   360
      TabIndex        =   18
      Top             =   3840
      Width           =   6255
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Suministro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lblSuministro 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblBoleta 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Boleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   570
      End
   End
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.OptionButton optPago 
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   6
      Top             =   6600
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optPago 
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   6240
      Width           =   615
   End
   Begin VB.ComboBox cboCriterio 
      Height          =   315
      ItemData        =   "frm_VTA_ServicioLuzDelSur.frx":0000
      Left            =   360
      List            =   "frm_VTA_ServicioLuzDelSur.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1860
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   1860
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_ServicioLuzDelSur.frx":0041
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_ServicioLuzDelSur.frx":05CB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtMontonMinimo 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      ColorDefault    =   -2147483639
      ColorDefault    =   -2147483639
      Enabled         =   0   'False
      Bloqueado       =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox txtImporteTotal 
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      ColorDefault    =   -2147483639
      ColorDefault    =   -2147483639
      Enabled         =   0   'False
      Bloqueado       =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox txtCriterio 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   1860
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlGrilla grdDocumentos 
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "S/."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   1740
      TabIndex        =   17
      Top             =   6660
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "S/."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1740
      TabIndex        =   16
      Top             =   6240
      Width           =   240
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
      Index           =   6
      Left            =   4380
      TabIndex        =   15
      Top             =   6900
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
      Index           =   5
      Left            =   6097
      TabIndex        =   14
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Importe Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   6600
      Width           =   1230
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Monto Mínimo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   6240
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_ServicioLuzDelSur.frx":0B55
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios - Luz del Sur"
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
      TabIndex        =   10
      Top             =   60
      Width           =   2415
   End
End
Attribute VB_Name = "frm_VTA_ServicioLuzDelSur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Public strCodigoPadre As String
Public strDescripcionPadre As String

Private Sub cboCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo handle
'Servicio
Dim strTipo As String
Dim strMonto As String
    If optPago(0).Value = True Then
        strTipo = "0"
        strMonto = txtMontonMinimo.Text
    ElseIf optPago(1).Value = True Then
        strTipo = "1"
        strMonto = txtImporteTotal.Text
    Else
        MsgBox "Indicar si desea cancelar el monto mínimo o el importe total", vbCritical, App.ProductName
        Exit Sub
    End If
    
    If Val(strMonto) = 0 Then
        MsgBox "El importe seleccionado es cero, verifique", vbCritical, App.ProductName
        Exit Sub
    End If
    
    objVenta.AgregaServicio strCodigoPadre, strDescripcionPadre, grdDocumentos.DataSource("COD_SERVICIO").Value, grdDocumentos.DataSource("DES_SERVICIO").Value, lblSuministro, Servicio, _
    Val(strMonto), "", Val(strMonto), grdDocumentos.Columns("COD_PRODUCTO").Value, lblBoleta.Caption, lblSuministro, strTipo
    frm_VTA_Servicios.grdServicios.Rebind
    Unload Me

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo handle
    'ctlGrilla1.Array1 = objServicio.ConectaLuzDelSur(1, txtCriterio.Text)
    Set ctlGrilla1.DataSource = objServicio.ConectaLuzDelSurDB(1, txtCriterio.Text)
    ctlGrilla1.Rebind
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

    

Private Sub ctlGrilla1_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo handle
If ctlGrilla1.ApproxCount <= 0 Then Exit Sub
    txtImporteTotal.Text = Val("" & ctlGrilla1.Columns("MTO_FACTURA").Value)
    txtMontonMinimo.Text = Val("" & ctlGrilla1.Columns("MTO_MINIMO").Value)
    lblCliente.BackStyle = 1
    lblSuministro.BackStyle = 1
    lblBoleta.BackStyle = 1
    lblFecha.BackStyle = 1

    lblCliente.Caption = "" & ctlGrilla1.Columns("DSC_NOMBRE").Value
    lblSuministro.Caption = "" & ctlGrilla1.Columns("NRO_SUMINISTRO").Value
    lblBoleta.Caption = "" & ctlGrilla1.Columns("NRO_BOLETA").Value
    lblFecha.Caption = "" & ctlGrilla1.Columns("FCH_FAC").Value
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click
End Sub

Private Sub Form_Load()
On Error GoTo handle

    SetteaFormulario Me
    
        Dim arrCampos, arrCaption, arrAncho, arrAlineacion
    arrCampos = Array("COD_SERVICIO", "DES_SERVICIO", "COD_PRODUCTO")
    arrCaption = Array("Codigo", "Servicio", "CodProducto")
    arrAncho = Array(1000, 4500, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgGeneral)
    grdDocumentos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDocumentos.Columns(2).Visible = False
    
    Set grdDocumentos.DataSource = objServicio.Lista("", strCodigoPadre)


    'txtCriterio.Text = "0007434"
       'ctlGrilla1.TipoArray = True
       'Dim arrCampos, arrCaption, arrAncho, arrAlineacion As Variant
       arrCampos = Array("Suministro", "flgVerificacion", "N° Boleta", "flgVeriBol", "Importe", "F.Factura", "f.Vencimiento", "Imp. Minimo", "Cliente", "FlgDeuda")
       arrCaption = Array("Suministro", "flgVerificacion", "N° Boleta", "flgVeriBol", "Importe", "F.Factura", "f.Vencimiento", "Imp. Minimo", "Cliente", "FlgDeuda")
       arrAncho = Array("800", "800", "800", "800", "800", "800", "800", "800", "800", "800")
       arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
       'ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub grdDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

