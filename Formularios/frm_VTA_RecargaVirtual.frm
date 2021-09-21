VERSION 5.00
Begin VB.Form frm_VTA_RecargaVirtual 
   BorderStyle     =   0  'None
   Caption         =   "Recarga Virtual"
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlTextBox txtPrecioUnitario 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_RecargaVirtual.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_RecargaVirtual.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrilla grdRecarga 
      Height          =   4995
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8811
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtImporteTotal 
      Height          =   375
      Left            =   5100
      TabIndex        =   3
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Tipo            =   3
      MaxLength       =   6
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
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
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
   Begin vbp_Ventas.ctlTextBox txtNumeroTelefonico 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   5220
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Tipo            =   5
      MaxLength       =   15
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
   Begin vbp_Ventas.ctlTextBox txtConfirmacion 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   6240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Tipo            =   5
      MaxLength       =   15
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmación Teléfono:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto de recarga:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   5280
      Width           =   1935
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
      TabIndex        =   6
      Top             =   6900
      Width           =   390
   End
End
Attribute VB_Name = "frm_VTA_RecargaVirtual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Public flgRecarga As Boolean
Public strCodigoPadre As String
Public strDescripcionPadre As String


Private Sub cmdAceptar_Click()
On Error GoTo handle
    
    If txtNumeroTelefonico.Text <> txtConfirmacion.Text Then
        MsgBox "El número de teléfono no es igual al número de confirmación", vbCritical, App.ProductName
        txtNumeroTelefonico.SetFocus
        Exit Sub
    End If
    If Len(txtNumeroTelefonico.Text) > txtNumeroTelefonico.MaxLength Then
        MsgBox "El número de telefono no puede ser mayor a " & txtNumeroTelefonico.MaxLength, vbCritical, App.ProductName
        txtNumeroTelefonico.SetFocus
        Exit Sub
    End If
    If Val(txtImporteTotal.Text) <= 0 And txtImporteTotal.Visible = True Then
        MsgBox "Ingresar el importe", vbCritical, App.ProductName
        txtImporteTotal.SetFocus
        Exit Sub
    End If
    Dim intCantidad As Integer
    Dim strImporte As String
    If grdRecarga.DataSource("FLG_CONTROL_CANT") = "1" Then
        intCantidad = Val(txtImporteTotal.Text) / Val(grdRecarga.DataSource("IMP_VALOR_CLIE"))
        strImporte = 0 'Val(grdRecarga.DataSource("IMP_VALOR_CLIE")) * intCantidad
    Else
        intCantidad = 1
        strImporte = Val(grdRecarga.DataSource("IMP_VALOR_CLIE"))
    End If
    
    objVenta.AgregaServicio strCodigoPadre, _
                        strDescripcionPadre, _
                        grdRecarga.Columns(0).Value, _
                        grdRecarga.Columns(1).Value, _
                        "", _
                        Servicio, _
                        strImporte, _
                        "", _
                        strImporte, _
                        grdRecarga.Columns(2).Value, "", txtNumeroTelefonico.Text, "0", intCantidad

    frm_VTA_Servicios.grdServicios.Rebind
    Unload Me

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    setteaFormulario Me
    SetteaGrilla
    'Set grdRecarga.DataSource = objServicio.ListaRecarga(flgRecarga)
    Set grdRecarga.DataSource = objServicio.Lista("", strCodigoPadre)
    txtImporteTotal.Visible = flgRecarga
    txtNumeroTelefonico.Visible = flgRecarga
    txtConfirmacion.Visible = flgRecarga
    Label3.Visible = flgRecarga
    Label2.Visible = flgRecarga
    Label4.Visible = flgRecarga
End Sub
Sub SetteaGrilla()
    Dim arrCampos, arrCaption, arrAncho, arrAlineacion
    arrCampos = Array("COD_SERVICIO", "DES_SERVICIO", "COD_PRODUCTO")
    arrCaption = Array("Codigo", "Servicio", "CodProducto")
    arrAncho = Array(1000, 4500, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgGeneral)
    grdRecarga.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub grdRecarga_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub grdRecarga_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If Not "" & grdRecarga.DataSource("FLG_SERVICIO_ALQUILER") = "" Then
        txtNumeroTelefonico.MaxLength = Val("" & grdRecarga.DataSource("FLG_SERVICIO_ALQUILER"))
        txtConfirmacion.MaxLength = Val("" & grdRecarga.DataSource("FLG_SERVICIO_ALQUILER"))
    Else
        txtNumeroTelefonico.MaxLength = 15
        txtConfirmacion.MaxLength = 15
    End If
End Sub
