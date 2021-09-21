VERSION 5.00
Begin VB.Form frm_ADM_TransportistaBuscar 
   Caption         =   "Búsqueda"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin vbp_Ventas.ctlTextBox txtBuscar 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
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
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_ADM_TransportistaBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objEntrega As New clsEntrega

Private Sub ctlGrilla1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    frm_ADM_Transportista.txtIdTransportista.Text = Me.ctlGrilla1.Columns("ID_TRANSPORTISTA").Value
    frm_ADM_Transportista.txtDesTransportista.Text = Me.ctlGrilla1.Columns("DES_TRANSPORTISTA").Value
    frm_ADM_Transportista.recibe
End If
End Sub

Private Sub Form_Load()
SeteaGrilla
Set Me.ctlGrilla1.DataSource = objEntrega.ListaTransportista("", "1", "", Trim(Me.txtBuscar.Text))
End Sub

Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant

    arrCampos = Array("ID_TRANSPORTISTA", "DES_TRANSPORTISTA")
    arrCaption = Array("Id.", "Descripción")
    arrAncho = Array(800, 2500)
    arrAlineacion = Array(dbgCenter, dbgLeft)
    Me.ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SeteaGrilla
        Set Me.ctlGrilla1.DataSource = objEntrega.ListaTransportista("", "1", "", Trim(Me.txtBuscar.Text))
    End If
End Sub
