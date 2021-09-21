VERSION 5.00
Begin VB.Form frm_ADM_DetGuias 
   Caption         =   "Detalle de Guías"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Guía de Remisión"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9735
      Begin vbp_Ventas.ctlTextBox txtbuscar 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   1200
         Width           =   7215
         _ExtentX        =   12726
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
      Begin VB.Label Label6 
         Caption         =   "Buscar :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblcantproductos 
         Caption         =   "5"
         Height          =   255
         Left            =   8160
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblestado 
         Caption         =   "Activo"
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblnumitems 
         Caption         =   "15"
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblfecha 
         Caption         =   "01/01/2011"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblnumguia 
         Caption         =   "00000001"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Cant. Productos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Estado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Items :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Guía :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   9735
      Begin vbp_Ventas.ctlGrilla ctlgrdguias 
         Height          =   3375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5953
         Resalte         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_ADM_DetGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim objEntrega As New clsEntrega
Public numGuia As String

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.lblnumguia.Caption = numGuia
CargaCabGuia (numGuia)
cargaDetalle numGuia, SetBuscar
End Sub

Private Sub CargaCabGuia(num As String)
On Error GoTo Control

Dim ODyn As oraDynaset
Set ODyn = objEntrega.ListaCabGuia(num, objUsuario.CodigoLocal)
ODyn.MoveFirst
Me.lblestado.Caption = ODyn("COD_ESTADO").Value
Me.lblnumguia.Caption = ODyn("NUM_GUIA").Value
Me.lblcantproductos.Caption = ODyn("CTDPRODUCTOS").Value
Me.lblnumitems.Caption = ODyn("NUMITEMS").Value
Me.lblfecha.Caption = "" & ODyn("FCH_EMISION").Value

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Sub cargaDetalle(num As String, busca As String)

On Error GoTo Control

Set Me.ctlgrdguias.DataSource = objEntrega.ListaDetGuia(num, busca, objUsuario.CodigoLocal)
SeteaGrilla

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant

    arrCampos = Array("NUM_ITEM", "NUM_GUIA", "COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "UNIDAD", "CANTIDAD")
    arrCaption = Array("Nº", "Nº Guía", "Codigo", "Descripcion", "Laboratorio", "UND", "Cantidad")
    arrAncho = Array(600, 1200, 1200, 2200, 2200, 800, 800)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgRight)

    Me.ctlgrdguias.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    Me.ctlgrdguias.Columns(0).Merge = False
    Me.ctlgrdguias.Columns(1).Merge = False
    Me.ctlgrdguias.Columns(2).Merge = False
    Me.ctlgrdguias.Columns(3).Merge = False
    Me.ctlgrdguias.Columns(4).Merge = False
    Me.ctlgrdguias.Columns(5).Merge = False
    Me.ctlgrdguias.Columns(6).Merge = False

    ctlgrdguias.Columns(1).BackColor = vbInfoBackground
    ctlgrdguias.Columns(2).BackColor = vbInfoBackground

End Sub

Function SetBuscar() As String
If Me.txtbuscar.Text = "" Then
SetBuscar = "@"
Else
SetBuscar = Me.txtbuscar.Text
End If
End Function

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cargaDetalle numGuia, "@"
   Me.ctlgrdguias.DataSource.FindFirst "COD_PRODUCTO='" & Trim(Me.txtbuscar.Text) & "'"
End If
End Sub
