VERSION 5.00
Begin VB.Form frm_ADM_GuiaAsociada 
   Caption         =   "Detalle de Guía"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Guía de Remisión"
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   9495
      Begin vbp_Ventas.ctlTextBox txtbuscar 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   840
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
         TabIndex        =   30
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblcantproductos 
         Caption         =   "5"
         Height          =   255
         Left            =   4800
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblestado 
         Caption         =   "Activo"
         Height          =   255
         Left            =   4680
         TabIndex        =   28
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblnumitems 
         Caption         =   "15"
         Height          =   255
         Left            =   8160
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblfecha 
         Caption         =   "01/01/2011"
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblnumguia 
         Caption         =   "00000001"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   600
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   600
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
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "[Esc] Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "[F2] Desasociar Entrega"
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "[F1] Asociar Entrega"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Listado de Guías"
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   9495
      Begin vbp_Ventas.ctlGrilla grdCabGuia 
         Height          =   1215
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2143
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Entregas Asociadas"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9495
      Begin VB.Label lblCantBultos 
         Caption         =   "5"
         Height          =   255
         Left            =   8160
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblCantPrecintos 
         Caption         =   "12"
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblCantEntregas 
         Caption         =   "15"
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblFechaEntrega 
         Caption         =   "01/01/2011"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblNumIngreso 
         Caption         =   "00000001"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Cant. Bultos:"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Cant. Precintos :"
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
         Left            =   3480
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Nº Entregas :"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Nº Ingreso :"
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
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   9495
      Begin vbp_Ventas.ctlGrilla grdDetGuia 
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3625
         Resalte         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_ADM_GuiaAsociada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strIdEntrega As String
Dim objEntrega As New clsEntrega

Public Sub cargaCabGuias()
On Error GoTo Handle
Dim ODyn As oraDynaset
Set ODyn = objEntrega.Lista(objUsuario.CodigoLocal, "", "", "", strIdEntrega)

If Not ODyn.EOF Then
ODyn.MoveFirst
    Me.lblCantBultos.Caption = "" & ODyn("ctd_bultos").Value
    Me.lblCantPrecintos.Caption = "" & ODyn("ctd_precintos").Value
    Me.lblFechaEntrega.Caption = "" & ODyn("fch_registra").Value
    Me.lblNumIngreso.Caption = "" & ODyn("id_entrega").Value
    Me.lblCantEntregas.Caption = "" & ODyn("num_guias").Value
    SeteaCabGuias
    Set Me.grdCabGuia.DataSource = objEntrega.ListaAsociados(objUsuario.CodigoLocal, strIdEntrega, "")
End If
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Public Sub cargaDetGuias(numGuia As String)
On Error GoTo Handle
Dim ODyn As oraDynaset
Set ODyn = objEntrega.ListaCabGuia(numGuia, objUsuario.CodigoLocal)
ODyn.MoveFirst
If Not ODyn.EOF Then
    Me.lblestado.Caption = "" & ODyn("COD_ESTADO").Value
    Me.lblnumguia.Caption = "" & ODyn("NUM_GUIA").Value
    Me.lblcantproductos.Caption = "" & ODyn("CTDPRODUCTOS").Value
    Me.lblnumitems.Caption = "" & ODyn("NUMITEMS").Value
    Me.lblfecha.Caption = "" & ODyn("FCH_EMISION").Value
    SeteaDetGuias
    Set Me.grdDetGuia.DataSource = objEntrega.ListaDetGuia(numGuia, "@", objUsuario.CodigoLocal)
End If
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Sub SeteaCabGuias()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
    arrCampos = Array("NUM_GUIA", "NUM_ENTREGA", "FCH_EMISION")
    arrCaption = Array("Nº Guía", "Nº Entrega", "Fec. Emision")
    arrAncho = Array(1200, 1300, 2000)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter)
    arrFoco = Array(False, False, False)
    Me.grdCabGuia.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    Me.grdCabGuia.Columns(0).Merge = False
    Me.grdCabGuia.Columns(1).Merge = False
    Me.grdCabGuia.Columns(2).Merge = False
End Sub

Sub SeteaDetGuias()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
    arrCampos = Array("NUM_ITEM", "NUM_GUIA", "COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "UNIDAD", "CANTIDAD")
    arrCaption = Array("Nº", "Nº Guía", "Codigo", "Descripcion", "Laboratorio", "UND", "Cantidad")
    arrAncho = Array(600, 1200, 1200, 2200, 2200, 800, 800)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgRight)
    Me.grdDetGuia.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    Me.grdDetGuia.Columns(0).Merge = False
    Me.grdDetGuia.Columns(1).Merge = False
    Me.grdDetGuia.Columns(2).Merge = False
    Me.grdDetGuia.Columns(3).Merge = False
    Me.grdDetGuia.Columns(4).Merge = False
    Me.grdDetGuia.Columns(5).Merge = False
    Me.grdDetGuia.Columns(6).Merge = False
End Sub

Private Sub Command1_Click()
    frm_ADM_TranspGuia.carga strIdEntrega, 0
End Sub

Private Sub Command2_Click()
    Dim msbo As Variant
    If grdCabGuia.ApproxCount <= 0 Then
        MsgBox "No exiten guias que desasociar.", vbCritical, "Error"
        Exit Sub
    End If
    msbo = MsgBox("¿Seguro que desea desasociar las guía Nº " & Me.grdCabGuia.Columns("NUM_GUIA").Value & "?", vbYesNo + vbInformation, App.ProductName)
    If msbo = vbYes Then
        objEntrega.Desasociar strIdEntrega, "" & Me.grdCabGuia.Columns("NUM_GUIA").Value
        Me.grdCabGuia.Limpiar
        Me.grdDetGuia.Limpiar
        cargaCabGuias
    End If
End Sub

Private Sub Command3_Click()
    frm_ADM_Entrega.Consulta
    frm_ADM_Entrega.grdRecepcion.DataSource.FindFirst "ID_ENTREGA='" & Trim(strIdEntrega) & "'"
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Command3_Click
End If
If KeyCode = vbKeyF1 Then
    Command1_Click
End If
If KeyCode = vbKeyF2 Then
    Command2_Click
End If
End Sub

Private Sub Form_Load()
cargaCabGuias
End Sub

Private Sub grdCabGuia_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If Me.grdCabGuia.ApproxCount > 0 Then
        cargaDetGuias (Me.grdCabGuia.Columns("NUM_GUIA"))
    End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.grdDetGuia.DataSource.FindFirst "COD_PRODUCTO='" & Trim(Me.txtbuscar.Text) & "'"
    End If
End Sub
