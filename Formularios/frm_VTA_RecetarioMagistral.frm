VERSION 5.00
Begin VB.Form frm_VTA_RecetarioMagistral 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlDataCombo ctlCboTipo 
      Height          =   315
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   615
      Left            =   5640
      Picture         =   "frm_VTA_RecetarioMagistral.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_RecetarioMagistral.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_RecetarioMagistral.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrilla grdInsumo 
      Height          =   2775
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4895
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtConcentracion 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   4950
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
   Begin vbp_Ventas.ctlTextBox txtCant 
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   5385
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Tipo            =   4
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
   Begin vbp_Ventas.ctlTextBox txtCtdBase 
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   4515
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Tipo            =   4
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
   Begin vbp_Ventas.ctlDataCombo ctlCboUnidadMedida 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   4080
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      MatchEntry      =   1
      Enabled         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtBuscar 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   556
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
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
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblTipo 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
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
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frm_VTA_RecetarioMagistral.frx":109E
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingreso de Insumos"
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
      Left            =   480
      TabIndex        =   18
      Top             =   120
      Width           =   2085
   End
   Begin VB.Label lblInsumo 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Index           =   0
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad Medida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   4140
      Width           =   1290
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   15
      Top             =   4575
      Width           =   1245
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   765
      TabIndex        =   14
      Top             =   5445
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Concentración %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   5010
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   975
      TabIndex        =   12
      Top             =   5880
      Width           =   555
   End
   Begin VB.Label lblPrecio 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   5820
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_RecetarioMagistral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Dim objRecetarioMagistral As New clsRecetarioMagistral
Dim strCadCodIns As String
Dim strCadDesIns As String
Dim strCadCodigo As String
Dim strCadProd As String
Dim strCadUnd As String
Dim strCadPctMargen As String

Dim strCadCant As String
Dim strCadPrecio As String
Dim strCadSubTotal As String
Dim strCod As String
Dim strCadCodUnd As String
Dim strCantAnt As String
Public pstrRucProv As String
Public strCtdBase As String
Public strCodInsumo As String
Public strConcentracion As String
Public strCantidad As String



Private Sub cmdAceptar_Click()


On Error GoTo handle

    strCadCodIns = grdInsumo.Columns("COD_TIPO_INSUMO").Value
    strCadDesIns = grdInsumo.Columns("DES_TIPO_INSUMO").Value
    strCadCodigo = grdInsumo.Columns("COD_PRODUCTO").Value
    strCadProd = grdInsumo.Columns("DES_PRODUCTO").Value
    strCadPrecio = grdInsumo.Columns("IMP_PRECIO_VTA").Value
    strCadUnd = grdInsumo.Columns("DES_UND_CAPACIDAD_ABREV").Value
    strCadCodUnd = grdInsumo.Columns("COD_UND_CAPACIDAD").Value


    
    
    Select Case strCadCodIns
        Case "001"
            If Val(txtCtdBase.Text) = 0 Then
                MsgBox "Es necesario indicar cantidad base", vbCritical + vbOKOnly, App.ProductName
                txtCtdBase.SetFocus
                Exit Sub
            End If
            strCadPctMargen = 0
            strCtdBase = ""
            strCadCant = Val(txtCtdBase.Text)
            strCadSubTotal = strCadPrecio * strCadCant
            strCantAnt = strCadCant
        Case "002", "003", "004"
            If Val(txtCtdBase.Text) = 0 Then
                MsgBox "La cantidad base no puede ser CERO", vbCritical + vbOKOnly, App.ProductName
                Exit Sub
            End If
            If Val(txtConcentracion.Text) = 0 And Val(txtCant.Text) = 0 Then
                MsgBox "Es necesario indicar cantidad ó concentración", vbCritical + vbOKOnly, App.ProductName
                txtConcentracion.SetFocus
                Exit Sub
            End If
        
            If strCantAnt <> "" Then
                strCtdBase = strCantAnt
                strCadCant = Val(txtCant.Text)
                strCadPctMargen = (strCadCant / strCtdBase) * 100
                strCadSubTotal = strCadPrecio * strCadCant
            Else
                'Esta linea de validacion se agrego el dia 21/03/2007'
                If frm_VTA_RecetarioM.GrdInsumos.ApproxCount <= 0 Then strCtdBase = ""
                
                strCadPctMargen = "" 'strCadSubTotal = ""
                strCadCant = Val(txtCant.Text)
                strCadSubTotal = strCadPrecio * strCadCant
            End If
        Case Else
            If frm_VTA_RecetarioM.GrdInsumos.ApproxCount <= 0 Then strCtdBase = ""
            strCadPctMargen = "" ' strCadSubTotal = ""
            strCadCant = Val(txtCant.Text)
            strCadSubTotal = strCadPrecio * strCadCant
    End Select


                If Val(strCadSubTotal) = 0 Then
                    MsgBox "el Importe Sub Total no puede ser CERO" + Chr(13) + "revisar las cantidad base o la concentración", vbCritical + vbOKOnly, App.ProductName
                    Exit Sub
                End If


                strCod = objProducto.ListaDevRM(objUsuario.CodigoLocal, pstrRucProv)





                Call frm_VTA_RecetarioM.psub_Agrega_Insumo(strCadCodIns, strCadDesIns, _
                                                           strCadCodigo, strCadProd, _
                                                           strCadUnd, strCadPctMargen, _
                                                           strCtdBase, strCadCant, _
                                                           strCadPrecio, strCadSubTotal, _
                                                           strCod, strCadCodUnd)
                                                           
                frm_VTA_RecetarioM.GrdInsumos.Rebind
                frm_VTA_RecetarioM.CantBase = txtCtdBase.Text
                                                           
    Unload Me
    Exit Sub

handle:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.ProductName

End Sub

Private Sub cmdBuscar_Click()

Buscar

                                                                    
End Sub




Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub ctlCboTipo_Click(Area As Integer)
    strCodInsumo = ctlCboTipo.BoundText
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handle
    ''''psub_KeyDownAplicacion KeyCode, Shift
    
    Select Case KeyCode
        Case vbKeyF1
            txtBuscar.SetFocus
        Case vbKeyF2
            grdInsumo.SetFocus
            
    End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub Form_Load()
    Limpia
    Inicio
End Sub

Private Sub grdInsumo_DblClick()
    If grdInsumo.ApproxCount = 0 Then Exit Sub
    CargarValores
     SendKeys "{TAB}"

End Sub


Public Sub CargarValores()

    strCadCodIns = grdInsumo.Columns("COD_TIPO_INSUMO").Value
    strCadDesIns = grdInsumo.Columns("DES_TIPO_INSUMO").Value
    strCadCodigo = grdInsumo.Columns("COD_PRODUCTO").Value
    strCadProd = grdInsumo.Columns("DES_PRODUCTO").Value
    strCadPrecio = grdInsumo.Columns("IMP_PRECIO_VTA").Value
    strCadUnd = grdInsumo.Columns("DES_UND_CAPACIDAD_ABREV").Value
    strCadCodUnd = grdInsumo.Columns("COD_UND_CAPACIDAD").Value

    txtCtdBase.Text = frm_VTA_RecetarioM.CantBase
    txtConcentracion.Text = strConcentracion
    txtCant.Text = strCantidad
    
    lblInsumo.Caption = strCadProd
    lblTipo.Caption = strCadDesIns
    lblPrecio.Caption = Format(Val(strCadPrecio), "#0.00")

    ctlCboUnidadMedida.BoundText = strCadCodUnd
    
    
    
    

End Sub



Private Sub Inicio()

    SetteaFormulario Me
    SetGrd
    
    Set ctlCboUnidadMedida.RowSource = objRecetarioMagistral.ListaUndCapacidad
    ctlCboUnidadMedida.ListField = "DES_UND_CAPACIDAD"
    ctlCboUnidadMedida.BoundColumn = "COD_UND_CAPACIDAD"
    
    
    Set ctlCboTipo.RowSource = objRecetarioMagistral.ListaTipoInsumo
    ctlCboTipo.ListField = "DES_TIPO_INSUMO"
    ctlCboTipo.BoundColumn = "COD_TIPO_INSUMO"
    
    ctlCboTipo.BoundText = strCodInsumo
    

End Sub

Private Sub grdInsumo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdInsumo_DblClick
    End Select
End Sub



Private Sub SetGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim i As Integer


    arrCampos = Array("COD_TIPO_INSUMO", "DES_TIPO_INSUMO", "COD_PRODUCTO", "DES_PRODUCTO", "DES_CLASE_COM", "DES_CATEGORIA_COM", "DES_UND_CAPACIDAD_ABREV", "COD_UND_CAPACIDAD", "PCT_MARGEN", "IMP_COSTO_UNI", "IMP_PRECIO_VTA")
    arrCaption = Array("CodInsumo", "Tipo", "Código", "Descripción", "Clase", "Categoría", "UM", "CodCapacidad", "% Margen", "Costo", "Precio")
    arrAncho = Array(0, 900, 800, 2500, 0, 0, 500, 0, 0, 0, 650)
    arrAlineacion = Array(dbgGeneral, dbgGeneral, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgGeneral, dbgGeneral)
    grdInsumo.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdInsumo.Columns("COD_TIPO_INSUMO").Visible = False
    grdInsumo.Columns("DES_CLASE_COM").Visible = False
    grdInsumo.Columns("DES_CATEGORIA_COM").Visible = False
    grdInsumo.Columns("COD_UND_CAPACIDAD").Visible = False
    grdInsumo.Columns("PCT_MARGEN").Visible = False
    grdInsumo.Columns("IMP_COSTO_UNI").Visible = False
    
    
    For i = 0 To grdInsumo.Columns.Count - 1
        grdInsumo.Columns(i).AllowFocus = False
    Next i
    
    

End Sub


Private Sub Limpia()
lblTipo.Caption = ""
lblInsumo.Caption = ""
lblPrecio.Caption = ""
txtCtdBase.Text = ""
txtConcentracion.Text = ""
txtCant.Text = ""
ctlCboUnidadMedida.BoundText = -1
End Sub


Public Sub Buscar()
   Set grdInsumo.DataSource = objProducto.ListaRegMagistral(objUsuario.CodigoEmpresa, _
                                                                    objUsuario.CodigoLocal, _
                                                                    Trim(txtBuscar.Text), _
                                                                    frm_VTA_RecetarioM.pstrRucProv, _
                                                                    strCodInsumo)
    If grdInsumo.ApproxCount = 0 Then
        Limpia
        SendKeys "+{TAB}"
    Else
        SendKeys "{TAB}"
    End If

End Sub

Private Sub grdInsumo_RegistroSeleccionado(ByVal DatoColumna0 As String)
            If DatoColumna0 = objRecetarioMagistral.CodInsumoBase Then
                txtConcentracion.Bloqueado = True
                txtCant.Bloqueado = True
                '''txtCtdBase.Bloqueado = False
            Else
                txtConcentracion.Bloqueado = False
                txtCant.Bloqueado = False
                '''txtCtdBase.Bloqueado = True
            End If
End Sub


Private Sub txtCant_Change()
' Formula para sacar el "Pct"  =>  (Cant / Base )*100 '
On Error GoTo handle
txtConcentracion.Text = Round((txtCant.Text / txtCtdBase.Text) * 100, 4)

Exit Sub
handle:
txtConcentracion.Text = ""
    
End Sub


Private Sub txtConcentracion_Change()
'(Pct / 100 )* Ctd Base


 On Error GoTo handle
txtCant.Text = Round((txtConcentracion.Text / 100) * txtCtdBase.Text, 4)

Exit Sub

handle:
    txtCant.Text = ""

End Sub

Private Sub txtConcentracion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtConcentracion_Change
    End If
End Sub
