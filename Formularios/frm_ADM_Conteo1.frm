VERSION 5.00
Begin VB.Form frm_ADM_Conteo1 
   Caption         =   "Conteo"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9551
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "[F11] Finalizar"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "[Esc] Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "[+] Modificar Cantidad"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   6360
      Width           =   2175
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Codigo de Barras:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1830
   End
End
Attribute VB_Name = "frm_ADM_Conteo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj As New clsEntrega
Public strIdEntrega As String
Dim objproducto As New clsProducto
Dim momento As Integer

Private Sub cmdAceptar_Click()
    cerrar
End Sub

Private Sub cmdAgregar_Click()
    agregarCant
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub


Private Sub ctlTextBox1_KeyPress(KeyAscii As Integer)
On Error GoTo CtrlError
If KeyAscii = 13 Then
    Dim rs As oraDynaset
    Set rs = objproducto.ListaBusqueda(Trim(Me.ctlTextBox1.Text))
    If rs(0) <> -1 Then
        obj.GrabaConteoAux strIdEntrega, "" & rs("COD"), "1", momento, "0"
        Me.ctlGrilla1.Limpiar
        Set Me.ctlGrilla1.DataSource = obj.ListaConteoAux(strIdEntrega, momento)
        ctlTextBox1.Text = ""
        Me.cmdCancelar.SetFocus
    Else
        MsgBox "No se encontraron Productos asociados", vbCritical + vbInformation, "Aviso"
        ctlTextBox1.Text = ""
    End If
    
End If
If KeyAscii = 43 Then
    KeyAscii = 8
End If
Exit Sub
CtrlError:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Then
        agregarCant
    End If
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = vbKeyF11 Then
        cerrar
    End If
End Sub

Public Sub CargaGrilla()
On Error GoTo Handle
    Me.ctlGrilla1.Limpiar
    SeteaGrilla
    Set Me.ctlGrilla1.DataSource = obj.ListaConteoAux(strIdEntrega, momento)
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Dim rs As oraDynaset
    Set rs = obj.momento(strIdEntrega)
    momento = Val(rs(0)) + 1
    CargaGrilla
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
        
End Sub

Sub agregarCant()
On Error GoTo Handle
If Me.ctlGrilla1.Columns("item").Value = 0 And Me.ctlGrilla1.Columns("bloqueado").Value = 0 Then
    frm_ADM_AgregaCantidad.codBarra = Me.ctlGrilla1.Columns("Codigo").Value
    frm_ADM_AgregaCantidad.codproducto = Me.ctlGrilla1.Columns("Codigo").Value
    frm_ADM_AgregaCantidad.cantidad = Me.ctlGrilla1.Columns("cantidad").Value
    frm_ADM_AgregaCantidad.strEntrega = strIdEntrega
    frm_ADM_AgregaCantidad.Show vbModal
Else
    MsgBox "Solo se puede Modificar el Ultimo Producto Ingresado", vbCritical, "Error"
End If
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "CANTIDAD", "ITEM", "FLG_BLOQUEADO")
    arrCaption = Array("Codigo", "Producto", "Laboratorio", "Cantidad", "item", "bloqueado")
    arrAncho = Array(800, 3500, 3500, 500, 500, 500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    arrFoco = Array(False, False, False, True, False, False)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco

Me.ctlGrilla1.Columns("item").Visible = False
Me.ctlGrilla1.Columns("bloqueado").Visible = False
End Sub

Sub cerrar()
'Me.ctlGrilla1.Update
Dim numGuiaDev As String
If MsgBox("¿Seguro que desea cerrar el Conteo?", vbYesNo, App.ProductName) = vbYes Then
    Dim rs As oraDynaset
    Set rs = obj.ConsolidaConteoAux(strIdEntrega)
        
    rs.MoveFirst
    While Not rs.EOF
        obj.AgregaProducto strIdEntrega, "" & rs("COD_PRODUCTO"), "" & rs("CTD_TOTAL"), "0"
        rs.MoveNext
    Wend
        
    obj.Cierra strIdEntrega, objUsuario.codigo, "0", objUsuario.NombrePC, numGuiaDev
    frm_ADM_Entrega.Form_Load
    frm_ADM_Entrega.grdRecepcion.DataSource.FindFirst "ID_ENTREGA='" & Trim(strIdEntrega) & "'"
    Unload Me
End If

End Sub

Sub recibe(codproducto As String, cantidad As String)

If cantidad = "0" Then
    obj.EliminaConteoAux strIdEntrega, codproducto
Else
    obj.EditaConteoAux strIdEntrega, codproducto, cantidad, "1"
End If
Me.CargaGrilla
Me.ctlGrilla1.MoveFirst
Unload frm_ADM_AgregaCantidad

Me.cmdCancelar.SetFocus
Me.ctlTextBox1.Text = ""

End Sub
