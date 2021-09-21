VERSION 5.00
Begin VB.Form frm_ADM_Conteo 
   Caption         =   "Revisión"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "[+] Modificar Cantidad"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "[Esc] Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "[F11] Finalizar"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
   End
   Begin vbp_Ventas.ctlGrillaArray ctlGrilla1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   14843
      _ExtentY        =   873
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1830
   End
End
Attribute VB_Name = "frm_ADM_Conteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As New clsEntrega
Dim objproducto As New clsProducto
Public strEntrega As String
Dim strReconteo As String
Private xarrDetalle As New XArrayDB

Sub agregarCant()
If Me.ctlGrilla1.Columns("item").Value = 0 Then
    frm_ADM_AgregaCantidad.codBarra = Me.ctlGrilla1.Columns("Codigo").Value
    frm_ADM_AgregaCantidad.cantidad = Me.ctlGrilla1.Columns("cantidad").Value
    frm_ADM_AgregaCantidad.Show vbModal
Else
    MsgBox "Solo se puede Modificar el Ultimo Producto Ingresado", vbCritical, "Error"
End If
End Sub

Sub cerrar()
Me.ctlGrilla1.Update
Dim numGuiaDev As String
If MsgBox("¿Seguro que desea cerrar el Conteo?", vbYesNo, App.ProductName) = vbYes Then
    Dim i As Integer
    For i = 0 To xarrDetalle.Count(1) - 1
        obj.AgregaProducto strEntrega, xarrDetalle(i, 0), xarrDetalle(i, 3), "0"
    Next
    obj.Cierra strEntrega, objUsuario.Codigo, "0", objUsuario.NombrePC, numGuiaDev
    frm_ADM_Entrega.Form_Load
    Unload Me
End If
End Sub

Private Sub cmdAceptar_Click()
    cerrar
End Sub

Private Sub cmdAgregar_Click()
    agregarCant
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub ctlGrilla1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'If Me.ctlGrilla1.Columns("item").Value <> 0 Then
'    MsgBox "No se puede editar esta fila", vbCritical, "Restricción"
'    Me.ctlGrilla1.Columns("Cantidad").Value = OldValue
'Else
'    If Me.ctlGrilla1.Columns("Cantidad").Value = "" Then
'        MsgBox "Ingrese Cantidad Valida", vbCritical, "Restricción"
'        Me.ctlGrilla1.Columns("Cantidad").Value = OldValue
'    End If
'End If
End Sub

Private Sub ctlGrilla1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cantidad As String
cantidad = Me.ctlGrilla1.Columns("cantidad").Value

'If KeyCode = 13 Then
'    Me.ctlGrilla1.AllowUpdate = False
'End If
End Sub

Private Sub ctlTextBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim rs As oraDynaset
    Set rs = objproducto.ListaBusqueda(Trim(Me.ctlTextBox1.Text)) 'obj.ListaProducto2(strEntrega, Trim(ctlTextBox1.Text), "0")
    If rs(0) <> -1 Then
        agregar "" & rs("COD"), "" & rs("DES"), "" & rs("DES_LAB"), "1"
    Else
        MsgBox "No se encontraron Productos asociados", vbCritical + vbInformation, "Aviso"
    End If
    ctlTextBox1.Text = ""
    ctlTextBox1.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
 End If
 If KeyCode = 107 Then
    agregarCant
End If
If KeyCode = vbKeyF11 Then
    cerrar
End If
End Sub

Private Sub Form_Load()
    SeteaGrilla
    xarrDetalle.ReDim 0, -1, 0, 5
    ctlGrilla1.Array1 = xarrDetalle
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "CTD_PRODUCTO", "ITEM")
    arrCaption = Array("Codigo", "Producto", "Laboratorio", "Cantidad", "item")
    arrAncho = Array(800, 3500, 3500, 500, 500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    arrFoco = Array(False, False, False, True, False)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco

Me.ctlGrilla1.Columns("item").Visible = False
End Sub

Public Function AgregaDetalle(ByVal codProducto As String, _
                                    ByVal desProducto As String, _
                                    ByVal deslaboratorio As String, _
                                    ByVal cantidad As Integer) As XArrayDB
    Dim ultimo As Integer
    Dim aux As Integer
    If xarrDetalle.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    encontro = False
    ultimo = xarrDetalle.Count(1)
    
    aux = xarrDetalle.Count(1)

    If xarrDetalle.Count(1) = 0 Then
    ultimo = 0
    End If
    xarrDetalle.AppendRows
    
    If xarrDetalle.Count(1) > 0 Then
    
    Dim j As Integer
    j = xarrDetalle.Count(1) - 1
    
    While (j > 0)
        xarrDetalle(j, 0) = xarrDetalle(j - 1, 0)
        xarrDetalle(j, 1) = xarrDetalle(j - 1, 1)
        xarrDetalle(j, 2) = xarrDetalle(j - 1, 2)
        xarrDetalle(j, 3) = xarrDetalle(j - 1, 3)
        xarrDetalle(j, 4) = xarrDetalle(j - 1, 4) + 1
        j = j - 1
    Wend
      
    End If
        
    xarrDetalle(0, 0) = codProducto
    xarrDetalle(0, 1) = desProducto
    xarrDetalle(0, 2) = deslaboratorio
    xarrDetalle(0, 3) = cantidad
    xarrDetalle(0, 4) = 0
    
    Set AgregaDetalle = xarrDetalle

End Function

Public Sub eliminarDetalle()
If xarrDetalle.Count(1) > 0 Then
    Dim j As Integer
    j = 0
    While (j < (xarrDetalle.Count(1) - 2))
        xarrDetalle(j, 0) = xarrDetalle(j + 1, 0)
        xarrDetalle(j, 1) = xarrDetalle(j + 1, 1)
        xarrDetalle(j, 2) = xarrDetalle(j + 1, 2)
        xarrDetalle(j, 3) = xarrDetalle(j + 1, 3)
        xarrDetalle(j, 4) = xarrDetalle(j + 1, 4) - 1
        j = j + 1
    Wend
End If
xarrDetalle.DeleteRows (xarrDetalle.Count(1) - 1)
End Sub

Sub agregar(ByVal Codigo As String, ByVal Descripcion As String, ByVal Laboratorio As String, ByVal cantidad As String)
    AgregaDetalle Codigo, Descripcion, Laboratorio, "1"
    ctlGrilla1.Rebind
End Sub
