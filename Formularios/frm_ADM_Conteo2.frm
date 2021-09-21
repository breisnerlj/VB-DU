VERSION 5.00
Begin VB.Form frm_ADM_Conteo2 
   Caption         =   "Reconteo"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8705
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "[F11] Finalizar Verificación"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "[F5] Imprimir"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "[Esc] Salir"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "[Enter] Ingresar Cantidad"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Sr(a). Jefe de Local, Verifique la cantidad recibida de los siguientes productos"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "frm_ADM_Conteo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As New clsEntrega
Public strEntrega As String
Dim strReconteo As String
Dim xarrDetalle As XArrayDB

Sub agregarCantidad()
    frm_ADM_AgregaCantidad2.strEntrega = strEntrega
    frm_ADM_AgregaCantidad2.codproducto = Me.ctlGrilla1.Columns("Codigo").Value
    frm_ADM_AgregaCantidad2.desProducto = Me.ctlGrilla1.Columns("Producto").Value
    frm_ADM_AgregaCantidad2.desUnidad = Me.ctlGrilla1.Columns("Laboratorio").Value
    frm_ADM_AgregaCantidad2.Show vbModal
End Sub

Public Function AgregaListaConteo(Codigo As String, _
                            Descripcion As String, _
                            Laboratorio As String, _
                            cantidad As String) As XArrayDB
    
    If cantidad = "" Then
        cantidad = "0"
    End If
    
    Dim ultimo As Integer
    Dim aux As Integer
    If xarrDetalle.Count(1) < 0 Then Exit Function
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xarrDetalle.Count(1)
    While i < aux
        If xarrDetalle(i, 0) = Codigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xarrDetalle.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xarrDetalle.AppendRows
    End If
j:
    If xarrDetalle.Count(1) = 0 Then ultimo = 0: xarrDetalle.AppendRows
    
    xarrDetalle(ultimo, 0) = Codigo
    xarrDetalle(ultimo, 1) = Descripcion
    xarrDetalle(ultimo, 2) = Laboratorio
    xarrDetalle(ultimo, 3) = Val(cantidad)
        
    Set AgregaListaConteo = xarrDetalle
    
End Function

Public Sub carga(ByVal Entrega As String, ByVal Reconteo As String, Optional codproducto As String = "0", Optional inicio As Integer = 0)
On Error GoTo CtrlErr
    
    strReconteo = Reconteo
    strEntrega = Entrega
    Dim odyn As oraDynaset
    
    Set odyn = obj.ListaConteoOrden(Entrega, Reconteo, codproducto)
    Set xarrDetalle = New XArrayDB
    'xarrDetalle.ReDim 0, -1, 0, 5
    LimpiarArray
    SeteaGrilla
    odyn.MoveFirst
    While Not odyn.EOF
        AgregaListaConteo "" & odyn("COD_PRODUCTO").Value, _
                          "" & odyn("DES_PRODUCTO").Value, _
                          "" & odyn("DES_LABORATORIO").Value, _
                          "" & odyn("CTD_PRODUCTO").Value
        odyn.MoveNext
    Wend
    
'    ctlGrilla1.Array1 = xarrDetalle
'    Me.ctlGrilla1.Rebind
    
    Set Me.ctlGrilla1.DataSource = odyn
    Me.ctlGrilla1.MoveFirst
    
    Me.Caption = "Esta realizando la entrega :" & Entrega
    
    If inicio = 0 Then
        Me.Show vbModal
    End If
    
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub

Private Function validaConteo() As Boolean
Dim j As Integer
Dim i As String
Dim Valor As Boolean
Valor = True
If Me.ctlGrilla1.ApproxCount > 0 Then
Me.ctlGrilla1.MoveFirst
    For j = 0 To Me.ctlGrilla1.ApproxCount - 1
        i = Me.ctlGrilla1.Columns(3).Value
        If i = "0" Then
            Valor = False
        End If
        Me.ctlGrilla1.MoveNext
    Next
End If
Me.ctlGrilla1.MoveFirst
validaConteo = Valor
End Function

Private Sub cmdAgregar_Click()
    agregarCantidad
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Sub finalizar()
On Error GoTo CtrlErr
    Me.ctlGrilla1.MoveFirst
    Dim i As Integer
    For i = 0 To Me.ctlGrilla1.ApproxCount - 1
        If Me.ctlGrilla1.Columns("Cantidad").Value = "" Then
            MsgBox "No se ha ingresado cantidad al producto " + Me.ctlGrilla1.Columns("Producto").Value, vbCritical + vbInformation, "Error"
            Exit Sub
        End If
        Me.ctlGrilla1.MoveNext
    Next

    Dim numGuiaDev As String
    Dim mensaje As String
    mensaje = "¿Esta seguro de confirmar la cantidad ingresada para cada producto?" + Chr(13) + _
    "Si continúa ya no podrá modificar las cantidades." + Chr(13) + _
    "¿Desea continuar con el proceso?"
    If MsgBox(mensaje, vbYesNo, App.ProductName) = vbYes Then
        obj.Cierra strEntrega, objUsuario.Codigo, "1", objUsuario.NombrePC, numGuiaDev
        'Call obj.GrabaFaltantesYSobrantes(strEntrega, "Faltante")
        'Call obj.GrabaFaltantesYSobrantes(strEntrega, "Sobrante")
        'Call obj.GrabaFaltantesYSobrantes(strEntrega, "@")
        'Falta uno? 28/11/2012
        frm_ADM_Entrega.Form_Load
        frm_ADM_Sobrantes.idEntrega = strEntrega
        frm_ADM_Sobrantes.Show vbModal
        Unload Me
    End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub

Private Sub cmdFinalizar_Click()
    finalizar
End Sub

Private Sub cmdImprimir_Click()
'Me.ctlGrilla2.MostrarImprimir
End Sub

Private Sub ctlGrilla1_AfterColUpdate(ByVal ColIndex As Integer)
   agregar "" & Me.ctlGrilla1.Columns(0).CellText(ctlGrilla1.Bookmark), "" & Me.ctlGrilla1.Columns(1).CellText(ctlGrilla1.Bookmark), "" & Me.ctlGrilla1.Columns(2).CellText(ctlGrilla1.Bookmark), "" & Me.ctlGrilla1.Columns(3).CellText(ctlGrilla1.Bookmark)
   'Me.ctlTextBox1.Enabled = True
   'Me.ctlTextBox1.Text = ""
   'Me.ctlGrilla1.Enabled = False
   'Me.ctlTextBox1.SetFocus
End Sub


Sub agregar(ByVal Codigo As String, ByVal Descripcion As String, ByVal Laboratorio As String, ByVal cantidad As String)
    Dim rs As oraDynaset
    Dim i As Integer
    
On Error GoTo CtrlErr
    
    Set rs = obj.AgregaProducto(strEntrega, Codigo, IIf(cantidad = "", 1, cantidad), strReconteo, "2")
    LimpiarArray
    SeteaGrilla
    rs.MoveFirst
    While Not rs.EOF
        AgregaListaConteo rs("COD_PRODUCTO").Value, _
                          rs("DES_PRODUCTO").Value, _
                          rs("DES_LABORATORIO").Value, _
                          rs("CTD_PRODUCTO").Value
        rs.MoveNext
    Wend
    ctlGrilla1.Array1 = xarrDetalle
    Me.ctlGrilla1.Rebind
    'Set ctlGrilla1.DataSource = rs
'    ctlGrilla1.Rebind

Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub

Public Sub LimpiarArray()
    xarrDetalle.Clear
    xarrDetalle.ReDim 0, -1, 0, 5
End Sub


Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "CTD_PRODUCTO")
    arrCaption = Array("Codigo", "Producto", "Laboratorio", "Cantidad")
    arrAncho = Array(800, 3500, 3500, 500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    'ctlGrilla1.AllowUpdate = True
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    'agregaValores
End If
End Sub

Public Sub recibe(ByVal codproducto As String, ByVal cantidad As String)
On Error GoTo CtrlErr
   
    obj.AgregaProducto strEntrega, codproducto, cantidad, "1", "2"
    carga strEntrega, "1", , 1
    Me.ctlGrilla1.DataSource.FindFirst "COD_PRODUCTO='" & Trim(codproducto) & "'"
    Me.ctlGrilla1.MoveNext
    Unload frm_ADM_AgregaCantidad2
    Me.ctlGrilla1.SetFocus

Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
Dim mensaje As String
Dim numGuiaDev As String
'Dim objDocumento As New clsGuia
Dim objDocumento As New clsDocumento
Dim arr() As String

If KeyCode = vbKeyF1 Then
    If strReconteo = 0 Then
        mensaje = "Esta seguro de cerrar el conteo."
    Else
        mensaje = "Esta seguro de cerrar el conteo." + Chr(13) + "También se procederá a cerrar la recepción."
    End If
    If MsgBox(mensaje, vbYesNo, App.ProductName) = vbYes Then
        obj.Cierra strEntrega, objUsuario.Codigo, strReconteo, objUsuario.NombrePC, numGuiaDev
        MsgBox "Se ha generado la(s) Guía(s) Nº " + Replace(numGuiaDev, "|", " - ") + "." + Chr(13) + "Verifique su Impresora", vbInformation, "Guías de Devolución"
        arr = Split(numGuiaDev, "|")
        Dim j As Integer
        For j = 0 To UBound(arr)
            'objDocumento.spImprime_Guia_Dev "", "", arr(j)
            objDocumento.ImprimirGuiaTransferencia arr(j)
        Next
        
        frm_ADM_Sobrantes.idEntrega = strEntrega
        frm_ADM_Sobrantes.Show vbModal
        
        Unload Me
    End If
End If

If KeyCode = 13 Then
    agregarCantidad
End If

If KeyCode = 27 Then
    Unload Me
End If

If KeyCode = vbKeyF5 Then
    Me.ctlGrilla1.MostrarImprimir
End If

If KeyCode = vbKeyF11 Then
    finalizar
End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub
