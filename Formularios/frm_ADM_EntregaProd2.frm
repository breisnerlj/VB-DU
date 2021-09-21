VERSION 5.00
Begin VB.Form frm_ADM_EntregaProd2 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrillaArray ctlGrilla1 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11245
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F1-Finalizar Conteo"
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
      Left            =   240
      TabIndex        =   2
      Top             =   7440
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Producto:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   810
   End
End
Attribute VB_Name = "frm_ADM_EntregaProd2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As New clsEntrega
Dim strEntrega As String
Dim strReconteo As String
Dim xarrDetalle As XArrayDB


Public Function AgregaListaConteo(Codigo As String, _
                            Descripcion As String, _
                            Laboratorio As String, _
                            Cantidad As Integer) As XArrayDB
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
    xarrDetalle(ultimo, 3) = Val(Cantidad)
        
    Set AgregaListaConteo = xarrDetalle
    
End Function

Public Sub carga(ByVal Entrega As String, ByVal Reconteo As String, Optional codProducto As String = "0", Optional inicio As Integer = 0)
    strReconteo = Reconteo
    strEntrega = Entrega
    Dim odyn As oraDynaset
    
    Set odyn = obj.ListaConteoOrden(Entrega, Reconteo, codProducto)
    Set xarrDetalle = New XArrayDB
    'xarrDetalle.ReDim 0, -1, 0, 5
    LimpiarArray
    SeteaGrilla
    odyn.MoveFirst
    While Not odyn.EOF
        AgregaListaConteo odyn("COD_PRODUCTO").Value, _
                          odyn("DES_PRODUCTO").Value, _
                          odyn("DES_LABORATORIO").Value, _
                          odyn("CTD_PRODUCTO").Value
        odyn.MoveNext
    Wend
    ctlGrilla1.Array1 = xarrDetalle
    Me.ctlGrilla1.Rebind
    Me.Caption = "Esta realizando la entrega :" & Entrega
    If Reconteo = "1" Then
     Label2.Caption = "F1-Finalizar ReConteo"
    Else
     Label2.Caption = "F1-Finalizar Conteo"
    End If
    If inicio = 0 Then
        Me.Show vbModal
    End If
    
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

Private Sub ctlGrilla1_AfterColUpdate(ByVal ColIndex As Integer)
   Agregar "" & Me.ctlGrilla1.Columns(0).CellText(ctlGrilla1.Bookmark), "" & Me.ctlGrilla1.Columns(1).CellText(ctlGrilla1.Bookmark), "" & Me.ctlGrilla1.Columns(2).CellText(ctlGrilla1.Bookmark), "" & Me.ctlGrilla1.Columns(3).CellText(ctlGrilla1.Bookmark)
   Me.ctlTextBox1.Enabled = True
   Me.ctlTextBox1.Text = ""
   'Me.ctlGrilla1.Enabled = False
   Me.ctlTextBox1.SetFocus
End Sub


Private Sub ctlTextBox1_KeyPress(KeyAscii As Integer)
Dim rs As oraDynaset
    If KeyAscii = vbKeyReturn Then
        If Me.ctlGrilla1.ApproxCount > 0 Then
            Set rs = obj.ListaProducto2(strEntrega, Trim(ctlTextBox1.Text), strReconteo)
                
                If rs.EOF Then
                    MsgBox "El Producto no existe "
                    Me.ctlTextBox1.SetFocus
                Else
                    Me.ctlTextBox1.Enabled = False
                    carga strEntrega, strReconteo, Trim(Me.ctlTextBox1.Text), 1
                    'Me.ctlGrilla1.Enabled = True
                    Me.ctlGrilla1.Bookmark = 0
                    
                End If
        Else
            MsgBox "No se encontraron items en la grilla.", vbCritical, "Error"
            Me.ctlTextBox1.Text = ""
            Me.ctlTextBox1.SetFocus
        End If
    End If
    Dim mensaje As String
'    If KeyCode = vbKeyF1 Then
'        If strReconteo = 0 Then
'            mensaje = "Esta seguro de cerrar el conteo."
'        Else
'            mensaje = "Esta seguro de cerrar el conteo." + Chr(13) + "También se procederá a cerrar la recepción."
'        End If
'        If MsgBox(mensaje, vbYesNo, App.ProductName) = vbYes Then
'            obj.Cierra strEntrega, objUsuario.Codigo, strReconteo, objUsuario.NombrePC
'            Unload Me
'        End If
'    End If
End Sub

Sub Agregar(ByVal Codigo As String, ByVal Descripcion As String, ByVal Laboratorio As String, ByVal Cantidad As String)
    Dim rs As oraDynaset
    Dim i As Integer
    Set rs = obj.AgregaProducto(strEntrega, Codigo, IIf(Cantidad = "", 1, Cantidad), strReconteo, "2")
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
    arrFoco = Array(False, False, False, True)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    ctlGrilla1.AllowUpdate = True
    ctlGrilla1.Columns(0).Merge = True
    ctlGrilla1.Columns(1).Merge = True
    ctlGrilla1.MarqueeStyle = dbgHighlightCell
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    'agregaValores
End If
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
        
        frm_ADM_Sobrantes.IdEntrega = strEntrega
        frm_ADM_Sobrantes.Show vbModal
        
        Unload Me
    End If
End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub

