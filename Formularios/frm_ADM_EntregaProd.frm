VERSION 5.00
Begin VB.Form frm_ADM_EntregaProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   8415
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
      TabIndex        =   3
      Top             =   7440
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Producto:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   690
   End
End
Attribute VB_Name = "frm_ADM_EntregaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As New clsEntrega
Dim strEntrega As String
Dim strReconteo As String
Dim xarrDetalle As XArrayDB
Public Sub carga(ByVal Entrega As String, ByVal Reconteo As String)
    strReconteo = Reconteo
    strEntrega = Entrega
    Set xarrDetalle = New XArrayDB
    xarrDetalle.ReDim 0, -1, 0, 5
    Set ctlGrilla1.DataSource = obj.ListaConteo(Entrega, Reconteo)
    Me.Caption = "Esta realizando la entrega :" & Entrega
    If Reconteo = "1" Then
     Label2.Caption = "F1-Finalizar ReConteo"
    Else
     Label2.Caption = "F1-Finalizar Conteo"
    End If
    SeteaGrilla
    Me.Show vbModal
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

Private Sub ctlTextBox1_KeyPress(KeyAscii As Integer)
Dim rs As oraDynaset
    If KeyAscii = vbKeyReturn Then
        If Me.ctlGrilla1.ApproxCount > 0 Then
            Set rs = obj.ListaProducto2(strEntrega, Trim(ctlTextBox1.Text), strReconteo)
                If rs.EOF Then
                    MsgBox "El Producto no existe "
                Else
                    Agregar "" & rs("COD_PRODUCTO"), "" & rs("DES_PRODUCTO"), "" & rs("DES_LABORATORIO"), "1"
                    ctlTextBox1.Text = ""
                    ctlTextBox1.SetFocus
                    Me.ctlGrilla1.Limpiar
                    Set Me.ctlGrilla1.DataSource = obj.ListaConteoOrden(strEntrega, strReconteo, "" & rs("COD_PRODUCTO"))
                End If
            'frm_ADM_EntregaProducto.Carga rs
        Else
            MsgBox "No se encontraron items en la grilla.", vbCritical, "Error"
            ctlTextBox1.Text = ""
            ctlTextBox1.SetFocus
        End If
    End If
End Sub

Sub Agregar(ByVal Codigo As String, ByVal Descripcion As String, ByVal Laboratorio As String, ByVal Cantidad As String)
    Dim rs As oraDynaset
    Set rs = obj.AgregaProducto(strEntrega, Codigo, IIf(Cantidad = "", 1, Cantidad), strReconteo)
    Set ctlGrilla1.DataSource = rs
'    ctlGrilla1.Rebind
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "CTD_PRODUCTO")
    arrCaption = Array("Codigo", "Producto", "Laboratorio", "Cantidad")
    arrAncho = Array(800, 3500, 3500, 500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
Dim mensaje As String
Dim numGuiaDev As String
If KeyCode = vbKeyF1 Then
    If strReconteo = 0 Then
        mensaje = "Esta seguro de cerrar el conteo."
    Else
        mensaje = "Esta seguro de cerrar el conteo." + Chr(13) + "También se procederá a cerrar la recepción."
    End If
    If MsgBox(mensaje, vbYesNo, App.ProductName) = vbYes Then
        If strReconteo = 0 Then
            obj.Cierra strEntrega, objUsuario.Codigo, strReconteo, objUsuario.NombrePC, numGuiaDev
        Else
            Dim arr As Variant
            Dim objDocumento As New clsDocumento
            obj.Cierra strEntrega, objUsuario.Codigo, strReconteo, objUsuario.NombrePC, numGuiaDev
            MsgBox "Se ha generado la(s) Guía(s) Nº " + Replace(numGuiaDev, "|", " - ") + "." + Chr(13) + "Verifique su Impresora", vbInformation, "Guías de Devolución"
            arr = Split(numGuiaDev, "|")
            Dim j As Integer
            For j = 0 To UBound(arr)
                objDocumento.ImprimirGuiaTransferencia arr(j)
            Next
        End If
        
        Unload Me
    End If
End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Atención - " & Err.Number
End Sub

