VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_Pedido_Mayorita_LoteFchVenc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cantidades Lotes y Fecha Vencimiento"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "frm_VTA_Pedido_Mayorita_LoteFchVenc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrillaArray grdLoteFchVen 
      Height          =   2895
      Left            =   0
      TabIndex        =   22
      Top             =   2760
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5106
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1058
      ModoBotones     =   10
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   495
      Left            =   6760
      Picture         =   "frm_VTA_Pedido_Mayorita_LoteFchVenc.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   495
      Left            =   6760
      Picture         =   "frm_VTA_Pedido_Mayorita_LoteFchVenc.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   6615
      Begin vbp_Ventas.ctlTextBox txtLote 
         Height          =   345
         Left            =   3360
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   20
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
      Begin vbp_Ventas.ctlTextBox txtFracciones 
         Height          =   345
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   10
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
      Begin vbp_Ventas.ctlTextBox txtUnidades 
         Height          =   345
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   10
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
      Begin MSMask.MaskEdBox txtFechaVencimiento 
         Height          =   345
         Left            =   4920
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFchVecmiento 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Vencimiento"
         Height          =   195
         Left            =   4935
         TabIndex        =   12
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label lblLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblUnidades 
         AutoSize        =   -1  'True
         Caption         =   "Unidades"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblFracciones 
         AutoSize        =   -1  'True
         Caption         =   "Fracciones"
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Pendiente :"
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
      Left            =   3480
      TabIndex        =   20
      Top             =   1350
      Width           =   990
   End
   Begin VB.Label lblPendiente 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   4560
      TabIndex        =   19
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label lblPedido 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   2400
      TabIndex        =   18
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pedido :"
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
      Left            =   1800
      TabIndex        =   17
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblstock 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   720
      TabIndex        =   16
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Stock :"
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
      Left            =   120
      TabIndex        =   15
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblCodigo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Código :"
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
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
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
      Left            =   1320
      TabIndex        =   13
      Top             =   600
      Width           =   1140
   End
End
Attribute VB_Name = "frm_VTA_Pedido_Mayorita_LoteFchVenc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim indice As Integer
Dim xTemp As New XArrayDB

Public Sub CargaDatos(ByVal Codigo As String)
On Error GoTo handle
    indice = m_objPedido.ListaProducto.Find(0, 0, Trim(Codigo), XORDER_ASCEND)
    lblCodigo.Caption = m_objPedido.ListaProducto(indice, 0)
    lblDescripcion.Caption = m_objPedido.ListaProducto(indice, 1)
    lblstock.Caption = m_objPedido.ListaProducto(indice, 6)
    lblPedido.Caption = m_objPedido.ListaProducto(indice, 4)
    lblPendiente.Caption = m_objPedido.fnConvierte(m_objPedido.ListaProducto(indice, 11), IIf(m_objPedido.ListaProducto(indice, 11) = "1", 0, m_objPedido.ListaProducto(indice, 8) - m_objPedido.ListaProducto(indice, 9)), IIf(m_objPedido.ListaProducto(indice, 11) = "1", m_objPedido.ListaProducto(indice, 8) - m_objPedido.ListaProducto(indice, 9), 0), m_objPedido.ListaProducto(indice, 12))
    If m_objPedido.ListaProducto(indice, 11) = "0" Then
        txtFracciones.Enabled = False
    End If
        
    CargaPedido True
    SeteaGrilla
    grdLoteFchVen.Array1 = xTemp
    CargarPorDefecto
    Me.Show vbModal
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub CargaPedido(Optional PrimeraVez As Boolean = False)
    
    Dim g As Integer
    xTemp.ReDim 0, -1, 0, 5
    Dim arrUnidades As Variant
    Dim arrFracciones As Variant
    Dim arrLote As Variant
    Dim arrFechaVencimiento As Variant
    If PrimeraVez = True Then
        arrUnidades = Split(m_objPedido.ListaProducto(indice, 17), "|")
        arrFracciones = Split(m_objPedido.ListaProducto(indice, 18), "|")
        arrLote = Split(m_objPedido.ListaProducto(indice, 19), "|")
        arrFechaVencimiento = Split(m_objPedido.ListaProducto(indice, 20), "|")
    Else
        arrUnidades = Split(m_objPedido.ListaProducto(indice, 13), "|")
        arrFracciones = Split(m_objPedido.ListaProducto(indice, 14), "|")
        arrLote = Split(m_objPedido.ListaProducto(indice, 15), "|")
        arrFechaVencimiento = Split(m_objPedido.ListaProducto(indice, 16), "|")
    End If
    While g < UBound(arrUnidades)
        xTemp.AppendRows
        xTemp(g, 0) = m_objPedido.ListaProducto(indice, 0)
        xTemp(g, 1) = arrUnidades(g)
        xTemp(g, 2) = arrFracciones(g)
        xTemp(g, 3) = arrLote(g)
        xTemp(g, 4) = arrFechaVencimiento(g)
        g = g + 1
    Wend
    grdLoteFchVen.Rebind
    lblstock.Caption = m_objPedido.ListaProducto(indice, 6)
    lblPedido.Caption = m_objPedido.ListaProducto(indice, 4)
    lblPendiente.Caption = m_objPedido.fnConvierte(m_objPedido.ListaProducto(indice, 11), IIf(m_objPedido.ListaProducto(indice, 11) = "1", 0, m_objPedido.ListaProducto(indice, 8) - m_objPedido.ListaProducto(indice, 9)), IIf(m_objPedido.ListaProducto(indice, 11) = "1", m_objPedido.ListaProducto(indice, 8) - m_objPedido.ListaProducto(indice, 9), 0), m_objPedido.ListaProducto(indice, 12))
    
End Sub

Private Sub SeteaGrilla()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim i As Integer

    arrCampos = Array("", "", "", "", "")
    arrCaption = Array("Producto", "Unidades", "Fracciones", "Lote", "Fch. Venc.")
    arrAncho = Array(1200, 1200, 1200, 1200, 1200)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft)

    grdLoteFchVen.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    With grdLoteFchVen
        .AllowUpdate = False
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

        For i = 0 To 4
            .Columns(i).AllowSizing = False
        Next i
    End With

End Sub

Private Sub CargarPorDefecto()
'txtUnidades.Text =
If m_objPedido.ListaProducto(indice, 10) <= m_objPedido.ListaProducto(indice, 8) - m_objPedido.ListaProducto(indice, 9) Then
        txtUnidades.Text = Val(fncGuion(m_objPedido.ListaProducto(indice, 6), 0, "F"))
        txtFracciones.Text = Val(IIf(InStr(1, m_objPedido.ListaProducto(indice, 6), "F") = 0, 0, fncGuion(m_objPedido.ListaProducto(indice, 6), 1, "F")))
    Else
        txtUnidades.Text = Val(fncGuion(lblPendiente.Caption, 0, "F"))
        txtFracciones.Text = Val(IIf(InStr(1, lblPendiente.Caption, "F") = 0, 0, fncGuion(lblPendiente.Caption, 1, "F")))
End If

End Sub

Private Sub cmdAgregar_Click()
If Verifica = True Then
    m_objPedido.addAtendido m_objPedido.ListaProducto(indice, 0), Val(txtUnidades.Text), Val(txtFracciones.Text), txtLote.Text, txtFechaVencimiento.Text, False
    CargaPedido
    CargarPorDefecto
    txtLote.Clear
    txtFechaVencimiento.Text = "__/__/____"
    txtUnidades.Focus
End If
End Sub

Function Verifica() As Boolean
 Dim arrVerLotes As Variant
 Dim arrVerFechaVencimiento As Variant
 Dim i As Integer
 Dim dfecha As Date

    Verifica = True
    If (Trim(txtLote.Text) <> "" And txtFechaVencimiento.Text = "__/__/____") Or (Trim(txtLote.Text) = "" And txtFechaVencimiento.Text <> "__/__/____") Then
      Verifica = False
      MsgBox "Faltan Ingresar Datos del Lote o la Fecha de Vencimiento", vbOKOnly + vbExclamation, "Validación"
      GoTo Termina
    End If
    
    
    If Not (Trim(txtLote.Text) = "" And txtFechaVencimiento.Text = "__/__/____") Then
    
        If Mid(txtFechaVencimiento.Text, 4, 2) > 12 Then
           Verifica = False
           MsgBox "La fecha de vencimiento es incorrecta", vbCritical + vbOKOnly, App.ProductName
           GoTo Termina
        End If
    
       If InStr(txtFechaVencimiento.Text, "_") <> 0 Then
          Verifica = False
          MsgBox "La fecha de vencimiento es incorrecta", vbCritical + vbOKOnly, App.ProductName
          GoTo Termina
       End If
       
       dfecha = (Day(objUsuario.sysdate) & "/" & Month(objUsuario.sysdate) & "/" & Year(objUsuario.sysdate))
   
       If Not IsDate(txtFechaVencimiento.Text) Then
          Verifica = False
          MsgBox "La fecha de vencimiento es incorrecta", vbCritical + vbOKOnly, App.ProductName
          GoTo Termina
       End If
       
       
       
       If CDate(txtFechaVencimiento.Text) < dfecha Then
          Verifica = False
          MsgBox "La fecha de vencimiento es incorrecta", vbCritical + vbOKOnly, App.ProductName
          GoTo Termina
       End If
    End If
    
    
    
    
    
   arrVerLotes = Split(m_objPedido.ListaProducto(indice, 15), "|")
   arrVerFechaVencimiento = Split(m_objPedido.ListaProducto(indice, 16), "|")
   
   If UBound(arrVerLotes) >= 0 Then
      For i = LBound(arrVerLotes) To UBound(arrVerLotes) - 1
        If arrVerLotes(i) = Trim(txtLote.Text) Then
            If arrVerFechaVencimiento(i) = txtFechaVencimiento.Text Then
                 Verifica = False
                 MsgBox "Ya Existe un Lote con Dicha Fecha de Vencimiento", vbOKOnly + vbExclamation, "Validación"
                 txtLote.Focus
                 GoTo Termina
            End If
        End If
      Next
   End If
   
   
    
Termina:
End Function

Private Sub cmdEliminar_Click()
If grdLoteFchVen.ApproxCount Then
    m_objPedido.addAtendido m_objPedido.ListaProducto(indice, 0), grdLoteFchVen.Columns(1), grdLoteFchVen.Columns(2), 0, 0, True, grdLoteFchVen.Bookmark
    grdLoteFchVen.Delete
    CargaPedido
    CargarPorDefecto
End If
End Sub

Private Sub grdLoteFchVen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cmdEliminar_Click
End If
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
   On Error GoTo Control

    Select Case Index
        Case 1:
            m_objPedido.GrabaDetalle indice
            Unload Me
        Case 2:
            m_objPedido.CancelaDetalle indice
            Unload Me
        Case Else
            MsgBox "No se encuentra implementado", vbCritical, App.ProductName
    End Select
   Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub


Private Sub txtFechaVencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Call cmdAgregar_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
ctlToolBar1_Click Cancelar, 2
End Sub
