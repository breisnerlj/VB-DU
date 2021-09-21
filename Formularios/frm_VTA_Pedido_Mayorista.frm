VERSION 5.00
Begin VB.Form frm_VTA_Pedido_Mayorista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Pedido"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   ControlBox      =   0   'False
   Icon            =   "frm_VTA_Pedido_Mayorista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   11415
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar los Pendientes"
         Height          =   255
         Left            =   9120
         TabIndex        =   1
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin vbp_Ventas.ctlTextBox txtCliente 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   255
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Tipo            =   2
         MaxLength       =   50
         EnabledFoco     =   0   'False
         Bloqueado       =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtPedido 
         Height          =   315
         Left            =   7320
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Tipo            =   3
         Alignment       =   2
         MaxLength       =   10
         EnabledFoco     =   0   'False
         Bloqueado       =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtUsuario 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   630
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Tipo            =   2
         MaxLength       =   50
         EnabledFoco     =   0   'False
         Bloqueado       =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtDireccionEntrega 
         Height          =   315
         Left            =   840
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Tipo            =   2
         MaxLength       =   250
         EnabledFoco     =   0   'False
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtOC 
         Height          =   315
         Left            =   840
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2000
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Tipo            =   2
         MaxLength       =   30
         EnabledFoco     =   0   'False
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtPoliza 
         Height          =   315
         Left            =   840
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Tipo            =   2
         MaxLength       =   15
         EnabledFoco     =   0   'False
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtNombreClienteAux 
         Height          =   315
         Left            =   3840
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1560
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   556
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Tipo            =   2
         MaxLength       =   250
         EnabledFoco     =   0   'False
         Bloqueado       =   -1  'True
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
      Begin VB.Label Label3 
         Caption         =   "Nombre Cliente:"
         Height          =   315
         Left            =   2640
         TabIndex        =   19
         Top             =   1650
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Poliza"
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   1650
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "O/C:"
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   2050
         Width           =   555
      End
      Begin VB.Label lblDireccionEntrega 
         Caption         =   "Dirección Entrega:"
         Height          =   435
         Left            =   120
         TabIndex        =   14
         Top             =   1125
         Width           =   675
      End
      Begin VB.Label lblPedido 
         AutoSize        =   -1  'True
         Caption         =   "Pedido :"
         Height          =   195
         Left            =   6600
         TabIndex        =   11
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   345
         Width           =   555
      End
      Begin VB.Label lblFPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Pedido"
         Height          =   195
         Left            =   6480
         TabIndex        =   9
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lblUsuario 
         Caption         =   "Usuario"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblDtpFchPedido 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000040&
         Height          =   315
         Left            =   7320
         TabIndex        =   7
         Top             =   685
         Width           =   1335
      End
   End
   Begin vbp_Ventas.ctlGrillaArray grdPedido 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5953
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1058
      ModoBotones     =   10
   End
   Begin VB.Label lblCboEstado 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   600
      Left            =   7680
      TabIndex        =   12
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frm_VTA_Pedido_Mayorista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCabecera As oraDynaset
Dim rstPedido As oraDynaset
'Private m_blnCancelar As Boolean
Public ixdbPedido As New XArrayDB

Public Sub Carga(ByVal Pedido As String)
       Set m_objPedido = New clsPedido
       m_objPedido.Pedido = Pedido
       Me.Show vbModal
End Sub

Private Sub Form_Load()
    'm_blnCancelar = False
    ixdbPedido.ReDim 0, -1, 0, 18
    grdPedido.Rebind
    SeteaGrilla
    CargaCabecera
    CargaDetalle
End Sub

Private Sub SeteaGrilla()

Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim i As Integer


    
  
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "")

    arrCaption = Array("Cod.", "Descripción", "Laboratorio", "Linea", _
                       "Pedido", "Atendido", "Stock", "Observación", _
                       "Pedido Calc", "Atendido Calc", "Stock Calc", "Fracciona", "Ctd Fraciona")

    arrAncho = Array(600, 3400, 1000, 1000, 800, 800, 800, 800, 250, 250, _
                     250, 800, 1000)

    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgLeft, _
                          dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgLeft, _
                          dbgRight, dbgRight, dbgRight)

    With grdPedido
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

        .AllowUpdate = False
        .Columns("STOCK").FetchStyle = True
        
         For i = 0 To 3
            .Columns(i).AllowSizing = True
        Next i

        For i = 4 To 6
            .Columns(i).AllowSizing = False
        Next i

       .Columns(7).AllowSizing = True
        
        For i = 8 To 12
            .Columns(i).AllowSizing = False
        Next i
        
        For i = 8 To 10
        .Columns(i).Visible = False
        Next i
        
    End With

End Sub

Private Sub CargaCabecera()
On Error GoTo Control

    Set rstCabecera = m_objPedido.Cabecera_Pedido(objUsuario.CodigoLocal, m_objPedido.Pedido)

    If rstCabecera.RecordCount = 0 Then
        MsgBox "No existe Cabecera en el Pedido Nº " & Trim(rstCabecera("NUM_PEDIDO").Value), vbCritical, "Error"
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    txtCliente.Text = UCase(Trim(rstCabecera("CLIENTE").Value))
    txtUsuario.Text = UCase(Trim(rstCabecera("USUARIO").Value))
    txtPedido.Text = rstCabecera("NUM_PEDIDO").Value
    txtDireccionEntrega.Text = rstCabecera("DIR_ENTREGA_PROD").Value
    lblDtpFchPedido.Caption = rstCabecera("FCH_EMISION").Value
    lblCboEstado.Caption = rstCabecera("EST_PEDIDO").Value
    If Not IsNull(rstCabecera("OBS_PEDIDO01").Value) Then txtOC.Text = rstCabecera("OBS_PEDIDO01").Value
     If Not IsNull(rstCabecera("CLIENTEAUX").Value) Then txtNombreClienteAux.Text = rstCabecera("CLIENTEAUX").Value
    
    If rstCabecera("COD_CLIENTE_AUX").Value <> "" And Not IsNull(rstCabecera("COD_CLIENTE_AUX").Value) Then
        txtPoliza.Enabled = True
        txtNombreClienteAux.Enabled = True
    Else
        txtPoliza.Enabled = False
        txtNombreClienteAux.Enabled = False
    End If
    
    If rstCabecera("ESTADO").Value = "ANU" Then ctlToolBar1.Buttons(10).Enabled = False
    
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub CargaDetalle()
On Error GoTo Control
    
    Set rstPedido = m_objPedido.Detalle_Pedido(objUsuario.CodigoLocal, txtPedido.Text)
    
    While Not rstPedido.EOF
        m_objPedido.AddProducto rstPedido!COD_PRODUCTO, rstPedido!DES_PRODUCTO, rstPedido!DES_LABORATORIO, rstPedido!DES_LINEA, rstPedido!PEDIDO_CALC, rstPedido!ATENDIDO_CALC, rstPedido!Stock, "" & rstPedido!OBS_ITEM, rstPedido!FLG_FRACCION, rstPedido!Ctd_Fraccion, rstPedido!mto_precio_pac
        rstPedido.MoveNext
    Wend
        Dim Y As Integer
        Set ixdbPedido = m_objPedido.ListaProducto
        grdPedido.Array1 = muestraArray(m_objPedido.ListaProducto, 1)
        grdPedido.Rebind
        grdPedido.MoveFirst

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Function muestraArray(Arr As XArrayDB, Optional flagTodos As Integer = 0) As XArrayDB
Dim xTemp As New XArrayDB
xTemp.ReDim 0, -1, 0, 18
Dim u, g As Integer
xTemp.Clear
While u < Arr.Count(1)
    If flagTodos = 1 Then
        If (Arr(u, 8) - Arr(u, 9)) <> 0 Then
            xTemp.AppendRows
              g = xTemp.Count(1) - 1
             xTemp(g, 0) = Arr(u, 0)
             xTemp(g, 1) = Arr(u, 1)
             xTemp(g, 2) = Arr(u, 2)
             xTemp(g, 3) = Arr(u, 3)
             xTemp(g, 4) = Arr(u, 4)
             xTemp(g, 5) = Arr(u, 5)
             xTemp(g, 6) = Arr(u, 6)
             xTemp(g, 7) = Arr(u, 7)
             xTemp(g, 8) = Arr(u, 8)
             xTemp(g, 9) = Arr(u, 9)
             xTemp(g, 10) = Arr(u, 10)
             If Arr(u, 11) = 0 Then
                xTemp(g, 11) = "No"
            Else
                xTemp(g, 11) = "Si"
             End If
             xTemp(g, 12) = Arr(u, 12)
        End If
    Else
        xTemp.AppendRows
          g = xTemp.Count(1) - 1
         xTemp(g, 0) = Arr(u, 0)
         xTemp(g, 1) = Arr(u, 1)
         xTemp(g, 2) = Arr(u, 2)
         xTemp(g, 3) = Arr(u, 3)
         xTemp(g, 4) = Arr(u, 4)
         xTemp(g, 5) = Arr(u, 5)
         xTemp(g, 6) = Arr(u, 6)
         xTemp(g, 7) = Arr(u, 7)
         xTemp(g, 8) = Arr(u, 8)
         xTemp(g, 9) = Arr(u, 9)
         xTemp(g, 10) = Arr(u, 10)
         If Arr(u, 11) = 0 Then
            xTemp(g, 11) = "No"
        Else
            xTemp(g, 11) = "Si"
         End If
         xTemp(g, 12) = Arr(u, 12)
    End If
    u = u + 1
Wend
Set muestraArray = xTemp
End Function

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
   On Error GoTo Control

    Select Case Index
        Case 1:
            'If grdPedido.ApproxCount > 0 Then
                    If Validar Then
                        Grabar
                    End If
'            Else
'                MsgBox "El Pedido No Tiene Items Pendientes Por Atender", vbOKOnly + vbExclamation, "Validación"
'            End If
        Case 2:
        If lblCboEstado.Caption <> "ATENCION TOTAL" Then
                If MsgBox("Estas seguro que Deseas Cancelar, se perderán los Despachos Realizados ", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                    'm_blnCancelar = True
                    intGrabadoPedido = 0
                    Set m_objPedido = Nothing
                    Unload Me
                End If
        Else
            intGrabadoPedido = 0
            Set m_objPedido = Nothing
            Unload Me
        End If
        
        Case Else
            MsgBox "No se encuentra implementado", vbCritical, App.ProductName
    End Select
   Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Function Validar() As Boolean
Validar = True

If txtDireccionEntrega.Text = "" Then
   Validar = False
   MsgBox "Falta Ingresar la Dirección de Entrega", vbOKOnly + vbExclamation, "Validación"
   txtDireccionEntrega.Focus
   GoTo Termina:
End If

If rstCabecera("COD_CLIENTE_AUX").Value <> "" And Not IsNull(rstCabecera("COD_CLIENTE_AUX").Value) Then
    If Trim(txtPoliza.Text) = "" Or Trim(txtNombreClienteAux.Text) = "" Then
       Validar = False
       MsgBox "Falta Ingresar Datos del Botiquín", vbOKOnly + vbExclamation, "Validación"
       txtPoliza.Focus
       GoTo Termina:
    End If
End If


Termina:
End Function

Private Sub Grabar()
Dim larrProducto() As String
Dim larrUnidadAte() As String
Dim larrFraccionAte() As String
Dim larrNumLote() As String
Dim larrFchVenc() As String
Dim i As Integer
Dim strMensaje As String, strNumDocumento As String, strImpToltal As Double

On Error GoTo Control

    ReDim larrProducto(0 To 0)
    ReDim larrUnidadAte(0 To 0)
    ReDim larrFraccionAte(0 To 0)
    ReDim larrNumLote(0 To 0)
    ReDim larrFchVenc(0 To 0)

    If ixdbPedido.LowerBound(1) <> "-1" Then

       For i = ixdbPedido.LowerBound(1) To ixdbPedido.UpperBound(1)
           larrProducto(UBound(larrProducto)) = ixdbPedido(i, 0)
           larrUnidadAte(UBound(larrUnidadAte)) = pfstr_Segmento(IIf(IsNull(ixdbPedido(i, 13)), "0", ixdbPedido(i, 13)), False, "|")
           larrFraccionAte(UBound(larrFraccionAte)) = pfstr_Segmento(IIf(IsNull(ixdbPedido(i, 14)), "0", ixdbPedido(i, 14)), False, "|")
           larrNumLote(UBound(larrNumLote)) = pfstr_Segmento(ixdbPedido(i, 15), False, "|")
           larrFchVenc(UBound(larrFchVenc)) = pfstr_Segmento(ixdbPedido(i, 16), False, "|")

           ReDim Preserve larrProducto(UBound(larrProducto) + 1)
           ReDim Preserve larrUnidadAte(UBound(larrUnidadAte) + 1)
           ReDim Preserve larrFraccionAte(UBound(larrFraccionAte) + 1)
           ReDim Preserve larrNumLote(UBound(larrNumLote) + 1)
           ReDim Preserve larrFchVenc(UBound(larrFchVenc) + 1)
       Next i

       ReDim Preserve larrProducto(UBound(larrProducto) - 1)
       ReDim Preserve larrUnidadAte(UBound(larrUnidadAte) - 1)
       ReDim Preserve larrFraccionAte(UBound(larrFraccionAte) - 1)
       ReDim Preserve larrNumLote(UBound(larrNumLote) - 1)
       ReDim Preserve larrFchVenc(UBound(larrFchVenc) - 1)

    End If

    If MsgBox("¿Seguro(a) de Guardar el Pedido Mayorista?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        Exit Sub
    End If

    strMensaje = m_objPedido.GrabaPedido(strNumDocumento, strImpToltal, _
                                         objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, _
                                         objUsuario.NombrePC, objUsuario.Codigo, _
                                         txtPedido.Text, objUsuario.CodigoLiquidacion, _
                                         larrProducto, larrUnidadAte, larrFraccionAte, larrNumLote, larrFchVenc, Trim(txtDireccionEntrega.Text), Trim(txtOC.Text), Trim(txtPoliza.Text), Trim(txtNombreClienteAux.Text))

    If strMensaje = "" Then
       intGrabadoPedido = 1
       MsgBox "Se grabo satisfactoriamente la Guia N°" & strNumDocumento, vbExclamation, App.ProductName
       Unload Me
    Else
        intGrabadoPedido = 0
        MsgBox strMensaje, vbCritical, App.ProductName
    End If

   Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub Check1_Click()
    grdPedido.Array1 = muestraArray(ixdbPedido, Check1.Value)
    grdPedido.Rebind
    grdPedido.MoveFirst
End Sub

Private Sub grdPedido_DblClick()
'    If Not (Val(grdPedido.Columns(8)) - Val(grdPedido.Columns(9))) <> 0 Then
'    Exit Sub
'    End If
    If grdPedido.ApproxCount > 0 Then
        If rstCabecera("ESTADO").Value = "ANU" Or rstCabecera("ESTADO").Value = "EMI" Or rstCabecera("ESTADO").Value = "INA" Then
            Exit Sub
        End If

        frm_VTA_Pedido_Mayorita_LoteFchVenc.CargaDatos grdPedido.Columns(0)
        grdPedido.Array1 = muestraArray(m_objPedido.ListaProducto, Check1.Value)
        grdPedido.Rebind
    End If
    
End Sub

Private Sub grdPedido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdPedido_DblClick
    End If
End Sub

Private Sub grdPedido_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
On Error GoTo Control

    Select Case col
        Case grdPedido.Columns(5).ColIndex
             Select Case grdPedido.Columns(5).CellValue(Bookmark)
                 Case 0
                      CellStyle.BackColor = RGB(175, 50, 50)
                      CellStyle.ForeColor = vbWhite
             End Select
    End Select

   Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

'Public Property Get Cancelar() As Boolean
'    Cancelar = m_blnCancelar
'End Property
'
'Public Property Get Pedido() As clsPedido
'    Set Pedido = m_objPedido
'End Property

Private Sub Form_Unload(Cancel As Integer)
    Set m_objPedido = Nothing
End Sub

