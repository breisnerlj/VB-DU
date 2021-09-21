VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_Pedido_Mayorista 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrilla grdPedidos 
      Height          =   4455
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7858
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   7095
      Begin vbp_Ventas.ctlDataCombo cboEstado 
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   735
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin MSComCtl2.DTPicker dtpFchInicio 
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   40004
      End
      Begin MSComCtl2.DTPicker dtpFchFin 
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   40004
      End
      Begin vbp_Ventas.ctlTextBox txtCliente 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Tipo            =   2
         MaxLength       =   50
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
      Begin vbp_Ventas.ctlTextBox txtPedido 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Tipo            =   3
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
      Begin VB.Label lblPedido 
         AutoSize        =   -1  'True
         Caption         =   "Pedido :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   825
         Width           =   585
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1305
         Width           =   555
      End
      Begin VB.Label lblFFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Fin"
         Height          =   195
         Left            =   3240
         TabIndex        =   8
         Top             =   330
         Width           =   390
      End
      Begin VB.Label lblFInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   555
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado :"
         Height          =   195
         Left            =   3240
         TabIndex        =   6
         Top             =   810
         Width           =   585
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_Pedido_Mayorista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPedido As New clsPedido

Private Sub Form_Load()
On Error GoTo Control

    setteaFormulario Me
    SeteaGrilla
    Inicio
    BuscaValores

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub SeteaGrilla()

Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim i As Integer

    arrCampos = Array("NUM_PEDIDO", "CLIENTE", "FCH_EMISION", "ITEMS", "EST_PEDIDO", "ATENDIDO", "PENDIENTE", "TOTAL")
    arrCaption = Array("#Pedido", "Cliente", "Fecha", "#Items", "Estado", "Ate.", "Pend.", "Total")
    arrAncho = Array(1100, 2000, 1000, 600, 600, 700, 700, 700)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgRight, dbgRight, dbgRight)
    
    With grdPedidos
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        .Columns(3).Visible = False
        .Columns("EST_PEDIDO").FetchStyle = True
        
        For i = 0 To 7
            .Columns(i).AllowSizing = False
        Next i
    End With

End Sub

Private Sub Inicio()
    ctlToolBar1.Buttons(1).Caption = "Atender"
    ctlToolBar1.Buttons(1).ToolTipText = "Atender pedido"
    ctlToolBar1.Buttons(2).Caption = "Facturar"
    ctlToolBar1.Buttons(2).ToolTipText = "Facturar guias generadas"
    dtpFchInicio.Value = objUsuario.sysdate - 7
    dtpFchFin.Value = objUsuario.sysdate

    Set cboEstado.RowSource = objPedido.ListaEstado("", "Todos")
        cboEstado.BoundColumn = "COD"
        cboEstado.ListField = "DES"
        cboEstado.BoundText = "CON"
End Sub

Private Sub BuscaValores()
Dim dblInicio As Double
Dim dblFin As Double

On Error GoTo Control
    
    dblInicio = Val(Format(dtpFchInicio.Value, "YYYYMMDD"))
    dblFin = Val(Format(dtpFchFin.Value, "YYYYMMDD"))
    
    If dblInicio > dblFin Then
        MsgBox "La Fecha inicial no puede ser mayor a la Fecha fin", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If

    'grdPedidos.Limpiar
    Set grdPedidos.DataSource = objPedido.Lista(objUsuario.CodigoLocal, txtPedido.Text, dtpFchInicio.Value, _
                                                dtpFchFin.Value, cboEstado.BoundText, txtCliente.Text)
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
   On Error GoTo Control

    Select Case Index
        Case 1:
           If grdPedidos.ApproxCount > 0 Then
                frm_VTA_Pedido_Mayorista.Carga grdPedidos.Columns(0).Value ' .Show vbModal
                If intGrabadoPedido = 1 Then
                    BuscaValores
                End If
            Else
              MsgBox "No se puede atender un pedido que no esta seleccionado", vbOKOnly + vbExclamation, "Aviso"
           End If
        Case 2:
          If grdPedidos.ApproxCount > 0 Then
                frm_VTA_GuiasxFacturar.Carga grdPedidos.Columns(0).Value, "TRA" ' .Show vbModal
          Else
                 MsgBox "No se puede facturar un pedido que no esta seleccionado", vbOKOnly + vbExclamation, "Aviso"
          End If
        Case 3:
            BuscaValores

        Case 4:
            BuscaValores

        Case 5:
            grdPedidos.MostrarImprimir
            
        Case 6:
            grdPedidos.MostrarExcel
        
        Case 7:
            grdPedidos.MostrarEmail
            
        Case 8:
            Unload Me

        Case Else
            MsgBox "No se encuentra implementado", vbCritical, App.ProductName
    End Select
   Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number

End Sub

Private Sub grdPedidos_DblClick()
    If grdPedidos.ApproxCount <= 0 Then Exit Sub
    Call ctlToolBar1_Click(Buscar, 1)
End Sub

Private Sub grdPedidos_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
On Error GoTo Control
    
    Select Case col
        Case grdPedidos.Columns("EST_PEDIDO").ColIndex
            Select Case grdPedidos.Columns("EST_PEDIDO").CellValue(Bookmark)
                Case "ATE"
                    CellStyle.BackColor = RGB(50, 175, 50)
                    CellStyle.ForeColor = vbWhite
                Case "EMI"
                    CellStyle.BackColor = RGB(50, 50, 175)
                    CellStyle.ForeColor = vbWhite
                Case "ANU"
                    CellStyle.BackColor = RGB(175, 50, 50)
                    CellStyle.ForeColor = vbWhite
            End Select
    End Select

   Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdPedidos_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    If grdPedidos.ApproxCount <= 0 Then Exit Sub
    Call ctlToolBar1_Click(Buscar, 1)
End If
End Sub

Private Sub grdPedidos_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdPedidos.ApproxCount <= 0 Then Exit Sub
    'txtCliente.Text = "" & grdPedidos.Columns("CLIENTE").Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPedido = Nothing
End Sub
