VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_DLV_BuscaTarjetas 
   Caption         =   "Buscador de Tarjetas de Crédito"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12495
   Icon            =   "frm_DLV_BuscaTarjetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Filtros de Búsqueda"
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   12495
      Begin MSComCtl2.DTPicker dtpFchIni 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102825985
         CurrentDate     =   39903
      End
      Begin MSComCtl2.DTPicker dtpFchFin 
         Height          =   315
         Left            =   3240
         TabIndex        =   1
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102825985
         CurrentDate     =   39903
      End
      Begin vbp_Ventas.ctlTextBox txtNroTarjeta 
         Height          =   315
         Left            =   6120
         TabIndex        =   2
         Top             =   315
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         Tipo            =   8
         Alignment       =   2
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
      Begin vbp_Ventas.ctlDataCombo cboTarjetas 
         Height          =   315
         Left            =   9720
         TabIndex        =   3
         Top             =   315
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
         Height          =   195
         Left            =   9000
         TabIndex        =   10
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblNroTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Nro Tarjeta :"
         Height          =   195
         Left            =   5160
         TabIndex        =   9
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblFchFin 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   375
         Width           =   510
      End
      Begin VB.Label lblFchInicio 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   375
         Width           =   555
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1058
      ModoBotones     =   7
   End
   Begin vbp_Ventas.ctlGrilla grdDatos 
      Height          =   6975
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12303
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_DLV_BuscaTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFpago As New clsFormaPago
Dim objDelivery As New clsDelivery

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case Index
        Case 1
            Buscar
        Case 2
            Buscar
        Case 3
            If grdDatos.ApproxCount <= 0 Then MsgBox "No se puede Imprimir sin datos", vbExclamation, App.ProductName: Exit Sub
            grdDatos.MostrarImprimir
        Case 4
            If grdDatos.ApproxCount <= 0 Then MsgBox "No se puede Exportar sin datos", vbExclamation, App.ProductName: Exit Sub
            grdDatos.MostrarExcel
        Case 5
            If grdDatos.ApproxCount <= 0 Then MsgBox "No se puede enviar por Email sin datos", vbExclamation, App.ProductName: Exit Sub
            grdDatos.MostrarEmail
        Case 6
            Unload Me
    End Select
End Sub

Private Sub dtpFchIni_Change()
    dtpFchFin.Value = dtpFchIni.Value
End Sub

Private Sub Form_Load()
    
    Set cboTarjetas.RowSource = objFpago.ListaTarjetasVenta
        cboTarjetas.BoundColumn = "COD"
        cboTarjetas.ListField = "DES"
        cboTarjetas.BoundText = "*"
    
    dtpFchIni.Value = objUsuario.sysdate
    dtpFchFin.Value = objUsuario.sysdate

    SeteaGrd
End Sub

Private Sub SeteaGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

    arrCampos = Array("ESTADO", "NUM_PROFORMA", _
                      "FCH_REGISTRA", "COD_HIJO", _
                      "DES_HIJO", "NUM_TARJETA", _
                      "NUM_AUTORIZACION", "IMP_MONEDA_NAC", _
                      "FCH_VENCIMIENTO", "COD_LOCAL_REF")
                      
    arrCaption = Array("Estado", "Pedido", _
                       "Fecha", "Tip.Tarjeta", _
                       "Tarjeta", "Nro.Tarjeta", _
                       "Autorizacion", "Importe", _
                       "Fch.Vec.", "Local")
                       
    arrAncho = Array(1500, 1000, _
                     2200, 900, _
                     1500, 1800, _
                     1000, 800, 1200, 1000)
                     
    arrAlineacion = Array(dbgLeft, dbgCenter, _
                          dbgCenter, dbgCenter, _
                          dbgLeft, dbgLeft, _
                          dbgCenter, dbgRight, dbgCenter, dbgLeft)
                          
    grdDatos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDatos.Columns(9).Visible = False

End Sub

Private Sub Buscar()

Dim dblInicio As Double
Dim dblFin As Double

    On Error GoTo CtrlErr
        Screen.MousePointer = vbHourglass

        dblInicio = Val(Format(dtpFchIni.Value, "YYYYMMDD"))
        dblFin = Val(Format(dtpFchFin.Value, "YYYYMMDD"))
        
        If dblInicio - dblFin = 0 Or dblInicio - dblFin = -1 Then
        
            If txtNroTarjeta.Text = "" Then MsgBox "Ingrese el Numero de Tarjeta", vbCritical, App.ProductName: txtNroTarjeta.SetFocus: Screen.MousePointer = vbDefault: Exit Sub
        
            Set grdDatos.DataSource = objDelivery.ListaPedido_x_Tarjetas(dtpFchIni.Value, _
                                                                         dtpFchFin.Value, _
                                                                         cboTarjetas.BoundText, _
                                                                         txtNroTarjeta.Text)
            Set objDelivery = Nothing
            Screen.MousePointer = vbDefault
        Else
            MsgBox "La busqueda no pude ser mayor de un día", vbOKOnly + vbExclamation, "Aviso"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdDatos_DblClick()
On Error GoTo CtrlErr

    If grdDatos.ApproxCount = 0 Then Exit Sub
    frm_VTA_DetallePedido.NumeroPedido = grdDatos.Columns("NUM_PROFORMA").Value
    frm_VTA_DetallePedido.CodigoLocal = grdDatos.Columns("COD_LOCAL_REF")
    frm_VTA_DetallePedido.ReCargaDetPedido
    frm_VTA_DetallePedido.Show vbModal
Exit Sub
CtrlErr:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
