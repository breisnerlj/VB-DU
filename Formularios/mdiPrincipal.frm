VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Modulo de ventas"
   ClientHeight    =   8595
   ClientLeft      =   2610
   ClientTop       =   2550
   ClientWidth     =   11805
   Icon            =   "mdiPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picComandos 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   11805
      TabIndex        =   2
      Top             =   7665
      Width           =   11805
      Begin VB.CommandButton cmdPuntos 
         Caption         =   "Puntos"
         Enabled         =   0   'False
         Height          =   615
         Left            =   9000
         Picture         =   "mdiPrincipal.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   60
         Width           =   825
      End
      Begin VB.TextBox txtToolTip 
         Height          =   285
         Left            =   11520
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   660
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdImpresion 
         Caption         =   "Imp.Pendien"
         Height          =   615
         Left            =   9960
         Picture         =   "mdiPrincipal.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   60
         Width           =   1060
      End
      Begin VB.CommandButton cmdProforma 
         Caption         =   "Proforma"
         Height          =   615
         Left            =   8160
         Picture         =   "mdiPrincipal.frx":0E1E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   825
      End
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "&B?squeda"
         Height          =   615
         Left            =   2030
         Picture         =   "mdiPrincipal.frx":13A8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdModalidad 
         Caption         =   "&Modalidad"
         Height          =   615
         Left            =   0
         Picture         =   "mdiPrincipal.frx":30A2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton cmdMantenimientos 
         Caption         =   "&Mantenimient"
         Height          =   615
         Left            =   960
         Picture         =   "mdiPrincipal.frx":362C
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "002"
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdDocumento 
         Caption         =   "&Documento"
         Height          =   615
         Left            =   5040
         Picture         =   "mdiPrincipal.frx":3BB6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton cmdFormaPago 
         Caption         =   "&Forma Pago"
         Height          =   615
         Left            =   6000
         Picture         =   "mdiPrincipal.frx":4140
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton cmdAdministrador 
         Caption         =   "Administrador"
         Height          =   615
         Left            =   2900
         Picture         =   "mdiPrincipal.frx":46CA
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "001"
         Top             =   60
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   10995
         Picture         =   "mdiPrincipal.frx":4C54
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   825
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   4080
         Picture         =   "mdiPrincipal.frx":51DE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   930
      End
      Begin VB.CommandButton cmdGrabaVenta 
         Caption         =   "&Graba Venta"
         Height          =   615
         Left            =   7080
         Picture         =   "mdiPrincipal.frx":5768
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   10080
         TabIndex        =   30
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   8520
         TabIndex        =   24
         Top             =   660
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   21
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   7560
         TabIndex        =   20
         Top             =   660
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   4200
         TabIndex        =   19
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   18
         Top             =   660
         Width           =   690
      End
      Begin VB.Label lblModalidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + Q"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   10940
         TabIndex        =   6
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   3120
         TabIndex        =   5
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   6360
         TabIndex        =   4
         Top             =   660
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   5400
         TabIndex        =   3
         Top             =   660
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6480
      Left            =   7125
      ScaleHeight     =   6480
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   1185
      Width           =   4680
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   840
         Top             =   4440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   391
         ImageHeight     =   141
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiPrincipal.frx":5CF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiPrincipal.frx":2E4FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiPrincipal.frx":3CBAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiPrincipal.frx":4A087
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiPrincipal.frx":56503
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgLogoBtl 
         Height          =   2115
         Left            =   360
         Top             =   1920
         Visible         =   0   'False
         Width           =   5865
      End
   End
   Begin VB.PictureBox pctDelivery 
      Align           =   1  'Align Top
      BackColor       =   &H00CFF6FC&
      Height          =   1185
      Left            =   0
      ScaleHeight     =   1125
      ScaleWidth      =   11745
      TabIndex        =   13
      Top             =   0
      Width           =   11805
      Begin VB.Frame Frame1 
         BackColor       =   &H00CFF6FC&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1095
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   15200
         Begin vbp_Ventas.ctlTextBox txtDireccion 
            Height          =   615
            Left            =   11160
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1085
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
         Begin VB.CommandButton cmdHistorial 
            Caption         =   "Historial  [ Cliente / Pedido ]"
            Height          =   375
            Left            =   11160
            TabIndex        =   28
            Top             =   120
            Width           =   2175
         End
         Begin vbp_Ventas.ctlTextBox txtDLVTelefono 
            Height          =   375
            Left            =   75
            TabIndex        =   0
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
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
         Begin VB.CommandButton cmdDLVBuscar 
            Caption         =   "&Buscar"
            Height          =   195
            Left            =   1560
            TabIndex        =   25
            Top             =   600
            Width           =   75
         End
         Begin vbp_Ventas.ctlCliente ctlCliente1 
            Height          =   960
            Left            =   1920
            TabIndex        =   32
            Top             =   120
            Width           =   9255
            _extentx        =   16325
            _extenty        =   1693
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel?fono[Alt+W]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   240
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1695
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00CFF6FC&
            Height          =   735
            Left            =   45
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "mdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private objDocumento As New clsDocumento
Private objProforma As New clsProforma
Private objConvenio As New clsConvenio
Private objMaquina As New clsMaquina
Private objLocal As New clsLocal



Public strcodConv As String
Private m_bInLable As Boolean
Private TT1 As New CBalloonToolTip                             '//Demo for On Demand tooltip
Dim objMoneda As New clsMoneda
'Public TipoDoc As String
'Public xNumProforma As String
'Public xCodLocal As String
Public WithEvents wskPrincipal As CSocketMaster
Attribute wskPrincipal.VB_VarHelpID = -1
Public strNombreFomularioOrigen As String

'** Variable para el caso de Ticketera **'
 Public pdblPagTkt As Double
 Public pdblTPagTkt As Double
 
Public Sub cmdAdministrador_Click()
On Error GoTo handle
 '   frm_VTA_ConsultaDoc.Show
    frm_VTA_Administrador.Padre = cmdAdministrador.Tag
    frm_VTA_Administrador.Show vbModal
    
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdBusqueda_Click()

'''''''''''If objVenta.ptmModalidad = Venta_Convenio Then frm_VTA_Convenio.Limpiar
'''''''''''            objVenta.ptmModalidad = Venta_Regular
'''''''''''            ptmTipoPrecio = Regular
'''''''''''            objVenta.CodigoTipoVenta = Venta_Regular
'''''''''''            objVenta.PctBeneficiario = 0
'''''''''''            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
'''''''''''            frmPedido.Label6.Visible = True
'''''''''''            frmPedido.lblTotal.Visible = True
'''''''''''            frmPedido.Label4.Visible = False
'''''''''''            frmPedido.lblPctCopago.Visible = False
'''''''''''            frmPedido.Label8.Visible = False
'''''''''''            frmPedido.lblcopago.Visible = False
On Error GoTo handle
            
            '** Se cambio el tipo de precio para el delivery a "003" **'
            '** 22/01/2008 Por Cristhian Rueda **'
            'validar si objCliente tiene data, de ser el caso mostrar frm busqueda
            If objVenta.CodigoCliente <> "" Then
                If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = 0 Then
                    frm_VTA_Busqueda.Datos Format("3", "000")
                  Else
                    frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
                End If
                'I.ECASTILLO 12.08.2021
                If gstrIndDCSAP = "1" Then
                    If objVenta.isDCSAP = "1" Then
                        MsgBox "Ten en cuenta que para esta direcci?n:" & vbNewLine & _
                               "? El Stock de fracciones es" & vbNewLine & _
                               "  independiente." & vbNewLine & _
                               "? No se puede convertir las cajas" & vbNewLine & _
                               "  en fracciones.", vbExclamation + vbOKOnly, App.ProductName
                    End If
                    If objVenta.busquedaNVeces > 0 And Len(Trim(objVenta.strBusqueda)) > 0 Then
                        frm_VTA_Busqueda.cmdBuscar_Click
                    End If
                End If
                'F.ECASTILLO 12.08.2021
                frm_VTA_Busqueda.SetFocus
                frm_VTA_Busqueda.txtBuscar.SetFocus
            Else
                frm_VTA_Busqueda.cmd_salir_Click
            End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

'    objVenta.ptmModalidad = Venta_Regular
'    ptmTipoPrecio = Regular
'    On Error GoTo handle
'    objVenta.CodigoTipoVenta = Venta_Regular
''    frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
    'frm_VTA_Busqueda.Datos Format(objVenta.CodigoTipoVenta, "000")
    'frm_VTA_Busqueda.Show
    'frm_VTA_Busqueda.SetFocus
    'frm_VTA_RecetarioM.pstrFlgRM = ""
'    Exit Sub
'handle:
'    MsgBox "Error cargar es posible que algunas funcionalidades no se encuentren validad", vbExclamation, App.ProductName
End Sub

Public Sub cmdCancelar_Click()
On Error GoTo handle
   If MsgBox("? Desea borrar todos los Datos del Documento.. ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        Cancelar
   End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
   
End Sub

Private Sub cmdDelivery_Click()

End Sub

Private Sub cmdDLVBuscar_Click()
'''    Dim rsTelefono As oraDynaset
'''    Dim objCliente As New clsCliente

    On Error GoTo handle
    
    If objVenta.ParametroValor("PERCAMTLF") = 1 Then
        If frmPedido.grdPedido.ApproxCount > 0 Then MsgBox "Si desea cambiar el telefono debe cancelar la venta.", vbCritical, "Aviso": Exit Sub
    End If
    
    txtDireccion.Text = ""
    frm_DLV_BuscaTelefono.Telefono = txtDLVTelefono.Text
    
    frm_DLV_BuscaTelefono.Show vbModal
    txtDireccion.Text = objVenta.DireccionClienteDLV
    '''txtDireccion.Text = objVenta.DesAuxCliDirecc
    '''txtDLVTelefono.Text = ctlCliente1.Telefono
    
    frm_VTA_Documento.pblnEditCli = False
    
    If objVenta.UbigeoEntrega = "" Then
        objVenta.UbigeoEntrega = ctlCliente1.Ubigeo
    End If
    
''    If ctlCliente1.Codigo = "" Then
''    Else
''    End If
    
    cmdBusqueda_Click
    
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub cmdDocumento_Click()

'fnImprimeCupon "0000000130", "0000000309"
On Error GoTo handle
    If objVenta.CodigoTipoVenta = Guias_Remision Then
        frm_VTA_GuiaRemision.Show
        frm_VTA_GuiaRemision.SetFocus
    Else
        frm_VTA_Documento.Show
        frm_VTA_Documento.SetFocus
    End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Public Sub cmdFormaPago_Click()
On Error GoTo handle

frmPedido.flgF6 = 1

    frm_VTA_FormaPago.Show
    frm_VTA_FormaPago.SetFocus
'If (frmPedido.optCredito.Visible = False And frmPedido.optEfectivo.Visible = False And frmPedido.optNinguna.Visible = False) Then
'    frm_VTA_FormaPago.Show
'    frm_VTA_FormaPago.SetFocus
'ElseIf frmPedido.optEfectivo.Value = True Then
'    frm_VTA_FormaPagoEfectivo.pstrDato = "001"
'    frm_VTA_FormaPagoEfectivo.pstrDatoDes = "EFECTIVO"
'    frm_VTA_FormaPagoEfectivo.Show
'    frm_VTA_FormaPagoEfectivo.grdEfectivo.SetFocus
'ElseIf frmPedido.optCredito.Value = True Then
'    frm_VTA_FormaPagoTarjeta.pstrDato = "002"
'    frm_VTA_FormaPagoTarjeta.pstrDatoDes = "TARJETA"
'    frm_VTA_FormaPagoTarjeta.Show
'    frm_VTA_FormaPagoTarjeta.SetFocus
'End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Sub grabaProforma()

    Dim dblDlvEmpAsume As Double
    Dim msgAsignaOrderDigProf As String
    'Dim flgReservaCapacidad As String
    objVenta.LocalAtencion = ctlCliente1.LocalAsignado
    objVenta.LocalDespacho = ctlCliente1.LocalDespacho
    'I.ECASTILLO 17.12.2020
    'flgReservaCapacidad = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVCAP") '1 => ACTIVO, 0 => INACTIVO
    
    Dim flg_ruteoA_cnv
    flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
    Dim sCia As String
    Dim rsCia As oraDynaset
    Dim flgFunLocal As String
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    'F.ECASTILLO 17.12.2020
'    If Not objConvenio.VtaCnv_x_Dlv_Empresa_Asume(objUsuario.CodigoEmpresa, objVenta.CodigoConvenio, "1", "1") = objConvenio.ValorPctEmpresa Then
'        If objVenta.FormaPago.Count(1) <= 0 Then
'            MsgBox "No ha seleccionado una forma de pago", vbCritical, App.ProductName
'            frm_VTA_FormaPago.Show
'            Exit Sub
'        End If
'    End If
    
    If objUsuario.EsDelivery = True Then
        
        dblDlvEmpAsume = objConvenio.VtaCnv_x_Dlv_Empresa_Asume(objUsuario.CodigoEmpresa, objVenta.CodigoConvenio, "1", "1")
    
        '*** Cambio para que permita validar cuando se edite el porcentaje copago ***'
        '*** 1002/2009 Por Crueda
        '*** 18/02/09 corregido por pherrera la validacion no funcionaba con venta regular. Permitia grabar proformas sin fp
    '    If objVenta.PctBeneficiario = 0 Then
    '        dblDlvEmpAsume = 100
    '    End If
    
    '    If dblDlvEmpAsume <> objConvenio.ValorPctEmpresa Then
        If objVenta.PctBeneficiario > 0 Or objVenta.ptmModalidad = Venta_Regular Or objVenta.ptmModalidad = Venta_Mayorista Then
            If objVenta.FormaPago.Count(1) <= 0 Then MsgBox "No ha seleccionado una forma de pago", vbCritical, App.ProductName: frm_VTA_FormaPago.Show: frm_VTA_FormaPago.SetFocus: Exit Sub
        End If
    
        '** Para que capture el valor del nombre de la persona que paga con tarjeta **'
        '** Hecho 07/11/2007 Por Cristhian Rueda                                    **'
        
        'objVenta.NomTitular = frm_VTA_FormaPagoTarjeta.PNombreTitular
        'objVenta.NumDNI = frm_VTA_FormaPagoTarjeta.PNumDni
    
        If Trim(txtDLVTelefono.Text) = "" Then
            MsgBox "No ha Ingresado un Cliente", vbCritical, "Aviso"
            Unload frm_VTA_Busqueda
            If frm_VTA_FormaPago.Visible = True Then
                Unload frm_VTA_FormaPago
            End If
            txtDLVTelefono.SetFocus
            Exit Sub
        End If
    
        If objVenta.UbigeoEntrega = "" Then MsgBox "No he grabado el ubigeo", vbCritical, App.ProductName: frm_VTA_Documento.Show:  Exit Sub
        
        'I.ECASTILLO 15.06.2021
        If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
            GoTo cnvNoRuteaAuto
        End If
        Set rsCia = Nothing
        sCia = ""
        Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, objVenta.LocalDespacho)
        If (rsCia.RecordCount > 0) Then
          sCia = CStr(rsCia(1))
        End If
        Set rsCia = Nothing
        flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
        If flg_2e_reserva = "1" And flgFunLocal = "1" Then
            GoTo CallServiceReserva
        End If
        
        objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, objVenta.LocalDespacho, "RESERVA_STOCK_2DA")
        
        If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
        Else
CallServiceReserva:
            'se agrega validaci?n para geolocalizacion, segmentos
            If Len(Trim(objVenta.dc_street)) = 0 Or _
               Len(Trim(objVenta.dc_city)) = 0 Or _
               Len(Trim(objVenta.dc_district)) = 0 Then
                MsgBox "No se ha geolocalizado, favor de no saltarse los pasos", vbCritical, App.ProductName
                frm_VTA_PreviaTomaPedido.Show
                Exit Sub
            End If
            If Len(Trim(objVenta.bk_ServiceType)) = 0 Then
                MsgBox "No escogio Segmentos/Horarios, favor de no saltarse los pasos", vbCritical, App.ProductName
                Exit Sub
            End If
        End If
cnvNoRuteaAuto:
        'F.ECASTILLO 15.06.2021
        
        frm_VTA_Previa.Show vbModal
        
        If frm_VTA_Previa.flgContinua = False Then Exit Sub
        
    End If
    'If objVenta.UbigeoEntrega = "" Then MsgBox "No he grabado el ubigeo", vbCritical, App.ProductName: frm_VTA_Documento.Show: Exit Sub
On Error GoTo handle
    Dim strProforma As String
    Dim strMensaje  As String
    
    If objVenta.EsRegaloTodo = False Then
        
    End If
    'I.ECASTILLO 27.10.2020
    'SI LOCAL ES DC CAPPA RESERVAR CAPACIDAD, SOLO GRABAR SI NO HUBO ERROR
'    If objVenta.isLocalDcCappa = "1" Then
'        If objVenta.flgDatosCapacidad = True Then
'            If reservaCapacidad = "0" Then
'                MsgBox "No se pudo realizar la reserva de capacidad, intente nuevamente o comuniquese con sistemas", vbOKOnly, App.ProductName
'                Exit Sub
'            End If
'        Else
'            MsgBox "Por tratarse de local DC Cappa debe verificar disponibilidad de capacidad", vbCritical + vbOKOnly, App.ProductName
'            frm_VTA_Documento.Show
'            frm_VTA_Documento.SetFocus
'            Exit Sub
'        End If
'    'Else
'        'strMensaje = objVenta.Grabar(godbVentas, strProforma, objVenta.ArchivoVital)
'    End If
    'F.ECASTILLO 27.10.2020
    'I.ECASTILLO 17.12.2020 | Reserva de Stock 2da Etapa
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto2
    End If
    Set rsCia = Nothing
    sCia = ""
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, objVenta.LocalDespacho)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "1" And flgFunLocal = "1" Then
        GoTo CallServiceReserva2
    End If
    
    objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, objVenta.LocalDespacho, "RESERVA_STOCK_2DA")
    
    If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
    Else
CallServiceReserva2:
        'validar si fracciones corresponden al del local
'        If objVenta.isLocalDcCappa = "1" Then
'            If reValidaFracc = "0" Then
'                MsgBox "Existen productos cuyo fraccionamiento no corresponde con el del local, favor de validar", vbCritical, App.ProductName
'                Exit Sub
'            End If
'        End If
'        If objVenta.isLocalDcCappa <> "1" Then
'            If objVenta.flgDatosCapacidad = True Then
'                If reservaCapacidad = "0" And objVenta.bk_chkRET <> "1" Then
'                    MsgBox "No se pudo realizar la reserva de capacidad, intente nuevamente o comuniquese con sistemas", vbOKOnly, App.ProductName
'                    Exit Sub
'                End If
'            End If
'        End If
    End If
cnvNoRuteaAuto2:
    'F.ECASTILLO 17.12.2020
    
    strMensaje = objVenta.Grabar(godbVentas, strProforma, objVenta.ArchivoVital)
    
    If Not strMensaje = "" Then
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    printCantCallsGMaps strProforma, strMensaje
    'I.ECASTILLO 27.10.2020
'    If objVenta.isLocalDcCappa = "1" Then
'        If asignaDigital(strProforma) = "0" Then
'            MsgBox "No se pudo asignar pedido a Digital" & vbNewLine & _
'                    "========================" & vbNewLine & _
'                    "Por favor verificar la bandeja de pendientes." & vbNewLine & _
'                    "Proforma: " & strProforma, vbOKOnly + vbCritical, App.ProductName
'            'debe liberar reserva (retornar capacidad)
'            'If retornaCapacidad = False Then
'            '    MsgBox "Al no poder asigarse el pedido a Digital se intento retornar la capacidad reservada," & vbNewLine & _
'            '            "pero esto no fue posible, favor de comunicarse con sistemas.", vbOKOnly, App.ProductName
'            'Else
'                'se retorno la capacidad
'            'End If
'            'Exit Sub
'        Else
'            If Len(Trim(objVenta.dc_orderDigital)) > 0 Then
'                msgAsignaOrderDigProf = objProforma.asignaOrderDigitalProf(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strProforma, objVenta.dc_orderDigital, "1")
'            End If
'        End If
'    End If
    'F.ECASTILLO 27.10.2020
    
    'I.ECASTILLO 17.12.2020 | Reserva de Stock 2da Etapa
    'asignar a digital solo si todos los productos tienen stock suficiente en local despacho
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto3
    End If
    If flgFunLocal = "1" Then
        GoTo CallServiceReserva3
    End If
    
    If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
    Else
CallServiceReserva3:
'        If objVenta.isLocalDcCappa <> "1" Then
            Dim necesita_transf As Boolean
            Dim xx, ii, jj As Integer
            xx = 0: necesita_transf = False
            For xx = 0 To objVenta.Producto2.UpperBound(1)
                Debug.Print objVenta.Producto2(xx, 0)
                Debug.Print objVenta.Producto2(xx, 30)
                Debug.Print objVenta.Producto2(xx, 31)
                If objVenta.Producto2(xx, 31) = "1" Then necesita_transf = True: Exit For
            Next xx
            If necesita_transf = False Then
                If asignaDigital(strProforma) = "0" Then
                    MsgBox "No se pudo asignar pedido a Digital" & vbNewLine & _
                            "========================" & vbNewLine & _
                            "Por favor verificar la bandeja de pendientes." & vbNewLine & _
                            "Proforma: " & strProforma, vbOKOnly + vbCritical, App.ProductName
                Else
                    If Len(Trim(objVenta.dc_orderDigital)) > 0 Then
                        msgAsignaOrderDigProf = objProforma.asignaOrderDigitalProf(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strProforma, objVenta.dc_orderDigital, "1")
                        
                        If gstrFlagReservaCap = "1" Then
                            If objVenta.flgDatosCapacidad = True Then
                                If reservaCapacidad(objVenta.dc_orderDigital) = "0" And objVenta.bk_chkRET <> "1" Then
                                    MsgBox "No se pudo realizar la reserva de capacidad.", vbOKOnly, App.ProductName
                                    'Exit Sub
                                    gstrFlagReservaCap = "2"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
'        End If
    End If
cnvNoRuteaAuto3:
    'F.ECASTILLO 17.12.2020
    
    'Dim objProforma As New clsProforma
    Dim x As String
    Dim strCadCodProducto As String
    Dim strCadCtdProducto  As String
    Dim strCadCtdProductoFrac  As String
    Dim strLocalOrigen  As String
    Dim strLocalDestino As String
    Dim i As Integer
    i = 0
    
    While i < objVenta.Distribucion.UpperBound(1)
        strCadCodProducto = strCadCodProducto & objVenta.Distribucion(i, 1) & "|"
        strCadCtdProducto = strCadCtdProducto & objVenta.Distribucion(i, 4) & "|"
        strCadCtdProductoFrac = strCadCtdProductoFrac & objVenta.Distribucion(i, 4) & "|"
        strLocalOrigen = strLocalOrigen & objVenta.Distribucion(i, 8) & "|"
        strLocalDestino = strLocalDestino & objVenta.Distribucion(i, 9) & "|"
        i = i + 1
    Wend
    
    If Not objVenta.Distribucion.UpperBound(1) = -1 Then
        x = objProforma.GrabaTransferencia(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strProforma, strCadCodProducto, strCadCtdProducto, strCadCtdProductoFrac, strLocalOrigen, strLocalDestino, objUsuario.Codigo)
    End If
    
    MsgBox "Se grabo satisfactoriamente la Proforma :" & strProforma, vbInformation, App.ProductName
    
    If objUsuario.EsDelivery = False Then
        objDocumento.ImprimirDocumento objVenta.CodigoDocumentoVenta, strProforma
    End If
    
    Set objProforma = Nothing
'    ==== JLOPEZ CAMBIO PARA QUE MUESTRE PANTALLA
'    ==== EN MODO DELIVERY PROV
 '   ==== 08/01/2007
'    subNuevo
'    frm_VTA_Busqueda.grdProductos.Limpiar
'    frm_VTA_Busqueda.grdAlternativos.Limpiar
'    frm_VTA_Busqueda.grdComplementarios.Limpiar
'    cmdModalidad_Click
    
    Call Salida

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    Set objProforma = Nothing
End Sub
Public Sub cmdGrabaVenta_Click() 'Optional ByVal formulario As String = "")

    Dim CantImp As Integer
    Dim u As Integer
    Dim objProducto As New clsProducto
    Dim strSecVentas As String
    u = 0

    While u < objVenta.Producto.Count(1)
        If objVenta.NumMaximoUnidades > 0 And objVenta.Producto(u, 6) = 0 And val(objVenta.Producto(u, 3)) > objVenta.NumMaximoUnidades Then
            MsgBox "El convenio no permite seleccionar m?s de " & objVenta.NumMaximoUnidades & " unidades por producto", vbCritical, App.ProductName
            Exit Sub
        ElseIf objVenta.NumMaximoUnidades > 0 And objVenta.Producto(u, 6) = 1 Then
            Dim xCtdFrac As Integer
            Dim xCtdUnidades As Double
            xCtdFrac = objProducto.intCtdFrac(objVenta.Producto(u, 0))
            xCtdUnidades = val(val(objVenta.Producto(u, 3))) / xCtdFrac
            If val(val(xCtdUnidades)) > objVenta.NumMaximoUnidades Then
                MsgBox "El convenio no permite seleccionar m?s de " & objVenta.NumMaximoUnidades & " unidades por producto", vbCritical, App.ProductName
                Exit Sub
            End If
        End If
        u = u + 1
    Wend
    'MsgBox formulario

    Dim strNumProforma As String
    Dim oTipDoc As OraParamArray
    Dim oNumDoc As OraParamArray
    Dim oTipDocCo As OraParamArray
    Dim oNumDocCo As OraParamArray
    Dim auxTipDoc As OraParamArray
    Dim auxNumDoc As OraParamArray
    Dim varMsgDoc As Variant
    Dim varMsgDocCo As Variant
    Dim objImpresion As New clsDocumento
    Dim objLocal As New clsLocal
    Dim strFlgImprimeAlm As String
    Dim rsLocal As oraDynaset
    Dim Devicename As String
    Dim bolEncontroImpAlm As Boolean
    Dim UltDocEmitido As oraDynaset
    Dim strUltDocEmi As String
    Dim i, j As Integer
    Dim orsListaDocumentoMaquina As oraDynaset
    Dim strFlgImprimeDoc As String
    Dim UbicaPrinter As String
    Dim Impresoras As Printer
    Dim bolPagarNavsatConEfe As Boolean
    Dim Y, z As Integer
    Dim RetPromoMensaje As String
    Dim strCajaAbierta As String
    Dim NomImpresora As String

    On Error GoTo CtrlErr
    
    Set rsLocal = objLocal.Lista(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
    
    If Not rsLocal.EOF Then strFlgImprimeAlm = IIf(IsNull(rsLocal("FLG_IMPRIME_ALM").Value), "0", rsLocal("FLG_IMPRIME_ALM").Value)

    If objVenta.CodigoDocumentoVenta = "PRO" Then grabaProforma: Exit Sub
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Bitacora :
    '   Creado el por jlopez
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    If objVenta.CodigoTipoVenta = Venta_Convenio And objVenta.CodigoConvenio = "" Then
        MsgBox "La modalidad de venta es Convenio , no ha seleccionado uno..", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Bitacora :
    '   Creado el 22/06/07 por crueda
    'If objUsuario.TipoMaquina = objUsuario.TipoMaquinaCabina Then grabaProforma: Exit Sub
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    If objUsuario.EsDelivery = True Then grabaProforma: Exit Sub
    
    'MsgBox "Modalidad de Venta " & objVenta.ptmModalidad & "", vbExclamation, Caption: Exit Sub
    
    '** Modalidad Venta (Regular, Mayorista, Cotizaci?n) **'
'    If (objVenta.CodigoTipoVenta = Venta_Mayorista Or objVenta.CodigoTipoVenta = Venta_Regular Or objVenta.CodigoTipoVenta = Cotizaciones) And objVenta.Producto.UpperBound(1) = -1 Then
'        MsgBox "La Modalidad debe ser " & Chr(13) & _
'               "        Venta Regular   o'" & Chr(13) & _
'               "        Venta Mayorista o'" & Chr(13) & _
'               "        Cotizaci?n ", vbExclamation, App.ProductName
'        Exit Sub
'    End If
    
    '** Validaciones Hechas el 23/08/2007 Por Crueda **'
    '** Modalidad Servicio **'
    If objVenta.CodigoTipoVenta = Servicio And objVenta.Servicios.UpperBound(1) = -1 Then
        MsgBox "La modalidad de venta es Servicio , no ha seleccionado uno..", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+  Bitacora :
    '+  Creado el 08/07/2008 por jlopez
    '+  Se valida si es una transacci?n navsat que el pago solo sea EFECTIVO
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    If objVenta.CodigoTipoVenta = Servicio And objVenta.Servicios.UpperBound(1) > -1 Then
        bolPagarNavsatConEfe = True
        For z = 0 To objVenta.Servicios.UpperBound(1)
           If objVenta.EsNavsat(objVenta.Servicios(z, 0)) Then
                For Y = 0 To objVenta.FormaPago.UpperBound(1)
                    If objVenta.FormaPago(Y, 0) <> gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_EFE") _
                        And objVenta.FormaPago(Y, 0) <> gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_RED") Then
                        bolPagarNavsatConEfe = False
                        Exit For
                    End If
                Next
                If Not bolPagarNavsatConEfe Then Exit For
           End If
        Next
    
        If Not bolPagarNavsatConEfe Then
            MsgBox "Las recargas virtuales NAVSAT, solo se pagan en EFECTIVO", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If
    
    '** Modalidad Recetario Magistral **'
    'If objVenta.CodigoTipoVenta = Recetario And (objVenta.RecetarioMagistral.UpperBound(1) = -1) Or (objVenta.RecetarioMagistral.UpperBound(1) = 0 And objVenta.Producto.UpperBound(1) = 0) Then
    
    If objVenta.CodigoTipoVenta = Recetario And objVenta.RecetarioMagistral.UpperBound(1) = -1 Then
        MsgBox "La modalidad de venta es Recetario Magistral , no ha seleccionado uno.. ", vbExclamation, App.ProductName
        Exit Sub
    End If
    '***************************************************'
    
    '** Validaci?n para que no permita vernder estando en la modalidad de Cajero Correspondal **'
    '** Fecha 19/11/2007 Por Cristhian Rueda                                                  **'
    If objVenta.CodigoTipoVenta = Cajero_Corresponsal Then
        MsgBox "La modalidad de venta es Cajero Corresponsal , no ha seleccionado uno.. ", vbExclamation, App.ProductName
        Exit Sub
    End If
        
    If objVenta.Totales(2) < 0 Then
       MsgBox "El Total a pagar es menor al Monto Total", vbInformation, App.ProductName
       Exit Sub
    End If
    
'    gclsOracle.Cerrar
'    If gclsOracle.Conexion(gvarTNSNAME, gvarUSUARIO, gvarPASSWORD) <> "" Then End
    
   
        If strcodConv = "" Then
'            If objVenta.FormaPago.Count(1) <= 1 Then
'            MsgBox "Ingrese forma de Pago", vbCritical, App.ProductName
'            Me.cmdFormaPago_Click
'            Exit Sub
'            End If
            If objVenta.EsRegaloTodo = False Then
                If objVenta.FormaPago.Count(1) <= 1 And objVenta.CodigoConvenio = "" And objVenta.ptmModalidad <> Cobro_Responsabilidad Then MsgBox "Ingrese forma de Pago", vbCritical, App.ProductName: frm_VTA_FormaPago.Show: Exit Sub
            End If
        End If
        If objVenta.Producto.Count(1) <= 0 Then MsgBox "Debe de seleccionar algun producto", vbCritical, App.ProductName: Exit Sub
        
        If objVenta.PctBeneficiario > 0 Then
                    If objVenta.flgPctBeneficiario <> "0" Then
                        If objVenta.fnRedondeo(objVenta.ImpPctBeneficiario) > objVenta.Totales(0) Then
                            MsgBox "El Importe de Co-Pago No pude ser mayor al monto", vbCritical, App.ProductName
                            Exit Sub
                        End If
                    End If
        End If
    
    If (objVenta.CodigoTipoVenta = Venta_Regular Or objVenta.CodigoTipoVenta = Venta_Convenio) And objVenta.HayFarmacos And "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "PIDERECETA", "10") = "1" Then
        frm_VTA_DatoReceta.Show vbModal
    End If
    '+ Modificaci?n 07/01/2009 Por Crueda
    '+ Permite que se haga unas validaciones previo al registro de la carga
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    strCajaAbierta = objDocumento.ValidaExisteLiquidacion(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objUsuario.CodigoLocal, objUsuario.Codigo)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++ Se aumento un parametro al GrabaDoc para que evalue segun el pctBeneficario en caso de convenio que este marcado deducible en base al precio publico
    '++ 09/01/2008 Por Cristhian Rueda
    Dim FlgAnulados As Boolean
    If objVenta.GrabarDoc(gclsOracle.ODataBase, oTipDoc, oNumDoc, oTipDocCo, oNumDocCo, RetPromoMensaje, , strcodConv, objVenta.NewPctBeneficario, "", FlgAnulados) = False Then Exit Sub
 '   If TipoDoc = "PRO" Then
 '       Dim MsJ
 '       MsJ = objProforma.CambiaEstado(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, xNumProforma, "005", objUsuario.Codigo)
 '   End If
 
     If RetPromoMensaje <> "" Then
         frmMensajePromo.pRetMensaje = RetPromoMensaje
         frmMensajePromo.Show vbModal
     End If
 
        
    objVenta.OrdenaDoc gclsOracle.ODataBase, oTipDoc, oNumDoc, oTipDocCo, oNumDocCo, auxTipDoc, auxNumDoc
    
    pdblPagTkt = 0: pdblTPagTkt = 0
    
    For i = 0 To oTipDoc.ArraySize - 1
        If oTipDoc.get_Value(i) <> "" Then
            If oTipDoc.get_Value(i) = "REC" Then
                Dim objServicio As New clsServicio
                Dim strNumOperacion As String
                Dim rsServicio As oraDynaset
                Set rsServicio = objServicio.ListaCobranza(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, oTipDoc.get_Value(i), oNumDoc.get_Value(i))
                strNumOperacion = "" & rsServicio("NUM_VOUCH_OPE").Value
                Set rsServicio = Nothing
                varMsgDoc = varMsgDoc & oTipDoc.get_Value(i) & " " & oNumDoc.get_Value(i) & " -->N? Operacion: " & strNumOperacion & Chr(13)
                Set objServicio = Nothing
            Else
                pdblTPagTkt = pdblTPagTkt + 1
                varMsgDoc = varMsgDoc & oTipDoc.get_Value(i) & " " & oNumDoc.get_Value(i) & Chr(13)
            End If
        '/////agregado por miguel laguna con pherrera 29/03/10////////
        Else
            Exit For
        '/////////////////////////////////////////////////
        End If
    Next i
    
    For i = 0 To oTipDocCo.ArraySize - 1
        If oTipDocCo.get_Value(i) <> "" Then
        varMsgDocCo = varMsgDocCo & oTipDocCo.get_Value(i) & " " & oNumDocCo.get_Value(i) & Chr(13)
        '/////agregado por miguel laguna con pherrera 29/03/10////////
        Else
            Exit For
        '///////////////////////////////////////////////////////
        End If
    
    Next i
    
    
    If Not FlgAnulados = True Then
    
    'Actualizacion 23/08/2012 para IMPRESORA LOCAL O RED MLEVANO
    NomImpresora = "" & gclsOracle.FN_Valor("BTLPROD.FN_NOM_IMPRESORA_TKB", objUsuario.NombrePC)
    If Not NomImpresora = "" Then
        Call objDocumento.LocalORed(objUsuario.CodigoEmpresa, objUsuario.NombrePC, "TKB")
    End If
    'Fin de la actualizacion
    
    MsgBox "Se realizo la transacci?n satisfactoriamente  - " & Chr(13) & Chr(13) & varMsgDoc & _
            IIf(varMsgDocCo = "", "", "Por convenio - " & varMsgDocCo), vbInformation + vbOKOnly, App.ProductName
    End If
            
            Dim lk As Integer
            lk = 0
            i = 0
            Dim xDocumento As New XArrayDB
            xDocumento.ReDim 0, -1, 0, 1
            While lk < auxTipDoc.ArraySize - 1
                If Not auxTipDoc(i) = "" Then
                    xDocumento.AppendRows
                    xDocumento(lk, 0) = auxTipDoc(i)
                    xDocumento(lk, 1) = auxNumDoc(i)
                End If
                i = i + 1
                lk = lk + 1
            Wend
            xDocumento.QuickSort xDocumento.LowerBound(1), xDocumento.UpperBound(1), 0, XORDER_ASCEND, XTYPE_STRING, 1, XORDER_ASCEND, XTYPE_STRING
            i = 0
    strUltDocEmi = ""
    strFlgImprimeDoc = ""
        
    For i = 0 To auxTipDoc.ArraySize - 1
        If auxTipDoc.get_Value(i) <> "" Then
            Set UltDocEmitido = objDocumento.UltDocEmitido(objUsuario.CodigoEmpresa, objUsuario.NombrePC)
            If Not UltDocEmitido.EOF Then strUltDocEmi = UltDocEmitido("COD_TIPO_DOCUMENTO").Value
            
    ''        Set orsListaDocumentoMaquina = objImpresion.ListaDocumentoMaquina(objUsuario.CodigoEmpresa, auxTipDoc.get_Value(i), objUsuario.NombrePC)
    ''
    ''        If orsListaDocumentoMaquina.EOF Then
    ''            strFlgImprimeDoc = "1"
    ''        Else
    ''            strFlgImprimeDoc = orsListaDocumentoMaquina("FLG_IMPRIME_DOC").Value
    ''        End If
            
            strFlgImprimeDoc = objImpresion.ImprimeRecibo(objUsuario.CodigoEmpresa, xDocumento(i, 0), xDocumento(i, 1))

            
            If xDocumento(i, 0) <> "" And strFlgImprimeDoc = "1" Then
                If xDocumento(i, 0) <> strUltDocEmi Then
                    MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                            "en posicion de " & xDocumento(i, 0), vbInformation, App.ProductName
                End If

                UbicaPrinter = Printer.Devicename
                frm_SERV_Impresoras.gNroCopia = 1
                If xDocumento(i, 0) = objUsuario.TipoDocRecibo Then
                          frm_SERV_Impresoras.Show vbModal
                End If

                

                For j = 1 To frm_SERV_Impresoras.gNroCopia
                        '''''para las colas de impresion
                    Dim objImpre As New clsImpresion
                    If objImpre.PuedeImprimir(objUsuario.CodigoEmpresa, objUsuario.NombrePC, xDocumento(i, 0)) = True Then
                    If objImpre.ListaPendiente(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.NombrePC, xDocumento(i, 0)) > 0 Then frmColaImpresion.Show vbModal: GoTo ira
                             pdblPagTkt = pdblPagTkt + j
                             ''PRUEBA DE TICKET PROFORMA
                             'If objVenta.NumPedidoPadre <> "" And objUsuario.CodTipoVenta = objUsuario.TipoVentaDlv Then  ''ES EL ULTIMO i = auxTipDoc.ArraySize - 1 And
                                'objImpresion.ImprimirDocumento xDocumento(i, 0), xDocumento(i, 1), "", objVenta.CodModalidadVenta, FlgAnulados, True
                             'Else ''ASI ESTABA
                                objImpresion.ImprimirDocumento xDocumento(i, 0), xDocumento(i, 1), "", objVenta.CodModalidadVenta, FlgAnulados
                                'strSecVentas = objVenta.fnDevuelveSecuencia(objUsuario.CodigoEmpresa, xDocumento(i, 0), xDocumento(i, 1))
                             'End If
                    End If
                    strSecVentas = objVenta.fnDevuelveSecuencia(objUsuario.CodigoEmpresa, xDocumento(i, 0), xDocumento(i, 1))
                    Set objImpre = Nothing
                Next j

                For Each Impresoras In Printers
                    If UbicaPrinter = Impresoras.Devicename Then Set Printer = Impresoras: Exit For
                Next
                objDocumento.GrabaUltDocEmitido objUsuario.CodigoEmpresa, objUsuario.NombrePC, xDocumento(i, 0)
            End If
            
'''''''''''''''            If auxTipDoc.get_Value(i) <> "" And strFlgImprimeDoc = "1" Then
'''''''''''''''                If auxTipDoc.get_Value(i) <> strUltDocEmi Then
'''''''''''''''                    MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
'''''''''''''''                            "en posicion de " & auxTipDoc.get_Value(i), vbInformation, App.ProductName
'''''''''''''''                End If
'''''''''''''''
'''''''''''''''                UbicaPrinter = Printer.Devicename
'''''''''''''''                frm_SERV_Impresoras.gNroCopia = 1
'''''''''''''''                If auxTipDoc.get_Value(i) = objUsuario.TipoDocRecibo Then
'''''''''''''''                          frm_SERV_Impresoras.Show vbModal
'''''''''''''''                End If
'''''''''''''''
'''''''''''''''
'''''''''''''''                For j = 1 To frm_SERV_Impresoras.gNroCopia
'''''''''''''''                        '''''para las colas de impresion
'''''''''''''''                    Dim objImpre As New clsImpresion
'''''''''''''''                    If objImpre.PuedeImprimir(objUsuario.CodigoEmpresa, objUsuario.NombrePC, auxTipDoc.get_Value(i)) = True Then
'''''''''''''''                    If objImpre.ListaPendiente(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.NombrePC, auxTipDoc.get_Value(i)) > 0 Then frmColaImpresion.Show vbModal: GoTo ira
'''''''''''''''                             pdblPagTkt = pdblPagTkt + j
'''''''''''''''                             objImpresion.ImprimirDocumento auxTipDoc.get_Value(i), auxNumDoc.get_Value(i)
'''''''''''''''                    End If
'''''''''''''''                    Set objImpre = Nothing
'''''''''''''''                Next j
'''''''''''''''
'''''''''''''''
'''''''''''''''
'''''''''''''''                For Each Impresoras In Printers
'''''''''''''''                    If UbicaPrinter = Impresoras.Devicename Then Set Printer = Impresoras: Exit For
'''''''''''''''                Next
'''''''''''''''                objDocumento.GrabaUltDocEmitido objUsuario.CodigoEmpresa, objUsuario.NombrePC, auxTipDoc.get_Value(i)
'''''''''''''''            End If
        End If
    Next i
            
    '+++++++'
    
    'Agregado por Jose Melgar, para Imprimir Ticket Termico
    'Fecha: 09/05/2012
    If objVenta.CodigoTipoVenta = Venta_Convenio Then
       Dim StrImprime As String
       Dim strNombreTermica As String
       Dim bolValorTermica As Boolean

       StrImprime = objDocumento.ListaDocCredito(strSecVentas)

       If StrImprime = "1" Then

          strNombreTermica = objVenta.NombreTermica(objUsuario.NombrePC)

          If strNombreTermica = "" Then
             MsgBox "No tiene la configurada la Impresora Termica en el Maestro de Maquinas", vbCritical, App.ProductName
          Else
              bolValorTermica = False
              For Each Impresoras In Printers
                  If UCase(strNombreTermica) = UCase(Impresoras.Devicename) Then
                     Set Printer = Impresoras
                     bolValorTermica = True
                  End If
              Next

              If bolValorTermica Then
                 MsgBox "Se proceder? a generar el Voucher Convenio" & Chr(13) & "Prepare su impresora.", vbInformation, App.ProductName
                 fn_ImprimeTicketTermico strSecVentas
              Else
                  MsgBox "No tiene la impresora termica o tiene configurado otro nombre en el maestro de maquinas", vbCritical, App.ProductName
              End If
           End If
       End If


    End If
  
    
    If objVenta.NumPedidoPadre <> "" And objUsuario.CodTipoVenta = objUsuario.TipoVentaDlv Then
        CantImp = objProforma.Cuenta_Impresoras_x_Maquina(objUsuario.CodigoEmpresa, objUsuario.NombrePC)
        
        If CantImp <= 0 Then '***
            frm_SERV_Impresoras.Show vbModal
            'UbicaPrinter = Printer.Devicename
            If frm_SERV_Impresoras.gNroCopia > 0 Then
               MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                      "en posici?n de Pedido Delivery", vbInformation, App.FileDescription
            End If
            
            Dim k%
            ''For k = 0 To 1 - 1
            For k = 1 To frm_SERV_Impresoras.gNroCopia
                'MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                       "en posici?n de Pedido Delivery", vbInformation, App.FileDescription
                'objProforma.Imprimir objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.NumPedidoPadre
                objProforma.ImprimirDelivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.NumPedidoPadre, "1"
            Next k
        ElseIf CantImp > 0 Then '***
            'mlaguna 05/04/2010 Envia directamente a la ticketera
            objProforma.ImprimirTicket_Delivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.NumPedidoPadre, "1", True
            'objProforma.ImprimirTicket_DeliveryNew objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.NumPedidoPadre, "1", True
        End If                                               '***
    End If
    
    '+++++++'
            
    For i = 0 To auxTipDoc.ArraySize - 1
        If auxTipDoc.get_Value(i) <> "" Then
            UbicaPrinter = Printer.Devicename
            
            
            Devicename = objLocal.NombreImpresoraAlmacen
            bolEncontroImpAlm = False
            For Each Impresoras In Printers
                If UCase(Impresoras.Devicename) = Trim(UCase(Devicename)) Then
                    Set Printer = Impresoras
                    bolEncontroImpAlm = True
                    Exit For
                End If
            Next Impresoras
            
            If strFlgImprimeAlm = "1" And (auxTipDoc.get_Value(i) = objUsuario.TipoDocBol Or auxTipDoc.get_Value(i) = objUsuario.TipoDocFac Or auxTipDoc.get_Value(i) = objVenta.TipoDocTKB Or auxTipDoc.get_Value(i) = objVenta.TipoDocTKF) And bolEncontroImpAlm Then
                                objDocumento.ImprimeDocAlamcenLocal objUsuario.CodigoEmpresa, _
                                            auxTipDoc.get_Value(i), _
                                            auxNumDoc.get_Value(i), _
                                            objUsuario.CodigoLocal
            End If
            For Each Impresoras In Printers
                If UbicaPrinter = Impresoras.Devicename Then Set Printer = Impresoras: Exit For
            Next
        '/////agregado por miguel laguna con pherrera 29/03/10////////
        Else
            Exit For
        '/////////////////////////////////////////////////////////////
        End If
    Next i
    '+++++++'
    Dim strMensajeRemesa As String
    strMensajeRemesa = objVenta.MensajeRemesa(objUsuario.Codigo, objUsuario.NombrePC)
    If strMensajeRemesa <> "" Then
        MsgBox strMensajeRemesa, vbCritical, App.ProductName
    End If
    '' aca imprimi los cupones
    '' 04-ABR-2014 TCT Proceso  de Impresion de Cupones
    Dim rsCupones As oraDynaset
    Dim strNombreCuponera As String
    
    If objVenta.EsImprCupon(objUsuario.CodigoLocal) = True Then
            If strSecVentas <> "0" Then
               Set rsCupones = objDocumento.ListaCupones(strSecVentas)
                    While Not rsCupones.EOF
                    strNombreCuponera = objVenta.NombreCuponera(objUsuario.NombrePC)
                        If strNombreCuponera = "" Then
                            MsgBox "No tiene la impresora de cupones Instalada", vbCritical, App.ProductName
                            Exit Sub
                        Else
                        Dim bolValor As Boolean
                        bolValor = False
                            For Each Impresoras In Printers
                                If UCase(strNombreCuponera) = UCase(Impresoras.Devicename) Then
                                Set Printer = Impresoras
                                bolValor = True
: Exit For
                                End If
                            Next
                            If bolValor Then
                            MsgBox "Se proceder? a generar Cup?n de Descuento." & Chr(13) & "Prepare su impresora.", vbInformation, App.ProductName
                            
                            ''' <30-ABR-14   TCT   Add indicador  de Impresion de Cupon>
                            Dim rs_01 As oraDynaset
                            Dim v_str As String
                            Set rs_01 = objDocumento.ListaCupon("" & rsCupones("COD_DOCUMENTO_PAGO"), "" & rsCupones("NUM_DOCUMENTO_PAGO"))
                            v_str = left(rs_01("DES_DOCUMENTO_PAGO"), 6)
                            If v_str = "LLENAR" Then
                             fnImprimeCupon2 "" & rsCupones("COD_DOCUMENTO_PAGO"), "" & rsCupones("NUM_DOCUMENTO_PAGO")
                            Else
                             fnImprimeCupon "" & rsCupones("COD_DOCUMENTO_PAGO"), "" & rsCupones("NUM_DOCUMENTO_PAGO")
                            
                            End If
                            ''' </30-ABR-14   TCT   Add indicador  de Impresion de Cupon>
                            
                            'fnImprimeCupon "" & rsCupones("COD_DOCUMENTO_PAGO"), "" & rsCupones("NUM_DOCUMENTO_PAGO")
                            'fnImprimeCupon2 "" & rsCupones("COD_DOCUMENTO_PAGO"), "" & rsCupones("NUM_DOCUMENTO_PAGO")
                            Else
                            MsgBox "No tiene la impresora de cupones Instalada o tiene diferente nombre", vbCritical, App.ProductName
                            Exit Sub
                            End If
                        End If
                        rsCupones.MoveNext
                    Wend
            End If
    End If
ira:
   ''' If Not strNombreFomularioOrigen = "frm_DLV_Pedido" Then
        Call Salida
         '''Else
        '''objVenta.LimpiaProductos
        '''strNombreFomularioOrigen = ""
        
    '''End If
      
    Exit Sub
CtrlErr:

    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
    If objUsuario.CodTipoVenta = objUsuario.TipoVentaDlv And objVenta.NumPedidoPadre <> "" Then  'PARA DELIVERY
        Call Salida
    End If
    'If Err.Number <> -2147221503 Then Call Salida

End Sub

Sub justifica_printer(x0, xf, y0, txt)
 ' x0, xf = posicion de los margenes izquierdo y derecho
' y0 = posicion vertical donde se desea empezar a escribir
' txt = texto a escribir

Dim x, Y, k, Ancho
Dim s As String, ss As String
Dim x_spc

s = txt
x = x0
Y = y0
Ancho = (xf - x0)

While s <> ""

ss = ""
While (s <> "") And (Printer.TextWidth(ss) <= Ancho)
ss = ss & left$(s, 1)
s = right$(s, Len(s) - 1)
Wend
If (Printer.TextWidth(ss) > Ancho) Then
s = right$(ss, 1) & s
ss = left$(ss, Len(ss) - 1)
End If
' aqui tenemos en ss lo maximo que cabe en una linea
If right$(ss, 1) = " " Then
ss = left$(ss, Len(ss) - 1)
Else
If (InStr(ss, " ") > 0) And (left$(s & " ", 1) <> " ") Then
While right$(ss, 1) <> " "
s = right$(ss, 1) & s
ss = left$(ss, Len(ss) - 1)
Wend
ss = left$(ss, Len(ss) - 1)
End If
End If
x_spc = 0
x = x0
If (Len(ss) > 1) And (s & "" <> "") Then
x_spc = (Ancho - Printer.TextWidth(ss)) / (Len(ss) - 1)
End If
Printer.CurrentX = x
Printer.CurrentY = Y

If x_spc = 0 Then
Printer.Print ss;
Else
For k = 1 To Len(ss)
Printer.CurrentX = x
Printer.Print Mid$(ss, k, 1);
x = x + Printer.TextWidth("*" & Mid$(ss, k, 1) & "*") - Printer.TextWidth("**")
x = x + x_spc
Next
End If

Y = Y + Printer.TextHeight(ss)
While left$(s, 1) = " "
s = right$(s, Len(s) - 1)
Wend
Wend

End Sub

Function fnImprimeCupon(ByVal CodigoDocDescto As String, ByVal NumDocDescto As String)
    
    Dim rs As oraDynaset
    Dim arr() As String
    Dim i As Byte
    Dim strFrase As String
    
    strFrase = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "FRASELEGAL", "10")
    
    Set rs = objDocumento.ListaCupon(CodigoDocDescto, NumDocDescto)
     arr = Split(rs("DES_SUBTITULO"), "|")
    
    'imgLogoBtl.Picture = LoadPicture("\\10.85.8.2\proyector\logos\1.bmp")
    
    If objVenta.EsMFA(objUsuario.CodigoLocal) = True Then
        imgLogoBtl.Picture = Me.ImageList1.ListImages(3).Picture
        Printer.PaintPicture imgLogoBtl, 570, 20, 2950, 850
    Else
        imgLogoBtl.Picture = Me.ImageList1.ListImages(1).Picture
        Printer.PaintPicture imgLogoBtl, 1000, 20, 2150, 950
    End If
    
    
    Printer.CurrentY = 1154
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    Printer.Font.Size = 12
    Printer.FontName = "IDAutomation.com Code39"  'LETRA DE CODIGO DE BARRA'
    Printer.Print Space(2) & "*" & rs("NUM_DOCUMENTO_PAGO2") & "*"
    Printer.FontName = "Courier New"
    Printer.Font.Size = 15
    Printer.Font.Bold = True
    Printer.Print Space(5) & rs("DES_TITULO")
    'Printer.Print Space(6) & Split(rs("DES_SUBTITULO"), "|")
    For i = 0 To UBound(arr)
       Printer.Font.Size = 10
       Printer.Print pfstr_Alineado(Trim(arr(i)), 34, centro)
    Next
    
    Printer.Font.Size = 10
    Dim strCliente As String
    strCliente = "" & rs("DES_CLIENTE")
    Dim intConstante As Integer
    intConstante = 0
    If (strCliente <> "") Then
    Printer.Print Space(6) & "Para: " & strCliente
    intConstante = 184
    End If
    Printer.Font.Size = 8
    Dim strMensaje As String
    strMensaje = "" & rs("DES_CONTENIDO")
    'justifica_printer 20, 4000, 2842 + intConstante, strMensaje
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Font.Bold = False
    Printer.Print ""
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, Space(4) & rs("NUM_DOCUMENTO_PAGO") & Space(10) & "Nro." & rs("SEC_VENTA_EMI")
    Printer.Print ""
    strMensaje = Space(1) & Chr(34) & rs("DES_FRASE") & Chr(34) + strFrase
    Printer.Font.Size = 7
    Printer.Font.Bold = True
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Print ""
    Printer.FontBold = False
    strMensaje = "" & rs("DES_LEGAL")
    Printer.Font.Size = 7
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Print ""
    Printer.Font.Size = 6
    strMensaje = Space(2) & rs("COD_USUARIO") & Space(8) & rs("FCH_REGISTRA") & Space(9) & "BTL" & rs("COD_LOCAL_EMISOR")
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Print
    Printer.EndDoc
    MsgBox "Se ha generado el Cup?n N? " & "" & rs("NUM_DOCUMENTO_PAGO2"), vbInformation, App.ProductName
End Function


Function fnImprimeCupon2(ByVal CodigoDocDescto As String, ByVal NumDocDescto As String)
    
    Dim rs As oraDynaset
    Dim rsDatos As oraDynaset
    Dim arr() As String
    Dim i As Byte
    Dim strFrase As String
    
    strFrase = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "FRASELEGAL", "10")
    
    Set rs = objDocumento.ListaCupon(CodigoDocDescto, NumDocDescto)
    If Not IsNull(rs("DES_SUBTITULO")) Then arr = Split(rs("DES_SUBTITULO"), "|")
    
    'imgLogoBtl.Picture = LoadPicture("\\10.85.8.2\proyector\logos\1.bmp")
    
    If objVenta.EsMFA(objUsuario.CodigoLocal) = True Then
        imgLogoBtl.Picture = Me.ImageList1.ListImages(3).Picture
        Printer.PaintPicture imgLogoBtl, 570, 20, 2950, 850
    Else
        imgLogoBtl.Picture = Me.ImageList1.ListImages(1).Picture
        Printer.PaintPicture imgLogoBtl, 1000, 20, 2150, 950
    End If
    
    
    Printer.CurrentY = 1154
    Printer.FontName = "Courier New"
    Printer.Font.Size = 15
    Printer.Font.Bold = True
    Printer.Print Space(5) & rs("DES_TITULO")
    'Printer.Print Space(6) & Split(rs("DES_SUBTITULO"), "|")
    If Not IsNull(rs("DES_SUBTITULO")) Then
        For i = 0 To UBound(arr)
           Printer.Font.Size = 10
           Printer.Print pfstr_Alineado(Trim(arr(i)), 34, centro)
        Next
    End If
    
    Printer.Font.Size = 10
    Dim strCliente As String
    strCliente = "" & rs("DES_CLIENTE")
    Dim intConstante As Integer
    intConstante = 0
    If (strCliente <> "") Then
    Printer.Print Space(6) & "Para: " & strCliente
    intConstante = 184
    End If
    Printer.Font.Size = 8
    Dim strMensaje As String
    strMensaje = "" & rs("DES_CONTENIDO")
    'justifica_printer 20, 4000, 2842 + intConstante, strMensaje
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Font.Bold = False
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    
    Set rsDatos = objDocumento.ListaDatosCupon(CodigoDocDescto)
    For i = 1 To rsDatos.RecordCount
        Printer.Print rsDatos("des_nombre")
        Printer.Print ""
        rsDatos.MoveNext
    Next
    
    Printer.Print ""
    Printer.Print ""
    strMensaje = Space(1) & Chr(34) & rs("DES_FRASE") & Chr(34) + strFrase
    Printer.Font.Size = 7
    Printer.Font.Bold = True
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Print ""
    Printer.FontBold = False
    strMensaje = "" & rs("DES_LEGAL")
    Printer.Font.Size = 7
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Print ""
    Printer.Font.Size = 6
    strMensaje = Space(2) & rs("COD_USUARIO") & Space(8) & rs("FCH_REGISTRA") & Space(9) & "BTL" & rs("COD_LOCAL_EMISOR")
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Print ""
    strMensaje = Space(22) & rs("NUM_TICKET")
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    Printer.Print
    Printer.EndDoc
    MsgBox "Se ha generado el Cup?n N? " & "" & rs("NUM_DOCUMENTO_PAGO2"), vbInformation, App.ProductName
End Function



Public Sub subNuevo()
    gblNumeroTelefonico = ""

    objVenta.CodMedico = ""
    objVenta.vExisteDNI_RENIEC = ""
    objVenta.EsVentaMonedero = False
    objVenta.NumeroTarjetaMonedero = ""
    objVenta.PuntosTarjetaMonedero = 0
    
    objMedico.vCodMedico = ""
    
    ctlCliente1.Limpiar
    If txtDireccion.Visible = True Then txtDireccion.Text = ""
    
    frmPedido.lblMedico.Caption = ""
    frmPedido.txtCMPBus.Text = ""
    frmPedido.lblTotal.Caption = "0.00": frmPedido.lblTotalPagar.Caption = "0.00":
    frmPedido.lblRedondeo.Caption = "0.00": frmPedido.lblPagado.Caption = "0.00":
    frmPedido.lblVuelto.Caption = "0.00"
    frmPedido.lblPctCopago.Caption = "0.00"
    frmPedido.LblCoPago.Caption = "0.00"
    frmPedido.pstrdxPrcCero = ""
    frmPedido.PintaIndicadores
    '-- Limpia Variables Publicas --'
    frmPedido.pstrCant = "": frmPedido.pstrPrecio = "": frmPedido.pstrProd = "": frmPedido.pstrSubTot = ""
    frmPedido.pstrCtdFracc = "": frmPedido.pstrCtdFracc = ""
    frmPedido.pstrPctUnit = "": frmPedido.pstrPrcUniKairo = ""
    frmPedido.pstrImpuesto = "": frmPedido.pstrComision = ""
    frmPedido.pstrPromocion = ""
    frmPedido.psub_BeginArry
            
    frm_VTA_Busqueda.grdProductos.Limpiar
    frm_VTA_Busqueda.grdAlternativos.Limpiar
    frm_VTA_Busqueda.grdComplementarios.Limpiar
    'modificado por pherrera 06/11/07 se caia cuando se limpiaba y la pantalla de busqueda no esta cargada
    If frm_VTA_Busqueda.Visible Then frm_VTA_Busqueda.txtBuscar.selection
    'I.ECASTILLO 17.12.2020 | 06.01.2021
    Dim flg_ruteoA_cnv
    flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto
    End If
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "0" Then
    Else
        frm_VTA_PreviaTomaPedido.flgContinua = False
        mdiPrincipal.ctlCliente1.seleccionManualLocal = False
        objVenta.respetaLocal = False
    End If
cnvNoRuteaAuto:
    'F.ECASTILLO 17.12.2020
On Error GoTo g
   Dim i As Integer
   Dim intCountForm As Integer
   intCountForm = Forms.Count
   While i < intCountForm
        Dim LIMP As Boolean
        LIMP = True
        If Forms(i).name = "frmPedido" Then LIMP = False
        If Forms(i).name = "mdiPrincipal" Then LIMP = False
        If LIMP = True Then Unload (Forms(i)): i = 0: intCountForm = intCountForm - 1
        'Debug.Print "-->" & i & "<--||" & intCountForm & "||*****" & Forms(i).name & "******"
        i = i + 1
   Wend
g:
    'Debug.Print "hito"
    
    
''''''''''''    '*********************************************************************'
''''''''''''    '*********************************************************************'
''''''''''''    ' Recetario Magistral '
''''''''''''
''''''''''''    frm_VTA_RecetarioM.pxdbInsumos.ReDim 0, -1, 0, 13
''''''''''''    'frm_VTA_RecetarioM.GrdInsumos.Rebind
''''''''''''
''''''''''''    frm_VTA_RecetarioM.pstrDatoProv = "": frm_VTA_RecetarioM.pstrDatoCliente = "": frm_VTA_RecetarioM.pstrDatoMedico = ""
''''''''''''    frm_VTA_RecetarioM.TxtMedico.Text = "": frm_VTA_RecetarioM.TxtProveedor.Text = "": frm_VTA_RecetarioM.txtCliente.Text = ""
''''''''''''    frm_VTA_RecetarioM.LblCliente.Caption = "": frm_VTA_RecetarioM.LblMedico.Caption = "": frm_VTA_RecetarioM.LblNomprov.Caption = ""
''''''''''''    'frm_VTA_RecetarioM.ctlCboTipCliente.BoundText = "*"
''''''''''''    frm_VTA_RecetarioM.GrdInsumos.Close
''''''''''''
''''''''''''    frm_VTA_RecetarioM.pstrIdCantidad = "": frm_VTA_RecetarioM.pstrIdPctDsc = ""
''''''''''''    frm_VTA_RecetarioM.pstrIdPreVta = "": frm_VTA_RecetarioM.pstrIdProducto = ""
''''''''''''    frm_VTA_RecetarioM.pstrIdProductoBtl = "": frm_VTA_RecetarioM.pstrDatoCliente = ""
''''''''''''    frm_VTA_RecetarioM.pstrDatoMedico = "": frm_VTA_RecetarioM.pstrDatoProv = ""
''''''''''''    frm_VTA_RecetarioM.pstrRucProv = ""
''''''''''''    '*********************************************************************'
''''''''''''    '*********************************************************************'
''''''''''''
''''''''''''    frm_VTA_FacServPrestados.pstrIDKeito = ""
''''''''''''    frm_VTA_RecetarioM.pstrFlgRM = ""
''''''''''''
''''''''''''    frm_VTA_Servicios.grdServicios.Array1 = objVenta.Servicios
''''''''''''    frm_VTA_Servicios.grdServicios.Rebind
    
    txtDLVTelefono.Text = ""
    
    frm_VTA_Busqueda.txtBuscar.Clear
    frm_VTA_Busqueda.SetFocus
    
    
    'frm_VTA_FormaPago.pstrFPago = ""
    
'''    objDocumento.CargaDatosPN "", "", _
'''                                "", "", _
'''                                "", ""
'''
'''    objDocumento.CargaDatosPJ "", "", _
'''                                "", "", _
'''                                ""
    
    'frm_VTA_FormaPago.GrdListaFP.Delete
    'frm_VTA_FormaPago.GrdListaFP.Rebind
    
End Sub
    
Private Sub cmdHistorial_Click()

On Error GoTo handle

    With frm_DLV_HistorialCliente
        If ctlCliente1.LocalAsignado = "" Then MsgBox "No tiene local asignado", vbCritical, App.ProductName: Exit Sub
        '.Mostrar objUsuario.CodigoEmpresa, ctlCliente1.LocalAsignado, ctlCliente1.Codigo
        '.Mostrar objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, ctlCliente1.Codigo
        .Mostrar objUsuario.CodigoEmpresa, ctlCliente1.LocalAsignado, ctlCliente1.Codigo
        
    End With
    Exit Sub
handle:
MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub cmdImpresion_Click()
   frmColaImpresion.Show vbModal
End Sub

Public Sub cmdMantenimientos_Click()
    frm_VTA_Mantenimientos.Padre = cmdMantenimientos.Tag
    frm_VTA_Mantenimientos.Show vbModal
End Sub

Private Sub cmdModalidad_Click()
    On Error GoTo CtrlErr
    frm_VTA_Modalidad.Show vbModal
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub cmdProforma_Click()
On Error GoTo CtrlErr
'    frmPedido.grdPedido.Rebind
'    objVenta.LimpiaServicio
'    objVenta.LimpiaConvenio
'    frmPedido.grdPedido.Rebind
'    objVenta.ptmModalidad = Cotizaciones
'    objVenta.CodigoTipoVenta = Cotizaciones
'    objVenta.PctBeneficiario = 0
'    frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
'    frmPedido.Label4.Visible = False
'    frmPedido.lblPctCopago.Visible = False
'    frmPedido.Label8.Visible = False
'    frm_VTA_RecetarioM.pstrFlgRM = ""
'    mdiPrincipal.cmdGrabaVenta.Enabled = True
'    frmPedido.lblcopago.Visible = False
'    frm_DLV_Pedido.blnActivaPed = False
    frm_VTA_Cotizacion.Show vbModal
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdPuntos_Click()
On Error GoTo h
    frm_VTA_EscaneaLogin.Show vbModal
    If frm_VTA_EscaneaLogin.ok = True Then
        frm_VTA_Puntos.Show vbModal
    End If
'
Exit Sub
h:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdSalir_Click()
    '271107 comentado por PHERRERA se queda colgada la app
    'psub_Cerrar
    Unload Me
End Sub

'Private Sub Command1_Click()
' On Error GoTo CtrlErr
'    frm_ADM_CierreDiario.Show vbModal
'    Exit Sub
'CtrlErr:
'    MsgBox Err.Description, vbCritical, App.ProductName
'End Sub

Private Sub MDIForm_Load()
'    Dim objProceso As New cls_OFF_Procesos
'    objProceso.ProcesaUsuario objUsuario.CodigoLocal
'    Set objProceso = Nothing
    'agregado por pherrer para mensajes de DLV
    On Error Resume Next
    'DESCOMENTAR EN LA COMPILACION
    Set wskPrincipal = New CSocketMaster
        
    With wskPrincipal
        .Protocol = sckUDPProtocol
        .LocalPort = pstrPuerto
        .Bind
    End With
    '/////////////////////////////////////
On Error GoTo handle
    pstrPaginaComunicandonos = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_RUTA_PAG_COMUNICANDONOS")
    'AGREGADO DE VALIDACION TIPO MAQUINA RUTEO
    'VALIDAR SI MAQUINA RUTEA VISUALIZA FRM MODALIDAD Y PEDIDO
    If objUsuario.TipoMaquina = objUsuario.TipoMaquinaVerif Or objUsuario.TipoMaquina = objUsuario.TipoMaquinaRuteo Then Exit Sub
    frmPedido.Show
    frm_VTA_Modalidad.Show vbModal
    Exit Sub
handle:
    'If Err.Number <> 401 Then
    MsgBox Err.Description, vbCritical, App.ProductName
     '   cmdBusqueda_Click
End Sub

Sub HabilitaPermisos()
On Error GoTo handle
    Dim j As Integer
    
    While j < Me.Controls.Count
        Select Case TypeName(Me.Controls(j))
        Case "CommandButton"
            If Not Me.Controls(j).Tag = "" Then Me.Controls(j).Enabled = False
        End Select
        j = j + 1
    Wend
    
    j = 0
    
    Dim objPermisos As New clsAutorizacion
    Dim rsPermisos As oraDynaset
        
        Set rsPermisos = objPermisos.ListaPermisos(objUsuario.Aplicacion, objUsuario.Codigo)
        
        While Not rsPermisos.EOF
            If Me.name = rsPermisos("DES_URL").Value Then
                Me.Controls(rsPermisos("DES_ICONO").Value).Enabled = True
                Me.Controls(rsPermisos("DES_ICONO").Value).Tag = rsPermisos("COD_MENU").Value
            End If
            rsPermisos.MoveNext
        Wend
    
    Set objPermisos = Nothing

    Exit Sub
handle:
    Set objPermisos = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    TT1.Destroy
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo handle
    
    wskPrincipal.CloseSck
    
    Set objVenta = Nothing
    Set objUsuario = Nothing
    
    Dim frmX As Form
    
    gclsOracle.Cerrar
    
    For Each frmX In Forms
        If frmX.name <> Me.name Then
            Unload frmX
            Set frmX = Nothing
        End If
    Next
    
    'gclsOracle.Cerrar
    'Unload frmPedido

    Exit Sub
handle:
    Set objVenta = Nothing
    Set objUsuario = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub picComandos_Click()
'MsgBox objVenta.PctBeneficiarioReal
End Sub

Private Sub picComandos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub picComandos_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    TT1.Destroy
End Sub

Private Sub Picture1_Resize()
On Error GoTo handle
    'check size
    If Picture1.Width < 120 Then
        Picture1.Width = 120
    End If
    'if Form1 is docked, position it so its resizing border is hidden
    'outside the confines of Picture1.
    If frmPedido.bDocked Then
        frmPedido.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX), Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo handle
    'As a simple alternative to using a splitter control we can resize
    'Picture1 by sending a message that will make windows draw a
    'resizing border for us.
    If Picture1.Visible Then
        ReleaseCapture 'need to do this or SendMessage fails
        'Send message to start resizing picture1
        SendMessage Picture1.hwnd, WM_NCLBUTTONDOWN, HTRIGHT, ByVal &O0
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtDLVTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
    'psub_KeyDownAplicacion KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        If Trim(txtDLVTelefono.Text) = "" Then MsgBox "Debe ingresar el n?mero de tel?fono.", vbOKOnly + vbExclamation, "Mensaje": txtDLVTelefono.SetFocus: Exit Sub
        cmdDLVBuscar_Click
    End If
End Sub

Private Sub Salida()
    
'    frm_VTA_Documento.blnTipoDoc = False
'    subNuevo
'
'
'    If objMaquina.Valida = False Then
'            End
'    Else
'            If objUsuario.CodigoLocal = objUsuario.LocalDelivery And objMaquina.Delivery = "1" Then
'                objUsuario.CodTipoVenta = objUsuario.TipoVentaDlv
'            ElseIf objUsuario.CodigoLocal <> objUsuario.LocalDelivery And objMaquina.Delivery = "0" Then
'                objUsuario.CodTipoVenta = objUsuario.TipoVentaLocal
'            Else
'                objUsuario.CodTipoVenta = objUsuario.TipoVentaWeb
'            End If
'    End If
'    objUsuario.TipoCambio = objMoneda.TCambio(objUsuario.TipoCambioDefault, objUsuario.TipoCambioMonedaDefault)
'    cmdModalidad_Click
    
    Cancelar
    
End Sub

Private Sub Cancelar()
    On Error GoTo handle
    
    gstrCodTarjetaMon = ""
    gstrCodTarjetaFid = ""
    gintFidelizado = 0
    frmPedido.lbl_Cliente.Caption = ""     'nuevo
    frmPedido.lblTotalDescuento.Caption = "0.00"
    frmPedido.pstrDniCli = ""
    frmPedido.pstrNomcli = ""
    frmPedido_Busca_Cli.v_delfrm = ""
    frm_VTA_Documento.blnTipoDoc = False
    frmPedido.pstrCodCliente_Ink = ""
    frmPedido.pstrPuntos_Ink = ""
    frmPedido.lblPuntosAcum = "0"
    subNuevo
    
    Unload frm_VTA_RecetarioM
        
    'SE PUSO ESTA VALIDACION PARA ESCOGER EL TIPO DE MAQUINA
    'SE HIZO EL CAMBIO EL 22/11 POR JLOPEZ
        
    If objUsuario.flgDeliveryProv = "1" Then
        frm_VTA_TipoMaquina.Show vbModal
        If objUsuario.Conectado Then
            tipoPantalla
        End If
    End If
        
    If objMaquina.valida = False Then
        End
    Else
        If objUsuario.CodigoLocal = objUsuario.LocalDelivery And objMaquina.Delivery = "1" Then
            objUsuario.CodTipoVenta = objUsuario.TipoVentaDlv
        ElseIf objUsuario.CodigoLocal <> objUsuario.LocalDelivery And objMaquina.Delivery = "0" Then
            objUsuario.CodTipoVenta = objUsuario.TipoVentaLocal
        Else
            objUsuario.CodTipoVenta = objUsuario.TipoVentaWeb
        End If
    End If
    objUsuario.TipoCambio = objMoneda.TCambio(objUsuario.TipoCambioDefault, objUsuario.TipoCambioMonedaDefault)
        
    If Not strNombreFomularioOrigen = "frm_DLV_Pedido" Then
        frm_VTA_Modalidad.Show vbModal
    Else
        strNombreFomularioOrigen = ""
        '--- inicializo las variables como si fuera modalidad venta regular
        '--- Jahzeel Lopez 13/06/2008
        '---
            
        objVenta.ptmModalidad = Venta_Regular
        ptmTipoPrecio = Regular
        objVenta.CodigoTipoVenta = Venta_Regular
        gblNumeroTelefonico = ""
    End If
    'I.ECASTILLO 17.12.2020
    objVenta.bk_codLocal = ""
    objVenta.bk_codLocalCapacidad = ""
    objVenta.bk_FechaCapacidad = ""
    objVenta.bk_HoraCapacidad = ""
    objVenta.bk_HoraCapacidad2 = ""
    objVenta.bk_codBeneficiario = ""
    objVenta.bk_strProforma = ""
    objVenta.bk_ServiceType = ""
    objVenta.bk_chkRET = ""
    objVenta.bk_flgPactado = ""
    objVenta.bk_deliveryTime = ""
    objVenta.bk_message = ""
    objVenta.bk_segmento = ""
    objVenta.bk_amount = ""
    objVenta.bk_starHour = ""
    objVenta.bk_endHour = ""

    objVenta.dc_street = ""
    objVenta.dc_number = ""
    objVenta.dc_apartment = ""
    objVenta.dc_country = ""
    objVenta.dc_city = ""
    objVenta.dc_district = ""
    objVenta.dc_latitude = ""
    objVenta.dc_longitude = ""
    objVenta.bk_codCliente = ""
    objVenta.bk_Ubigeo = ""
    objVenta.dc_referencia = ""
    objVenta.dc_urbanizacion = ""
    printCantCallsGMaps "", "Cancela"
    'F.ECASTILLO 17.12.2020
    objVenta.busquedaNVeces = 0
    objVenta.strBusqueda = ""
    frmPedido.cancelaOptions
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

'Recibiendo data por el winsock
Private Sub wskPrincipal_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Dim Texto As String
Dim Clave As String
Dim SubClave As String
Dim CadenaLimpia As String
On Error GoTo Control

wskPrincipal.GetData data
data = Hex2Str(data)

Texto = strDecrypt(data)
'MsgBox (Texto)
Clave = Mid(Texto, 1, 4)
SubClave = Mid(Texto, 5, 4)
CadenaLimpia = Mid(Texto, 9, 500)

Select Case Clave
Case "MENS"
    Select Case SubClave
        Case "DELI"
            TT1.BackColor = &HFFDE79
            TT1.Style = TTBalloon
            TT1.Icon = TTIconInfo
            TT1.Title = "Delivery"
            TT1.TipText = CadenaLimpia
            TT1.PopupOnDemand = True
            TT1.VisibleTime = 600000
            TT1.CreateToolTip txtToolTip.hwnd
            TT1.Show 0, txtToolTip.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
        Case "MENS"
            TT1.Style = TTBalloon
            TT1.Icon = TTIconInfo
            TT1.Title = "Mensaje"
            TT1.TipText = CadenaLimpia
            TT1.PopupOnDemand = True
            TT1.VisibleTime = 6000
            TT1.CreateToolTip txtToolTip.hwnd
            TT1.Show 0, txtToolTip.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
        Case "ADVE"
            TT1.Style = TTBalloon
            TT1.Icon = TTIconWarning
            TT1.Title = "Advertencia"
            TT1.TipText = CadenaLimpia
            TT1.PopupOnDemand = True
            TT1.VisibleTime = 6000
            TT1.CreateToolTip txtToolTip.hwnd
            TT1.Show 0, txtToolTip.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
    End Select
End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

'esta codigo estaba debajo del FORM_LOAD (pherrera 24/02/09)
'Sub tipoPantalla()
'    pctDelivery.Visible = False
'    ctlCliente1.Modo = False
'    Caption = gstrAplicacion & " * Ver: " & gstrVersion & " [" & gvarUSUARIO & "@" & gvarTNSNAME & "] " & "Local " & objUsuario.CodigoLocal & " de la Empresa " & objUsuario.Empresa
'    Select Case objUsuario.TipoMaquina
'        Case "001" 'cuando es la maquina del administrador
'        Case "002" 'cuando es la maquina de un cajero
'        Case "003" 'cuando es una cabina
'            pctDelivery.Visible = True
'            txtDLVTelefono.SetFocus
'            ctlCliente1.Modo = True
'        Case "004" ' cuando es el modulo de ruteo
'            frm_DLV_Seguimiento.Show
'        Case Else
'            MsgBox "No se ha definido el tipo de maquina", vbCritical, App.ProductName
'            End
'    End Select
'
'    'frm_DLV_Seguimiento
'
'    If objUsuario.TipoMaquina = "003" Then
'    Else
'    End If
'End Sub


Function fn_ImprimeTicketTermico(ByVal strSecVenta As String)

    Dim rs As oraDynaset
    Dim intVal As Integer
    Dim arr() As String
    Dim i As Integer
    
   
         Set rs = objDocumento.ListaTermico(strSecVenta, objUsuario.CodigoLocal)
    
        intVal = 0
        
        While Not rs.EOF
              arr = Split(CStr(rs("CONCATENA")), "|")
              imgLogoBtl.Picture = Me.ImageList1.ListImages(1).Picture
              Printer.FontName = "Courier New"
              Printer.PaintPicture imgLogoBtl, 1000, 20, 2150, 950
              Printer.FontSize = 10
              Printer.Print ""
              Printer.Print ""
              Printer.Print ""
              Printer.Print ""
              Printer.Print ""
              Printer.FontSize = 12
              Printer.FontBold = True
              Printer.Print Space(8) & "VENTA CONVENIO"
              Printer.Print ""
              Printer.FontBold = False
              Printer.FontSize = 8
              Printer.Print Space(2) & "Terminal  : " & objUsuario.NombrePC
              Printer.Print Space(2) & "Vendedor  : " & rs("VENDEDOR")
              Printer.Print Space(2) & "Convenio  : " & rs("CONVENIO")
              Printer.Print Space(2) & "F Emisi?n : " & rs("EMISION")
              Printer.Print ""
              Printer.FontSize = 10
              Printer.FontName = "IDAutomation.com Code39"  'LETRA DE CODIGO DE BARRA'
              Printer.Print Space(5) & "*" & objUsuario.CodigoLocal & "-" & strSecVenta & "*"
              Printer.FontName = "Courier New"
              Printer.FontSize = 11
              Printer.FontBold = True
              Printer.Print ""
              For i = 0 To UBound(arr)
                  Printer.Print Space(2) & " " & Trim(arr(i))
              Next
              Printer.Print ""
              Printer.FontBold = False
              Printer.FontSize = 9
              Printer.FontName = "Courier New"
              Printer.Print ""
              Printer.Print Space(2) & "Firma :_______________________"
              Printer.Print Space(2) & "Nombre: " & rs("NOMBRE")
              Printer.Print Space(2) & "DNI   : " & rs("DNI")
              Printer.Print Space(2) & "Telef : " & rs("TELEFONO")
              Printer.FontBold = True
              Printer.FontSize = 8
              Printer.Print ""
              Printer.FontName = "Courier New"
              Printer.FontSize = 10
              Printer.Print ""
              Printer.Print "*** Documento sin valor Fiscal ***"
                
              Printer.FontBold = False
              Printer.EndDoc
              rs.MoveNext
        Wend
    
 
    
    
Termina:
        Exit Function
End Function
'I.ECASTILLO 27.10.2020
Function reservaCapacidad(ByVal strProforma As String) As String
    Dim resp As String
    Dim Fecha As String
    Dim hora As String
'    Fecha = frm_VTA_Documento.DTPicker1.Value
'    Fecha = CStr(Format(Fecha, "yyyy-mm-dd"))
'    hora = frm_VTA_Documento.DTPicker3.Value
'    hora = CStr(Format(hora, "hh:mm:ss"))
    objVenta.bk_FechaCapacidad = CStr(Format(objVenta.bk_FechaCapacidad, "yyyy-mm-dd"))
    objVenta.bk_HoraCapacidad = CStr(Format(objVenta.bk_HoraCapacidad, "hh:mm:ss"))
    objVenta.bk_HoraCapacidad2 = CStr(Format(objVenta.bk_HoraCapacidad2, "hh:mm:ss"))
    resp = objProforma.reservaCapacidad(objVenta.bk_codLocalCapacidad, _
                                        objVenta.bk_FechaCapacidad, _
                                        objVenta.bk_HoraCapacidad, _
                                        objVenta.bk_ServiceType, _
                                        strProforma)
    reservaCapacidad = resp
End Function
Function retornaCapacidad() As String
    Dim resp As String
    Dim Fecha As String
    Dim hora As String
    Fecha = frm_VTA_Documento.DTPicker1.Value
    Fecha = CStr(Format(Fecha, "yyyy-mm-dd"))
    hora = frm_VTA_Documento.DTPicker3.Value
    hora = CStr(Format(hora, "hh:mm:ss"))
    resp = objProforma.retornaCapacidad(frm_VTA_ListaCapacidades.strLocal, _
                                        Fecha, _
                                        hora)
    retornaCapacidad = resp
End Function
Function asignaDigital(ByVal NumProforma As String) As String
    Dim resp As String
    asignaDigital = "0"
    resp = objProforma.asignaDigital(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, NumProforma)
    asignaDigital = resp
End Function
'F.ECASTILLO 27.10.2020
Function reValidaFracc() As String
    Dim valFracDcCap As String
    Dim response As String
'    response = "1"
'    Dim rs2 As oraDynaset
'    Set rs2 = objProducto.ListaFraccionamientoSugeridoV2(lblCodigo.Caption, mdiPrincipal.ctlCliente1.LocalDespacho)
'    'unidDcCap = "" & rs2("UNID_VTA").Value
'    valFracDcCap = "" & rs2("VAL_FRAC_LOCAL").Value
'    If Len(Trim(valFracDcCap)) = 0 Then chkFraccionamiento.Enabled = False: strFracciona = 0
'
'    reValidaFracc = response
End Function


Public Function printCantCallsGMaps(Optional prof As String, Optional msg As String)
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim errors As New XArrayDB
    Dim i As Integer
    Dim aux As Integer
    Dim objWS As New clsWebService
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    errors.ReDim 0, -1, 0, 9
    i = 0
    
    aux = garrCallGoogleMaps.Count(1)
    While i < aux
        params.AppendRows
        values.AppendRows
        params(i, 0) = garrCallGoogleMaps(i, 0)
        values(i, 0) = garrCallGoogleMaps(i, 1)
        i = i + 1
    Wend
    If Len(Trim(prof)) > 0 Then
        aux = garrCallGoogleMaps.Count(1)
        params.AppendRows
        values.AppendRows
        params(aux, 0) = "Proforma"
        values(aux, 0) = prof
    End If
    If Len(Trim(msg)) > 0 Then
        aux = garrCallGoogleMaps.Count(1)
        params.AppendRows
        values.AppendRows
        params(aux, 0) = "Mensaje"
        values(aux, 0) = msg
    End If
    objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors
    Set garrCallGoogleMaps = Nothing
    garrCallGoogleMaps.ReDim 0, -1, 0, 10
End Function

