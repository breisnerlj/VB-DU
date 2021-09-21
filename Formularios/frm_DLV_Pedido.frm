VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_DLV_Pedido 
   BorderStyle     =   0  'None
   Caption         =   "Pedido Delivery"
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   Icon            =   "frm_DLV_Pedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      Begin TabDlg.SSTab SSTab1 
         Height          =   6375
         Left            =   0
         TabIndex        =   2
         Top             =   720
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   11245
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Por Despachar"
         TabPicture(0)   =   "frm_DLV_Pedido.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "grdTransferencias"
         Tab(0).Control(1)=   "grdPedidoDLV"
         Tab(0).Control(2)=   "Label3"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Por Anular"
         TabPicture(1)   =   "frm_DLV_Pedido.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdAnulacion"
         Tab(1).Control(1)=   "CmdAnulacionDoc"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Re - impresión"
         TabPicture(2)   =   "frm_DLV_Pedido.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtDesDireccion"
         Tab(2).Control(1)=   "txtNumProforma"
         Tab(2).Control(2)=   "grdProformado"
         Tab(2).Control(3)=   "dtpFchIni"
         Tab(2).Control(4)=   "dtpFchFin"
         Tab(2).Control(5)=   "Label7"
         Tab(2).Control(6)=   "Label6"
         Tab(2).Control(7)=   "Label1"
         Tab(2).Control(8)=   "Label2"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Provincia"
         TabPicture(3)   =   "frm_DLV_Pedido.frx":035E
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label4"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Frame3"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).ControlCount=   2
         Begin vbp_Ventas.ctlTextBox txtDesDireccion 
            Height          =   315
            Left            =   -73560
            TabIndex        =   9
            Top             =   1560
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            Tipo            =   2
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
         Begin vbp_Ventas.ctlTextBox txtNumProforma 
            Height          =   315
            Left            =   -73560
            TabIndex        =   8
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Tipo            =   7
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
         Begin vbp_Ventas.ctlGrilla grdTransferencias 
            Height          =   2415
            Left            =   -74940
            TabIndex        =   13
            Top             =   3840
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4260
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
         Begin VB.CommandButton CmdAnulacionDoc 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Anulación"
            Height          =   555
            Left            =   -69600
            Picture         =   "frm_DLV_Pedido.frx":037A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anulación de documento"
            Top             =   5340
            Width           =   1095
         End
         Begin vbp_Ventas.ctlGrilla grdPedidoDLV 
            Height          =   3015
            Left            =   -74940
            TabIndex        =   3
            Top             =   480
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5318
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
         Begin vbp_Ventas.ctlGrilla grdAnulacion 
            Height          =   4695
            Left            =   -74880
            TabIndex        =   4
            Top             =   480
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   8281
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
         Begin vbp_Ventas.ctlGrilla grdProformado 
            Height          =   4155
            Left            =   -74880
            TabIndex        =   11
            Top             =   2040
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   7329
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
         Begin MSComCtl2.DTPicker dtpFchIni 
            Height          =   315
            Left            =   -73560
            TabIndex        =   6
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   82182145
            CurrentDate     =   39345
         End
         Begin MSComCtl2.DTPicker dtpFchFin 
            Height          =   315
            Left            =   -73560
            TabIndex        =   7
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   82182145
            CurrentDate     =   39345
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5715
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   6795
            Begin VB.CommandButton cmdAnular 
               Caption         =   "Anular"
               Height          =   705
               Left            =   3438
               Picture         =   "frm_DLV_Pedido.frx":0904
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   4920
               Width           =   1020
            End
            Begin VB.CommandButton cmdDetalle 
               Caption         =   "Asignar"
               Height          =   705
               Left            =   45
               Picture         =   "frm_DLV_Pedido.frx":0E8E
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   4920
               Width           =   1020
            End
            Begin VB.CommandButton cmdLlevando 
               Caption         =   "Llevando"
               Height          =   705
               Left            =   1176
               Picture         =   "frm_DLV_Pedido.frx":1418
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   4920
               Width           =   1020
            End
            Begin VB.CommandButton cmdLlegadaDestino 
               Caption         =   "Llega Destino"
               Height          =   705
               Left            =   4590
               Picture         =   "frm_DLV_Pedido.frx":19A2
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   4920
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.CommandButton cmdEntregado 
               Caption         =   "Entregado"
               Height          =   705
               Left            =   5715
               Picture         =   "frm_DLV_Pedido.frx":1F2C
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   4920
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.CommandButton cmdLlegadaLocal 
               Caption         =   "Llegada a local"
               Height          =   705
               Left            =   2307
               Picture         =   "frm_DLV_Pedido.frx":24B6
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   4920
               Width           =   1020
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Option1"
               Height          =   255
               Left            =   3120
               TabIndex        =   20
               Top             =   0
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Option2"
               Height          =   255
               Left            =   4560
               TabIndex        =   19
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Frame Frame2 
               Caption         =   "Tiempos"
               Height          =   1155
               Left            =   0
               TabIndex        =   17
               Top             =   3360
               Width           =   6735
               Begin vbp_Ventas.ctlGrilla grdTiempo 
                  Height          =   795
                  Left            =   120
                  TabIndex        =   18
                  Top             =   240
                  Width           =   6555
                  _ExtentX        =   11562
                  _ExtentY        =   1402
                  MenuPopUp       =   0   'False
                  Resalte         =   0   'False
               End
            End
            Begin MSComctlLib.ImageList ImageList2 
               Left            =   2040
               Top             =   1920
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   5
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frm_DLV_Pedido.frx":2A40
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frm_DLV_Pedido.frx":2E6B
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frm_DLV_Pedido.frx":32A1
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frm_DLV_Pedido.frx":3689
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frm_DLV_Pedido.frx":3A5E
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin vbp_Ventas.ctlGrilla grdPedidos 
               Height          =   2895
               Left            =   0
               TabIndex        =   26
               Top             =   360
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   5106
               MenuPopUp       =   0   'False
               Resalte         =   0   'False
            End
            Begin VB.CommandButton cmdAvisado 
               Caption         =   "&Avisado"
               Height          =   650
               Left            =   240
               Picture         =   "frm_DLV_Pedido.frx":3E45
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   3720
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Pedidos pendientes de Ruteo"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   120
               Width           =   2100
            End
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   31
            Top             =   1620
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nº Proforma:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   30
            Top             =   1260
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Opción habilitada SOLO para locales de Provincia."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   660
            Width           =   3345
         End
         Begin VB.Label Label3 
            Caption         =   "Transferencias"
            Height          =   255
            Left            =   -74940
            TabIndex        =   14
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   12
            Top             =   540
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   10
            Top             =   900
            Width           =   465
         End
      End
      Begin vbp_Ventas.ctlToolBar ToolBar_DLV_Pedido 
         Height          =   600
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1058
         ModoBotones     =   7
      End
   End
End
Attribute VB_Name = "frm_DLV_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProforma As New clsProforma
Dim objDocumento As New clsDocumento
Dim objlocal As New clsLocal
Dim objPedido As New clsProforma
'Dim odynTransf As oraDynaset

Public odynPedido As oraDynaset
Public blnActivaPed As Boolean

Dim strFlgImprimeDoc As String
Dim UbicaPrinter As String
Dim Impresoras As Printer
Dim Devicename As String
Dim bolEncontroImpAlm As Boolean

'Private Sub Command1_Click()
'BuscaPorDespachar
'End Sub

Private Sub CmdAnular_Click()
Dim strMensaje As String
Dim Bookmark As Variant
On Error GoTo Control

    If Not cmdAnular.Enabled Then Exit Sub

    If grdPedidos.ApproxCount = 0 Then Exit Sub
    If MsgBox("Se va anular la proforma " & grdPedidos.Columns("NUM_PROFORMA") & " , desea continuar ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        strMensaje = objPedido.Anula(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objUsuario.CODIGO)
        If strMensaje <> "" Then
            MsgBox strMensaje, vbCritical, App.ProductName
            Bookmark = grdPedidos.Bookmark
            grdPedidos.SetFocus
        End If
        Bookmark = grdPedidos.Bookmark
        ListaPedido
        grdPedidos.Bookmark = Bookmark
    End If

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Activate()
    blnActivaPed = True
End Sub

Private Sub Form_Load()
    
    setteaFormulario Me
    SeteaGrila
    SeteaGrillaAnu
    SeteaGrillaReImp
    SeteaGrillaTransf
    Frame3.Visible = False
    grdPedidoDLV.Columns(3).FetchStyle = True
    If objUsuario.flgDeliveryProv = "1" Then
        ListaPedido
    End If
    
    SSTab1.Tab = 0
    
    dtpFchIni.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    dtpFchFin.Value = Format(dtpFchIni.Value, "dd/mm/yyyy") 'Format(objUsuario.sysdate, "dd/mm/yyyy")
    
    BuscaPorDespachar
    
    
End Sub



Private Sub grdPedidoDLV_DblClick()
    frm_VTA_Cotizacion.strNombreFomularioOrigen = Me.name
    If grdPedidoDLV.ApproxCount = 0 Then Exit Sub
    objVenta.LimpiaProductos
    frm_VTA_Cotizacion.pstrNumProf = grdPedidoDLV.DataSource("NUM_PROFORMA").Value
    
    If Mid(Trim(grdPedidoDLV.DataSource("COD_MODALIDAD_VENTA").Value), 1, 3) = "TRA" Then
        frm_VTA_Cotizacion.TipoDoc = True
    Else
        frm_VTA_Cotizacion.TipoDoc = False
    End If
    frm_VTA_Cotizacion.Show
End Sub

Private Sub grdPedidoDLV_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
CellStyle.Font.Bold = True
End Sub

Private Sub grdPedidoDLV_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
    If grdPedidoDLV.Columns(6).CellText(Bookmark) = "TRA" Then
        RowStyle.BackColor = &H80FF80
    End If
End Sub

Private Sub grdPedidoDLV_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdPedidoDLV_DblClick
    End Select
End Sub

Private Sub grdPedidoDLV_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdPedidoDLV.ApproxCount <= 0 Then Exit Sub
    Set grdTransferencias.DataSource = objPedido.ListaTransfPedido(objUsuario.CodigoEmpresa, grdPedidoDLV.Columns("NUM_PROFORMA").Value, objUsuario.CodigoLocal)
End Sub



Private Sub ToolBar_DLV_Pedido_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
On Error GoTo Control

Dim CantImp As Integer

    Select Case Index
        Case 1
            If SSTab1.Tab = 0 Then
                '''If grdPedidoDLV.ApproxCount <= 0 Then Exit Sub
                BuscaPorDespachar
            ElseIf SSTab1.Tab = 1 Then
                AnulaDoc
            ElseIf SSTab1.Tab = 2 Then
                BuscaRe_Imp
            ElseIf SSTab1.Tab = 3 Then
                If objUsuario.flgDeliveryProv = "1" Then
                    ListaPedido
                End If
            End If
            
        Case 2
            If SSTab1.Tab = 0 Then
                '''''SeteaGrila
                BuscaPorDespachar
            ElseIf SSTab1.Tab = 1 Then
                AnulaDoc
            ElseIf SSTab1.Tab = 2 Then
                BuscaRe_Imp
            ElseIf SSTab1.Tab = 3 Then
                If objUsuario.flgDeliveryProv = "1" Then
                    ListaPedido
                End If
            End If
            
        Case 3
            Dim k As Integer
            If SSTab1.Tab = 0 Then 'normal
                If grdPedidoDLV.ApproxCount = 0 Then Exit Sub
                CantImp = objProforma.Cuenta_Impresoras_x_Maquina(objUsuario.CodigoEmpresa, objUsuario.NombrePC) 'mlaguna cantidad de impresoras por maquina 06/04/2010
                If grdTransferencias.ApproxCount <= 0 Then
                    MsgBox "Esta opción solo esta habilitada para los pedidos que tienen transferencia.", vbOKOnly + vbInformation, "Mensaje"
                    Exit Sub
                End If
                If grdPedidoDLV.Columns("NUM_PROFORMA").Value <> "" Then
                    If CantImp <= 0 Then
                            frm_SERV_Impresoras.Show vbModal
                            For k = 1 To frm_SERV_Impresoras.gNroCopia
                               MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                                      "en posición de Pedido Delivery", vbInformation, App.FileDescription
                                 objProforma.ImprimirDelivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdPedidoDLV.Columns("NUM_PROFORMA").Value, "1"
                            Next k
                    ElseIf CantImp > 0 Then
                            objProforma.ImprimirTicket_Delivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdPedidoDLV.Columns("NUM_PROFORMA").Value, "1", False  'Prueba
                            'objProforma.ImprimirTicket_DeliveryNew objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdPedidoDLV.Columns("NUM_PROFORMA").Value, "1", False  'Prueba
                    End If
                End If
                
            ElseIf SSTab1.Tab = 2 Then
                If grdProformado.ApproxCount = 0 Then Exit Sub
                CantImp = objProforma.Cuenta_Impresoras_x_Maquina(objUsuario.CodigoEmpresa, objUsuario.NombrePC) 'mlaguna cantidad de impresoras por maquina 06/04/2010
                
                If grdProformado.Columns("num_proforma").Value <> "" Then
                    
                    If CantImp <= 0 Then
                            frm_SERV_Impresoras.Show vbModal
                            k = 0
                            For k = 1 To frm_SERV_Impresoras.gNroCopia
                               MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                                      "en posición de Pedido Delivery", vbInformation, App.FileDescription
                                  objProforma.ImprimirDelivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdProformado.Columns("NUM_PROFORMA").Value, "1"
                            Next k
                    ElseIf CantImp > 0 Then
                            objProforma.ImprimirTicket_Delivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdProformado.Columns("NUM_PROFORMA").Value, "1", True
                            'objProforma.ImprimirTicket_DeliveryNew objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdProformado.Columns("NUM_PROFORMA").Value, "1", True
                    End If
                End If
               
                Dim i%
                For i = 0 To 1
                    If grdProformado.Columns("num_proforma").Value <> "" Then
                        UbicaPrinter = Printer.Devicename
                        
                        
                        Devicename = objlocal.NombreImpresoraAlmacen
                        bolEncontroImpAlm = False
                        For Each Impresoras In Printers
                            If UCase(Impresoras.Devicename) = Trim(UCase(Devicename)) Then
                                Set Printer = Impresoras
                                bolEncontroImpAlm = True
                                Exit For
                            End If
                        Next Impresoras

                        For Each Impresoras In Printers
                            If UbicaPrinter = Impresoras.Devicename Then Set Printer = Impresoras: Exit For
                        Next
                    End If
                Next i
            End If
            
        Case 4
            If SSTab1.Tab = 0 Then
                grdPedidoDLV.MostrarExcel
            ElseIf SSTab1.Tab = 1 Then
                grdAnulacion.MostrarExcel
            ElseIf SSTab1.Tab = 2 Then
                grdProformado.MostrarExcel
            End If
            
        Case 5
            If SSTab1.Tab = 0 Then
                grdPedidoDLV.MostrarEmail
            ElseIf SSTab1.Tab = 1 Then
                grdAnulacion.MostrarEmail
            ElseIf SSTab1.Tab = 2 Then
                grdProformado.MostrarEmail
            End If
            
        Case 6
            Unload Me
        Case Else
            MsgBox "Esta opción se encuentra Deshabilitada", vbExclamation, App.ProductName
    End Select

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub CmdAnulacionDoc_Click()
    On Error GoTo CtrlErr
    If grdAnulacion.ApproxCount = 0 Then
      Exit Sub
    End If

    If MsgBox("Desea anular la " + grdAnulacion.Columns("COD_TIPO_DOCUMENTO").Value + " Nº " + grdAnulacion.Columns("NUM_DOCUMENTO").Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        grdAnulacion.SetFocus
        Exit Sub
    End If
    
    Dim gvarError As String
    Dim ValorRet As String
    gvarError = objDocumento.Anula(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Trim(grdAnulacion.Columns("COD_TIPO_DOCUMENTO").Value), Replace(grdAnulacion.Columns("NUM_DOCUMENTO").Value, "-", ""), objUsuario.EstadoEmitido, objUsuario.CODIGO, ValorRet)
    
    If gvarError = "" Then
        MsgBox "Se anulo el siguiente documento:" + Chr(13) + ValorRet, vbInformation, App.ProductName
        AnulaDoc
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    grdAnulacion.SetFocus
End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub BuscaPorDespachar()
    Dim rsTmp As OracleInProcServer.oraDynaset
    'Set rsTmp = New OracleInProcServer.oraDynaset
    Set rsTmp = objProforma.ListaPedidoDLV(objUsuario.CodigoEmpresa, _
                                           objUsuario.CodigoLocal, _
                                           CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                           CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")), _
                                           objUsuario.PedidoAvisado)
    
    Set grdPedidoDLV.DataSource = rsTmp
    Set rsTmp = Nothing
                                                             
End Sub

Private Sub AnulaDoc()
        Set grdAnulacion.DataSource = objProforma.Documento_x_Anular(objUsuario.CodigoEmpresa, _
                                                                     objUsuario.CodigoLocal)
End Sub

Private Sub BuscaRe_Imp()
'''''    Set grdProformado.DataSource = objDelivery.ListaPedidoDLV(objUsuario.CodigoEmpresa, _
                                                              objUsuario.CodigoLocal, _
                                                              CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                              CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")), _
                                                              objUsuario.PedidoProformado)
                                                              
                                                              
                                                              
                                                              
    Set grdProformado.DataSource = objProforma.ListaPedidoDLVReimp(objUsuario.CodigoEmpresa, _
                                                              objUsuario.CodigoLocal, _
                                                              CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                              CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")), _
                                                              Trim(Replace(txtNumProforma.Text, "-", "")), _
                                                              Trim(txtDesDireccion.Text))
                                                              
End Sub

Sub SeteaGrila()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim Columna As TrueDBGrid70.Column
  
  
    arrCampos = Array("CIA", "NUM_PROFORMA", _
                      "FCH_REGISTRA", "MOTORIZADO", _
                      "DES_AUX_CLI_DIRECC", _
                      "MTO_TOTAL", "COD_TIPO_DOC", _
                      "COD_LOCAL_REF", "COD_CLIENTE", _
                      "NOMBRES", "COD_MODALIDAD_VENTA", _
                      "COD_ESTADO", "FCH_HORA_PACT_ENTR", _
                      "FCH_HORA_PACT_RECOG", "FLG_TRANSFERENCIA", _
                      "FLG_URGENTE", "FLG_ENTREGA_LOCAL", _
                      "FLG_RESERVA_STK", "MTO_BASE_IMP", _
                      "MTO_EXONERADO", "MTO_IMPUESTO", _
                      "MTO_VUELTO", "COD_CONVENIO", _
                      "DES_CLIENTE", "DES_AUX_CLI_NOMBRE")
                      
    arrCaption = Array("", "Proforma", _
                        "", "Motorizado", _
                        "Dirección Cliente", _
                        "Importe", "Tipo", _
                        "Local Refer.", "", _
                       "", "", _
                       "", "", _
                       "", "", _
                       "", "", _
                       "", "", _
                       "", "", _
                       "", "", _
                       "", "Nombre Cliente")
                       
    arrAncho = Array(0, 1000, _
                        0, 1800, _
                        2700, 600, _
                        600, 500, _
                        500, 0, _
                        0, 0, _
                        0, 0, _
                        0, 0, _
                        0, 0, _
                        0, 0, _
                        0, 0, _
                        0, 0, _
                        2000, 0)
                     
    arrAlineacion = Array(dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgRight, _
                          dbgRight, dbgLeft, _
                          dbgCenter, dbgCenter, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft)
                          
    grdPedidoDLV.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdPedidoDLV.Columns(0).Visible = False
    grdPedidoDLV.Columns(2).Visible = False
    'grdPedidoDLV.Columns(6).Visible = False
    grdPedidoDLV.Columns(7).Visible = False
    grdPedidoDLV.Columns(8).Visible = False
    grdPedidoDLV.Columns(9).Visible = False
    grdPedidoDLV.Columns(10).Visible = False
    grdPedidoDLV.Columns(11).Visible = False
    grdPedidoDLV.Columns(12).Visible = False
    grdPedidoDLV.Columns(13).Visible = False
    grdPedidoDLV.Columns(14).Visible = False
    grdPedidoDLV.Columns(15).Visible = False
    grdPedidoDLV.Columns(16).Visible = False
    grdPedidoDLV.Columns(17).Visible = False
    grdPedidoDLV.Columns(18).Visible = False
    grdPedidoDLV.Columns(19).Visible = False
    grdPedidoDLV.Columns(20).Visible = False
    grdPedidoDLV.Columns(21).Visible = False
    grdPedidoDLV.Columns(22).Visible = False
    grdPedidoDLV.Columns(23).Visible = False
    'grdPedidoDLV.Columns(24).Visible = False
    
    For Each Columna In grdPedidoDLV.Columns
        Columna.AllowSizing = True
    Next
    grdPedidoDLV.FetchRowStyle = True
End Sub

Sub SeteaGrillaTransf()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("NUM_PROFORMA", "LOCAL_ORIGEN", "DES_ESTADO_PEDIDO", "COD_PRODUCTO", "DESCRIPCION")
                      
    arrCaption = Array("Transferencia", "Local Origen", "Estado", "Codigo", "Producto")
                       
    arrAncho = Array(1400, 1400, 1000, 900, 2800)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
                          
    grdTransferencias.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub


Sub SeteaGrillaAnu()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("NUM_PROFORMA", "NUM_DOCUMENTO", "COD_TIPO_DOCUMENTO", _
                      "FCH_REGISTRA", "COD_USUARIO", _
                      "NOMBRES")
                      
    arrCaption = Array("Proforma", "Documento", "Tipo", _
                       "Fecha Emisión", "Codigo", _
                       "Usuario")
                       
    arrAncho = Array(1200, 1200, 1000, _
                     2000, 900, _
                     1800)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft)
                          
    grdAnulacion.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdAnulacion.Columns("COD_USUARIO").Visible = False
    'grdAnulacion.Columns("COD_TIPO_DOCUMENTO").Visible = False
End Sub

Sub SeteaGrillaReImp()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("NUM_PROFORMA", "FCH_REGISTRA", "NOMBRES", "MOTORIZADO", "DES_AUX_CLI_DIRECC")
                      
    arrCaption = Array("Proforma", "Fecha Emision", "Cliente", "Motorizado", "Direcciòn")
                       
    arrAncho = Array(1200, 1800, 2000, 1800, 2000)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
                          
    grdProformado.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub



Sub Opciones_Formato(arrCampos As Variant, arrCaption As Variant, arrAncho As Variant, arrAlineacion As Variant, indice As Boolean, Optional arrFoco As Variant)
Dim Columna As TrueDBGrid70.Column
    If indice = True Then
        arrCampos = Array("Sema", "FCH_REGISTRA", "NUM_PROFORMA", "DES_ESTADO", "FLG_URGENTE_OP", "COD_LOCAL_REF", "DES_MOTORIZADO", "DES_AUX_CLI_DIRECC", "UBIGEO_LARGO", "FLG_NEC_TRANF", "NombreConvenio", "flg_convenio", "FCH_HORA_PACT_ENTR", "TE", "REFERENCIA", "FLG_NEC_TRANSF", "COD_ESTADO", "COD_LOCAL", "FLG_TRANSFERENCIA", "FLG_URGENTE", "OBS_NOTA_RUTEO")
        arrCaption = Array("Sema", "Fecha Generada", "Num. Pedido", "Estado", "URGENTE", "Local", "Motorizado", "Dirección", "Dirección", "FLG_NEC_TRANF", "Convenio", "flg_convenio", "Fecha Entrega", "T.E.", "Referencia", "Tranf", "COD_ESTADO", "COD_LOCAL", "", "FLG_URGENTE", "Observacion Ruta")
        arrAncho = Array(300, 1550, 1000, 500, 200, 400, 1000, 0, 3500, 700, 2000, 100, 1800, 600, 2200, 0, 0, 0, 0, 0, 0)
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft)
    Else
        arrCampos = Array("Sema", "FCH_REGISTRA", "NUM_PROFORMA", "DES_ESTADO", "FLG_URGENTE_OP", "COD_LOCAL_REF", "DES_MOTORIZADO", "DES_AUX_CLI_DIRECC", "UBIGEO_LARGO", "FLG_NEC_TRANF", "NombreConvenio", "flg_convenio", "FCH_HORA_PACT_ENTR", "TE", "REFERENCIA", "FLG_NEC_TRANSF", "COD_ESTADO", "COD_LOCAL", "FLG_TRANSFERENCIA", "FLG_URGENTE", "OBS_NOTA_RUTEO")
        arrCaption = Array("Sema", "Fecha Generada", "Num. Pedido", "Estado", "URGENTE", "Local", "Motorizado", "Dirección", "Dirección", "FLG_NEC_TRANF", "Convenio", "flg_convenio", "Fecha Entrega", "T.E.", "Referencia", "Tranf", "COD_ESTADO", "COD_LOCAL", "", "FLG_URGENTE", "Observacion Ruta")
        arrAncho = Array(300, 1550, 1000, 500, 200, 400, 0, 0, 3500, 700, 2000, 100, 1800, 600, 0, 0, 0, 0, 0, 0, 3200)
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft)
    End If
    
    grdPedidos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdPedidos.RowHeight = 0
    grdPedidos.RowHeight = grdPedidos.RowHeight * 1.2
    grdPedidos.Columns(3).Font.Bold = True
    If Option1.Value = True Then
        grdPedidos.Columns(5).BackColor = pdblColorFondo
        grdPedidos.Columns(6).BackColor = pdblColorFondo
    End If
    grdPedidos.Columns(7).Visible = False
    grdPedidos.Columns(7).AllowSizing = False
    grdPedidos.Columns(9).Visible = False
    grdPedidos.Columns(9).AllowSizing = False
    grdPedidos.Columns(11).Visible = False
    grdPedidos.Columns(19).Visible = False
    grdPedidos.Columns(15).Visible = False
    grdPedidos.Columns(16).Visible = False
    grdPedidos.Columns(17).Visible = False
    grdPedidos.Columns(18).Visible = False
    grdPedidos.Columns(19).Visible = False
    grdPedidos.Columns(4).FetchStyle = True
    grdPedidos.Columns(15).FetchStyle = True
    grdPedidos.Columns(19).FetchStyle = True
    
    For Each Columna In grdPedidos.Columns
        Columna.AllowSizing = False
    Next
    
          
    If indice = True Then
        grdPedidos.Columns(6).Visible = True
        grdPedidos.Columns(14).Visible = True
        grdPedidos.Columns(20).Visible = False
        grdPedidos.MultiSelect = 2
    ElseIf indice = False Then
        grdPedidos.Columns(6).Visible = False
        grdPedidos.Columns(14).Visible = False
        grdPedidos.Columns(20).Visible = True
        grdPedidos.MultiSelect = 1
    End If
    
   ' grdPedidos.FetchRowStyle = True
    
    '''''''''''''''''''''''''''''''''
    grdPedidos.Columns(0).ValueItems.Translate = True
    Dim ValueItem3 As New TrueDBGrid70.ValueItem
    ValueItem3.DisplayValue = ImageList2.ListImages(3).Picture
    ValueItem3.Value = "1"
    grdPedidos.Columns(0).ValueItems.Add ValueItem3
    Set ValueItem3 = Nothing

    Dim ValueItem4 As New TrueDBGrid70.ValueItem
    ValueItem4.DisplayValue = ImageList2.ListImages(4).Picture
    ValueItem4.Value = "2"
    grdPedidos.Columns(0).ValueItems.Add ValueItem4
    Set ValueItem4 = Nothing
    
    Dim ValueItem5 As New TrueDBGrid70.ValueItem
    ValueItem5.DisplayValue = ImageList2.ListImages(5).Picture
    ValueItem5.Value = "3"
    grdPedidos.Columns(0).ValueItems.Add ValueItem5
    Set ValueItem5 = Nothing
    
End Sub

Sub ListaPedido()
    Dim Estado As String
    Dim posActual As Variant
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant
    
    Frame3.Visible = True
    If Option2.Value = True Then Estado = objPedido.PedidoVerificado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, False, arrFoco)
    If Option1.Value = True Then Estado = objPedido.PedidoAsignado: Call Opciones_Formato(arrCampos, arrCaption, arrAncho, arrAlineacion, True, arrFoco)
    'posActual = grdPedidos.Bookmark
    Set grdPedidos.DataSource = objPedido.ListaDeliveryProvincia(objUsuario.CodigoEmpresa, Estado, objUsuario.CodigoLocal)
    'grdPedidos.Bookmark = posActual
End Sub

Private Sub EstadoBotones()
    If Option1.Value = True Then
        'cmdVerificado.Enabled = False
        cmdDetalle.Enabled = True
        cmdAvisado.Enabled = True
        cmdLlevando.Enabled = True
        cmdLlegadaDestino.Enabled = True
        cmdEntregado.Enabled = True
        cmdLlegadaLocal.Enabled = True
        'cmdLiberar.Enabled = True
        'cmdAnular.Enabled = True
        'cmdReclamo.Enabled = True
    End If
    
    If Option2.Value = True Then
        'cmdVerificado.Enabled = True
        cmdDetalle.Enabled = True
        cmdAvisado.Enabled = False
        cmdLlevando.Enabled = False
        cmdLlegadaDestino.Enabled = False
        cmdEntregado.Enabled = False
        cmdLlegadaLocal.Enabled = False
        'cmdLiberar.Enabled = False
        'cmdAnular.Enabled = True
        'cmdReclamo.Enabled = True
    End If
End Sub


Private Sub grdPedidos_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If (grdPedidos.Columns("COD_ESTADO").Value = objPedido.PedidoEntregado) Then
        'cmdAnular.Enabled = False
        'cmdLiberar.Enabled = False
    Else
        'cmdAnular.Enabled = True
        'cmdLiberar.Enabled = True
    End If

   Set grdTiempo.DataSource = objPedido.ListaEstados(objUsuario.CodigoEmpresa, grdPedidos.Columns("NUM_PROFORMA"))
   
End Sub

Private Sub cmdLlevando_Click()
Dim strMensaje As String
Dim Bookmark As Variant
Dim k, i As Integer

On Error GoTo Control

    If Not cmdLlevando.Enabled Then Exit Sub
    
    If grdPedidos.Columns("COD_ESTADO") = objPedido.PedidoAvisado Then
        If MsgBox("Desea cambiar a Llevando sin necesidad que el pedido este Proformado ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirme") = vbNo Then
            If (grdPedidos.ApproxCount = 0) Or (grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoProforma) Then Exit Sub
        End If
    End If
    
    Bookmark = grdPedidos.Bookmark
    
    i = grdPedidos.SelBookmarks.Count - 1
    If i > 0 Then
        For k = i To 0 Step -1
            grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
            strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlevando, objUsuario.CODIGO)
        Next
    ElseIf grdPedidos.ApproxCount > 0 Then
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlevando, objUsuario.CODIGO)
    End If
    
    If strMensaje <> "" Then
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    
    ListaPedido
    grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub


'Private Sub CmdActualizar_Click()
'    On Error GoTo Control
'    ListaPedido
'
'Exit Sub
'Control:
'    MsgBox Err.Description, vbOKOnly + vbCritical, "Error :" & Err.Number
'End Sub

Private Sub cmdLlegadaDestino_Click()
    Dim strMensaje As String
    Dim Bookmark As Variant
    Dim k, i As Integer
    
On Error GoTo Control

    If Not cmdLlegadaDestino.Enabled Then Exit Sub

    If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoLlevando Then Exit Sub
    
    i = grdPedidos.SelBookmarks.Count - 1
    For k = i To 0 Step -1
        grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegada, objUsuario.CODIGO)
    Next
    
    If grdPedidos.SelBookmarks.Count = 0 Then
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegada, objUsuario.CODIGO)
    End If
    If strMensaje <> "" Then
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    Bookmark = grdPedidos.Bookmark
    ListaPedido
    grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdEntregado_Click()
    Dim strMensaje As String
    Dim Bookmark  As Variant
    Dim k, i As Integer
    
On Error GoTo Control

    If Not cmdEntregado.Enabled Then Exit Sub
    If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoLlegada Then Exit Sub
    
    i = grdPedidos.SelBookmarks.Count - 1
    For k = i To 0 Step -1
        grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoEntregado, objUsuario.CODIGO)
    Next
    
    If grdPedidos.SelBookmarks.Count = 0 Then
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoEntregado, objUsuario.CODIGO)
    End If
    If strMensaje <> "" Then
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    Bookmark = grdPedidos.Bookmark
    ListaPedido
    grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdLlegadaLocal_Click()
    Dim Bookmark As Variant
    Dim strMensaje As String
    Dim k, i As Integer
    
On Error GoTo Control


    Call cmdLlegadaDestino_Click
    Call cmdEntregado_Click
    

    If Not cmdLlegadaLocal.Enabled Then Exit Sub
    If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") <> objPedido.PedidoEntregado Then Exit Sub
    
    i = grdPedidos.SelBookmarks.Count - 1
    For k = i To 0 Step -1
        grdPedidos.Bookmark = grdPedidos.SelBookmarks.Item(k)
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegadaLocal, objUsuario.CODIGO)
    Next
    
    If grdPedidos.SelBookmarks.Count = 0 Then
        strMensaje = objPedido.CambiaEstado(objUsuario.CodigoEmpresa, grdPedidos.Columns("COD_LOCAL"), grdPedidos.Columns("NUM_PROFORMA"), objPedido.PedidoLlegadaLocal, objUsuario.CODIGO)
    End If
    If strMensaje <> "" Then
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
    Bookmark = grdPedidos.Bookmark
    ListaPedido
    grdPedidos.Bookmark = Bookmark

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdDetalle_Click()

    If Not cmdDetalle.Enabled Then Exit Sub
    If grdPedidos.ApproxCount = 0 Then Exit Sub
    'If grdPedidos.ApproxCount = 0 Or grdPedidos.Columns("COD_ESTADO") = objPedido.PedidoAsignado Or Option1.Value = True Then Exit Sub
    If grdPedidos.DataSource("COD_LOCAL") = "" Then Exit Sub
    On Error GoTo handle
    With frm_DLV_AsignaMotorizado
        .strLocal = "" & grdPedidos.DataSource("COD_LOCAL")
        .strNumProforma = "" & grdPedidos.DataSource("NUM_PROFORMA")
        .strLocalPedido = "" & grdPedidos.DataSource("COD_LOCAL_PRECIO")
        .CodCliente = "" & grdPedidos.DataSource("COD_CLIENTE")
        .CodDireccionCli = "" & grdPedidos.DataSource("COD_DIRECCION_CLI")
        frm_DLV_AsignaMotorizado.Show vbModal
        ListaPedido
    End With
''''''''''''''    frm_DLV_Ruteo.strLocal = grdPedidos.Columns("COD_LOCAL")
''''''''''''''    frm_DLV_Ruteo.strLocalPedido = grdPedidos.Columns("COD_LOCAL_REF")
''''''''''''''    frm_DLV_Ruteo.strnumProforma = grdPedidos.Columns("NUM_PROFORMA")
''''''''''''''    frm_DLV_Ruteo.bolEsLlamadoCab = False
''''''''''''''    If grdPedidos.Columns("COD_ESTADO") = objPedido.PedidoAsignado Or Option1.Value = True Then
''''''''''''''        Call CambioTipoTransferencia
''''''''''''''    End If
''''''''''''''    If Option2.Value = True Then
''''''''''''''        If grdPedidos.Columns("FLG_TRANSFERENCIA") = 1 Then
''''''''''''''            Call CambioTipoTransferencia
''''''''''''''        End If
''''''''''''''    End If
''''''''''''''    frm_DLV_Ruteo.Show vbModal
    Exit Sub
    
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    
End Sub


Sub CambioTipoTransferencia()
    frm_DLV_Ruteo.cboRuta1.Enabled = False
    frm_DLV_Ruteo.cboLocalAsig.Enabled = False
    frm_DLV_Ruteo.cmdTranferencia.Enabled = False
    frm_DLV_Ruteo.Check1.Enabled = False
    'frm_DLV_Ruteo.chkLocConStock.Enabled = False
    frm_DLV_Ruteo.grdPedidoStock.Enabled = False
    
    frm_DLV_Ruteo.cboRuta1.Visible = False
    frm_DLV_Ruteo.cboLocalAsig.Visible = False
    frm_DLV_Ruteo.cmdTranferencia.Visible = False
    frm_DLV_Ruteo.Check1.Visible = False
    'frm_DLV_Ruteo.chkLocConStock.Visible = False
    
End Sub
