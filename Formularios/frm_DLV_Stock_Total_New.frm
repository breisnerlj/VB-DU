VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DLV_Stock_Total_New 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Total de la Cadena"
   ClientHeight    =   7545
   ClientLeft      =   1410
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Transferencia"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      TabIndex        =   27
      Top             =   5400
      Width           =   5895
      Begin vbp_Ventas.ctlGrillaArray ctlGrillaArray1 
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1720
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la transferencia"
      Height          =   1575
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5895
      Begin VB.Label lblDescripcion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   555
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000040&
         Height          =   315
         Left            =   900
         TabIndex        =   25
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Local Asignado"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lblLocalAsignado 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   1140
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Stock"
         Height          =   195
         Left            =   1560
         TabIndex        =   21
         Top             =   900
         Width           =   420
      End
      Begin VB.Label lblStockLocal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1290
         TabIndex        =   20
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Pedida"
         Height          =   195
         Left            =   3195
         TabIndex        =   19
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label lblCantPedido 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3300
         TabIndex        =   18
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Falta completar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4500
         TabIndex        =   17
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label lblFalta 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   4620
         TabIndex        =   16
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblLocalAsignado2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   1140
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Busqueda de Stock"
      Height          =   3700
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      Begin VB.CommandButton cmdVerValFrac 
         Caption         =   "&Ver Frac"
         Height          =   500
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3120
         Width           =   900
      End
      Begin vbp_Ventas.ctlGrillaArray ctlGrillaArray2 
         Height          =   1815
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
      End
      Begin VB.CommandButton cmdTransferir 
         Caption         =   "Transferir"
         Height          =   495
         Left            =   4800
         Picture         =   "frm_DLV_Stock_Total_New.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   960
      End
      Begin VB.CheckBox chkFraccionamiento 
         Caption         =   "&Fraccionamiento"
         Height          =   195
         Left            =   3000
         TabIndex        =   7
         Top             =   3300
         Width           =   1575
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Toda la Cadena"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Zona"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Lima"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin vbp_Ventas.ctlTextBox txtCantidad 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Tipo            =   3
         MaxLength       =   3
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
      Begin vbp_Ventas.ctlDataCombo cboZona 
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   810
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboCia 
         Height          =   315
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   2900
         _ExtentX        =   5106
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   3300
         Width           =   630
      End
      Begin VB.Label Label28 
         Caption         =   "Cia:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   420
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   3135
      Picture         =   "frm_DLV_Stock_Total_New.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   1815
      Picture         =   "frm_DLV_Stock_Total_New.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Stock_Total_New.frx":109E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Stock_Total_New.frx":1638
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_DLV_Stock_Total_New.frx":1BD2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_DLV_Stock_Total_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodigoProducto As String
Public strDescripcionProducto As String
Public strTranferencias As Boolean
Public G_LOCAL_ASIGNADO As String
Public G_LOCAL_SAP_ASIGNADO As String
Public strStockLocal As String
Public strCantPedida As String
Public strFalta As String
Public strCodZona As String
Public intCantFraccionamiento As Integer
'Public intCantProd As Integer
'Public intCantProd1 As Integer
'Public intCantProdFrac As Integer
'Public intCantProdFrac1 As Integer
Public DIF As Integer
Public FlgFraccionado As Boolean
Dim objProducto  As New clsProducto
Dim objLocal As New clsLocal
Dim objZona As New clsZona
Dim falta As Integer
Dim Ingresa As Integer
Dim resto As Double
Dim division As Double
Dim objWS As New clsWebService

Private Sub cboCia_Change()
  ' listar el stock segun cia
  Dim vsCia As String
  vsCia = cboCia.BoundText
  If vsCia = "99" Then
    cmdVerValFrac.Visible = True ' ver fracciones para Mi Farma
    ctlGrillaArray2.Columns(7).Visible = True
  Else
   cmdVerValFrac.Visible = False
   ctlGrillaArray2.Columns(7).Visible = False
  End If
  'MsgBox "Cia Elegida :" + vsCia
  Busca
  
End Sub

Private Sub cboZona_Change()
Busca
End Sub

Private Sub chkFraccionamiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim x As Integer
    
    On Error GoTo handle
    
    For i = 0 To objVenta.Transferencia.UpperBound(1)
        objVenta.AgregaDistribucion objVenta.Transferencia(i, 0), _
                            objVenta.Transferencia(i, 1), _
                            objVenta.Transferencia(i, 2), _
                            objVenta.Transferencia(i, 4), _
                            objVenta.Transferencia(i, 7), _
                            objVenta.Transferencia(i, 5), _
                            objVenta.Transferencia(i, 6), _
                            objVenta.Transferencia(i, 8), _
                            objVenta.Transferencia(i, 9), _
                            objVenta.Transferencia(i, 10), _
                            objVenta.Transferencia(i, 11)
    Next

    frm_DLV_Ruteo.grdTransferencia.Rebind
    frm_DLV_Ruteo.EvaluaLocalesTransf
    Unload Me
    
    Exit Sub
    
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Sub Busca()
    Dim sCodEmp As String
    Dim Tipo As String
    On Error GoTo handle
    
    sCodEmp = cboCia.BoundText
    
    If Not strCodigoProducto = "" Then
        Me.MousePointer = vbHourglass
        If optFiltro(0).Value = True Then
            'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(objUsuario.CodigoEmpresa, strCodigoProducto, , , "001")
            'MsgBox "mdiPrincipal.ctlCliente1.sCia: " + mdiPrincipal.ctlCliente1.sCia
            ' toda la cadena
              'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(sCodEmp, strCodigoProducto, , , "001")
              Tipo = "all"
        ElseIf optFiltro(1).Value = True Then
            'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(objUsuario.CodigoEmpresa, strCodigoProducto, "1", , "001")
            ' lima
             'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(sCodEmp, strCodigoProducto, "1", , "001")
             Tipo = "lima"
        ElseIf optFiltro(2).Value = True Then
            'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(objUsuario.CodigoEmpresa, strCodigoProducto, "0", , "001")
            ' Provincia
             'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(sCodEmp, strCodigoProducto, "0", , "001")
             'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(sCodEmp, strCodigoProducto, "0", , "001", "", mdiPrincipal.ctlCliente1.LocalDespacho)
             Tipo = "provincia"
        ElseIf optFiltro(3).Value = True Then
            'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(objUsuario.CodigoEmpresa, strCodigoProducto, "", cboZona.BoundText, "001")
            'Set ctlGrillaArray2.DataSource = objproducto.ListaStockTotal(objUsuario.CodigoEmpresa, strCodigoProducto, "", cboZona.BoundText, "001", "", mdiPrincipal.ctlCliente1.LocalDespacho)
            ' zonas
             Tipo = "zona"
             'Set ctlGrillaArray2.DataSource = objProducto.ListaStockTotal(sCodEmp, strCodigoProducto, "", cboZona.BoundText, "001", "", mdiPrincipal.ctlCliente1.LocalDespacho)
        End If
        'I. ECASTILLO 05.07.2020
        Dim Codigo As Variant
        Codigo = Array(strCodigoProducto)
        If Tipo <> "zona" Then StockTotal sCodEmp, Tipo, IIf(G_LOCAL_ASIGNADO = "", mdiPrincipal.ctlCliente1.LocalDespacho, G_LOCAL_ASIGNADO), Codigo
        'F. ECASTILLO 05.07.2020
        Me.MousePointer = vbDefault
    End If
Exit Sub
handle:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
'I. ECASTILLO 05.07.2020
Private Function StockTotal(CIA As String, Tipo As String, codigoLocal As String, Codigo As Variant)
On Error GoTo Err
    Dim arr As New XArrayDB
    Dim ArrCode() As String
    Dim obj As New Dictionary
    Dim x, i As Integer
    Dim resource As oraDynaset
    Dim codProductoPosu, codBTL, codLocalPosu As String
    
    ReDim ArrCode(UBound(Codigo))
    For i = 0 To UBound(Codigo)
        codProductoPosu = ""
        codProductoPosu = objProducto.GetCodPosu(Codigo(i))
        ArrCode(i) = codProductoPosu
    Next i
    codLocalPosu = objLocal.GetCodPosu(codigoLocal)
    Set obj = objWS.GetStockTotal(CIA, Tipo, codLocalPosu, ArrCode)
    arr.Clear
    arr.ReDim 0, -1, 0, 7
    Debug.Print obj("data")(1)("drugstoreId")
    '"cod_local", "cod_local_sap", "des_local", "stock", "precio", _
     "flg_fraccionamiento", "cod_indicador_receta", "unid_vta"
    For x = 1 To obj("data").Count()
            arr.AppendRows
            arr(x - 1, 1) = obj("data")(x)("drugstoreId")                                               'cod_local_sap
            codBTL = objLocal.GetCodBTL(arr(x - 1, 1))
            arr(x - 1, 0) = codBTL                                                                      'cod_local
            arr(x - 1, 2) = obj("data")(x)("drugstore")                                                 'des_local
            arr(x - 1, 3) = obj("data")(x)("listProducts")(1)("fractionType")(2)("stock")               'stock
            arr(x - 1, 4) = obj("data")(x)("listProducts")(1)("price")                                  'precio
            arr(x - 1, 5) = obj("data")(x)("listProducts")(1)("isFractional")                           'flg_fraccionamiento
            arr(x - 1, 6) = obj("data")(x)("listProducts")(1)("prescription")                         'cod_indicador_receta
            arr(x - 1, 7) = obj("data")(x)("listProducts")(1)("fractionType")(2)("fractionatedDesc")    'unid_vta
    Next x
    ctlGrillaArray2.Array1 = arr
    ctlGrillaArray2.Rebind
    Exit Function
Err:
    Err.Raise Err.Number, "frm_DLV_Stock_Total", Err.Description
End Function
'E. ECASTILLO 05.07.2020

Private Sub cmdTransferir_Click()
On Error GoTo handle
''''    Dim arrCampos As Variant
''''    Dim arrCaption As Variant
''''    Dim arrAncho As Variant
''''    Dim arrAlineacion As Variant
''''    arrCampos = Array("", "", "", "", "", "", "", "", "", "")
''''    arrCaption = Array("Local", "Codigo", "Descripción", "Und/Frac", "Unidades", "Ctd Frac", "xTipoventa", "Fracciones", "Origen", "Destino")
''''    arrAncho = Array(0, 0, 0, 1000, 1000, 0, 0, 1500, 1200, 0)
''''    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
''''    ctlGrillaArray1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
''''    ctlGrillaArray1.Columns(0).Visible = False
''''    ctlGrillaArray1.Columns(1).Visible = False
''''    ctlGrillaArray1.Columns(2).Visible = False
''''    ctlGrillaArray1.Columns(7).Visible = False
''''    ctlGrillaArray1.Columns(5).Visible = False
''''    ctlGrillaArray1.Columns(6).Visible = False
''''    ctlGrillaArray1.Columns(9).Visible = False

    If chkFraccionamiento.Value = 1 Then
        Ingresa = Val(txtCantidad.Text)
    Else
        Ingresa = Val(txtCantidad.Text) * intCantFraccionamiento
    End If
    
    Dim intUnidad, intFraccion As Double
    
    '''intFraccion = Mid(IIf(IsNull(ctlGrillaArray2.Columns(2).Value), 0, ctlGrillaArray2.Columns(2).Value), InStr(1, IIf(IsNull(ctlGrillaArray2.Columns(2).Value), 0, ctlGrillaArray2.Columns(2).Value), "F") + 1, Len(IIf(IsNull(ctlGrillaArray2.Columns(2).Value), 0, ctlGrillaArray2.Columns(2).Value)))
    '''intUnidad = Val(Mid(IIf(IsNull(ctlGrillaArray2.Columns(2).Value), 0, ctlGrillaArray2.Columns(2).Value), 1, IIf(InStr(1, IIf(IsNull(ctlGrillaArray2.Columns(2).Value), 0, ctlGrillaArray2.Columns(2).Value), "F") = 0, 1, InStr(1, IIf(IsNull(ctlGrillaArray2.Columns(2).Value), 0, ctlGrillaArray2.Columns(2).Value), "F")))) * intCantFraccionamiento
    
    
    
    Dim c As Integer
    Dim s As String
    
    s = IIf(IsNull(ctlGrillaArray2.Columns("STOCK").Value), 0, ctlGrillaArray2.Columns("STOCK").Value)
    c = InStr(s, "F")
    
    
    
    If c > 0 Then
        intFraccion = Val(Mid(s, c + 1, Len(s)))
        intUnidad = Val(Mid(s, 1, c - 1)) * intCantFraccionamiento
    Else
        intFraccion = 0
        intUnidad = Val(Mid(s, 1, Len(s))) * intCantFraccionamiento
    End If
    
    
    
    
    
    If Ingresa > Val(intFraccion + intUnidad) Then
        MsgBox "La cantidad es mayor que el stock", vbCritical + vbOKOnly, App.ProductName
        txtCantidad.SetFocus
        Exit Sub
    End If
    

    If Val(txtCantidad.Text) = 0 Then
        MsgBox "Ingresar cantidad a transferir", vbInformation + vbOKOnly, App.ProductName
        txtCantidad.SetFocus
        Exit Sub
        
    End If
    CantidadFaltante
    If Ingresa > falta Then
        'MsgBox "La cantidad a Transferir es Mayor al Faltante", vbInformation + vbOKOnly, App.ProductName
        If MsgBox("La cantidad a Transferir es Mayor al Faltante" & Chr(13) & "Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirme") = vbNo Then
            txtCantidad.SetFocus
            Exit Sub
        End If
    End If
    '''ctlGrillaArray1.Array1 = objVenta.AgregaDistribucion(objUsuario.CodigoLocal, strCodigoProducto, lblDescripcion.Caption, Val(txtCantidad.Text), chkFraccionamiento.Value, 2, Pedido_DLV, ctlGrillaArray2.Columns(0), G_LOCAL_ASIGNADO)
    ctlGrillaArray1.Array1 = objVenta.AgregaTransferencia(objUsuario.codigoLocal, strCodigoProducto, lblDescripcion.Caption, Val(txtCantidad.Text), chkFraccionamiento.Value, 2, Pedido_DLV, ctlGrillaArray2.Columns(0).Value, G_LOCAL_ASIGNADO, ctlGrillaArray2.Columns(1).Value, G_LOCAL_SAP_ASIGNADO)
    
    ''''objVenta.ProductoTransferido = strCodigoProducto
    SendKeys "{TAB}"
    ctlGrillaArray1.Rebind
    CantidadFaltante
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub CantidadFaltante()
    Dim intCant, i As Integer
    
    division = 0: resto = 0: intCant = 0
    If DIF < 0 Then
        For i = 0 To objVenta.Transferencia.UpperBound(1)
            If objVenta.Transferencia(i, 3) = "F" Then
                intCant = intCant + Val(objVenta.Transferencia(i, 4))
            Else
                intCant = intCant + (Val(objVenta.Transferencia(i, 4)) * intCantFraccionamiento)
            End If
        Next
        
        If intCant > 0 And intCant > Abs(DIF) Then
            lblFalta.Caption = ""
        Else
            falta = Abs(DIF + intCant)
            division = Int(Abs(DIF + intCant) / intCantFraccionamiento)
            resto = Abs(DIF + intCant) Mod intCantFraccionamiento
            lblFalta.Caption = IIf(division = 0, "", division) & IIf(resto = 0, "", "F" & resto)
        End If
    End If
End Sub

Private Sub cmdVerValFrac_Click()
    Dim s As String
    Dim msg As String
  
    If ctlGrillaArray2.ApproxCount = 0 Then Exit Sub
  
    s = CStr("" & gclsOracle.FN_Valor("CMR.PKG_PRODUCTO.FN_GET_VALFRAC_PROD", "", lblCodigo.Caption, ctlGrillaArray2.Columns(0).Value))
    
    msg = " Local : " & vbCrLf & _
          vbTab & ctlGrillaArray2.Columns("cod_local_sap").Value & " - " & ctlGrillaArray2.Columns("des_local").Value & vbTab & _
          vbCrLf & vbCrLf & _
          " Producto : " & vbCrLf & _
          vbTab & lblCodigo.Caption & " - " & lblDescripcion.Caption & vbTab & _
          vbCrLf & vbCrLf & _
          " Fraccionamiento : " & vbCrLf & _
          vbTab & s
  
    If Not s = "" Then
        MsgBox msg, vbInformation + vbOKOnly, "Valor Fracción"
    Else
        MsgBox "No existen datos de fraccionamiento" & vbCrLf & s, vbCritical + vbOKOnly, App.ProductName
    End If
End Sub



Private Sub ctlGrillaArray2_DblClick()
    If strTranferencias = True Then
        If txtCantidad.Visible = True Then txtCantidad.SetFocus
        Exit Sub
    End If
    On Error GoTo CtrlErr
    If ctlGrillaArray2.ApproxCount = 0 Then Exit Sub
    If ctlGrillaArray2.ApproxCount > 0 And ctlGrillaArray2.Columns(0).Value <> "" Then
        'CAMBIAR POR LA NUEVA FUNCION
        If ctlGrillaArray2.Columns(3).Value = "0" Then
            MsgBox "El producto no tiene Stock", vbCritical + vbOKOnly, "Atención"
            ctlGrillaArray2.SetFocus
            Exit Sub
        End If
        '------------------------------
        Dim strIdFrac  As String
        Dim strIndicadorReceta  As String
        Dim strIndicador As String
        
        strIdFrac = objProducto.ListaDevFracciona(strCodigoProducto, objUsuario.codigoLocal, objVenta.CodModalidadVenta)
        strIndicador = objProducto.CodIndicadorReceta(strCodigoProducto)
        
        strIndicadorReceta = objProducto.IndicadorReceta(strIndicador)
        'frm_VTA_CantidadProducto.flgEspecieValorada = grdProductos.DataSource("FLG_ESP_VAL").Value
        frm_VTA_CantidadProducto.subDatos strCodigoProducto, strDescripcionProducto, "001", "Lista de Productos", Producto_Normal, strIdFrac, strIndicadorReceta, strIndicador
        
        Unload Me
    End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbInformation, Err.Number
End Sub

Private Sub ctlGrillaArray2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ctlGrillaArray2_DblClick
    End If
End Sub

Private Sub ctlGrillaArray2_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If ctlGrillaArray2.ApproxCount <> 0 Then
        If division > 0 Then
            txtCantidad.Text = Val(IIf(IsNull(ctlGrillaArray2.Columns("STOCK").Value), 0, ctlGrillaArray2.Columns("STOCK").Value))
            chkFraccionamiento.Value = 0
        End If
        If resto > 0 Then
            txtCantidad.Text = Mid(IIf(IsNull(ctlGrillaArray2.Columns("STOCK").Value), 0, ctlGrillaArray2.Columns("STOCK").Value), InStr(1, IIf(IsNull(ctlGrillaArray2.Columns("STOCK").Value), 0, ctlGrillaArray2.Columns("STOCK").Value), "F") + 1, Len(IIf(IsNull(ctlGrillaArray2.Columns("STOCK").Value), 0, ctlGrillaArray2.Columns("STOCK").Value)))
            chkFraccionamiento.Value = 1
        End If
    End If
End Sub

Private Sub ctlGrillaArray1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete And ctlGrillaArray1.ApproxCount <> 0
            ctlGrillaArray1.Delete
            ctlGrillaArray1.Rebind
            CantidadFaltante
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub Form_Activate()
    cboZona.BoundText = strCodZona
End Sub

Private Sub Form_Load()
Dim i As Integer
On Error GoTo Control
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Local", "Codigo", "Descripción", "U/F", "Ctd Prod.", "Precio", "xTipoventa", "frac", "Origen", "Destino", "Origen", "Destino")
    arrAncho = Array(0, 0, 0, 700, 1000, 0, 0, 0, 0, 0, 1500, 1500)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    
    ctlGrillaArray1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    ctlGrillaArray1.Columns(0).Visible = False
    ctlGrillaArray1.Columns(1).Visible = False
    ctlGrillaArray1.Columns(2).Visible = False
    ctlGrillaArray1.Columns(5).Visible = False
    ctlGrillaArray1.Columns(6).Visible = False
    ctlGrillaArray1.Columns(7).Visible = False
    ctlGrillaArray1.Columns(8).Visible = False
    ctlGrillaArray1.Columns(9).Visible = False
    ctlGrillaArray1.Columns(11).Visible = False

    arrCampos = Array("cod_local", "cod_local_sap", "des_local", _
                      "stock", "precio", _
                      "flg_fraccionamiento", "cod_indicador_receta", "unid_vta")
                      
    arrCaption = Array("", "C. Local", "Local", _
                       "Stock", "Precio", _
                       "", "", "UNID.VTA")
                       
    arrAncho = Array(800, 800, 2500, _
                     800, 800, _
                     800, 800, 1000)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
    ctlGrillaArray2.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    ctlGrillaArray2.Columns(0).Visible = False
    ctlGrillaArray2.Columns(4).Visible = False
    ctlGrillaArray2.Columns(5).Visible = False
    ctlGrillaArray2.Columns(6).Visible = False

lblCodigo.Caption = strCodigoProducto
lblDescripcion.Caption = strDescripcionProducto
lblLocalAsignado.Caption = G_LOCAL_ASIGNADO
lblLocalAsignado2.Caption = G_LOCAL_SAP_ASIGNADO
lblStockLocal.Caption = strStockLocal
lblCantPedido.Caption = strCantPedida
lblFalta.Caption = strFalta
objVenta.Transferencia.ReDim 0, -1, 0, 15
optFiltro(3).Value = True

frm_VTA_CantidadProducto.flgStockTotal = True

If strTranferencias = True Then
    
    cmdVerValFrac.Visible = False ' JCT, no visible para trans
    Label3.Visible = True
    txtCantidad.Visible = True
    chkFraccionamiento.Visible = True
    cmdTransferir.Visible = True
    Label2.Visible = True
    Label4.Visible = True
    Label7.Visible = True
    Label9.Visible = True
    'lblLocalAsignado.Visible = True
    lblLocalAsignado2.Visible = True
    lblStockLocal.Visible = True
    lblCantPedido.Visible = True
    lblFalta.Visible = True
    cmdAceptar.Enabled = True
    cmdCancelar.Picture = ImageList1.ListImages(2).Picture
    cmdCancelar.Caption = "&Cancelar"
    Busca
Else
    Label3.Visible = False
    txtCantidad.Visible = False
    chkFraccionamiento.Visible = False
    cmdTransferir.Visible = False
    Label2.Visible = False
    Label4.Visible = False
    Label7.Visible = False
    Label9.Visible = False
    'lblLocalAsignado.Visible = False
    lblLocalAsignado2.Visible = False
    lblStockLocal.Visible = False
    lblCantPedido.Visible = False
    lblFalta.Visible = False
    cmdAceptar.Enabled = False
    cmdCancelar.Picture = ImageList1.ListImages(3).Picture
    cmdCancelar.Caption = "&Salir"
End If

    Set cboZona.RowSource = objZona.Lista
    cboZona.ListField = "DES_ZONA"
    cboZona.BoundColumn = "COD_ZONA"
    
    '++++++++++ Begin jct 28Mar12, carga cia
    
    Set cboCia.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCia.ListField = "Des"
    cboCia.BoundColumn = "Cod"
    cboCia.BoundText = mdiPrincipal.ctlCliente1.sCia
    'mdiprincipal.ctlCliente1.
    
    '++++++++++ end
    
    ctlGrillaArray1.Array1 = objVenta.Transferencia
    
    For i = 0 To objVenta.Distribucion.UpperBound(1)
        If objVenta.Distribucion(i, 1) = strCodigoProducto Then
            ctlGrillaArray1.Array1 = objVenta.AgregaTransferencia(objVenta.Distribucion(i, 0), _
                            objVenta.Distribucion(i, 1), _
                            objVenta.Distribucion(i, 2), _
                            objVenta.Distribucion(i, 4), _
                            objVenta.Distribucion(i, 7), _
                            objVenta.Distribucion(i, 5), _
                            objVenta.Distribucion(i, 6), _
                            objVenta.Distribucion(i, 8), _
                            objVenta.Distribucion(i, 9), _
                            objVenta.Distribucion(i, 10), _
                            objVenta.Distribucion(i, 11))
        End If
    Next

    If mdiPrincipal.ctlCliente1.sCia <> "99" Then
     cmdVerValFrac.Visible = False ' JCT, no visible para trans
     ctlGrillaArray2.Columns(7).Visible = False
    End If
    
   CantidadFaltante
   '' Autor:Juan Arturo Escate Espichan
   '' Proposito: A solicitud de RMattos que se se carge por defecto la zona del local de despacho
   '' Fecha: 18/08/2014
   Dim strZonaDespacho As String
   strZonaDespacho = "" & objZona.DevuelveZona(objVenta.LocalDespacho)
   If Not strZonaDespacho = "" Then
        cboZona.BoundText = strZonaDespacho
   End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frm_VTA_CantidadProducto.flgStockTotal = False
    Set objProducto = Nothing
End Sub

Private Sub optFiltro_Click(Index As Integer)
 
 '+++++ jct
'' Dim s As String
'' s = cboCia.BoundText
'' If (Len(s) = 0) Then
''  MsgBox "Debe Elegir una Cia..."
''  Exit Sub ' no continuar
'' End If
 '+++++ jct
 
    Select Case Index
        Case 3 ' zonas
            cboZona.Visible = True
            Busca
        Case Else
            cboZona.Visible = False
             Busca
        End Select
        
End Sub

Public Sub Datos(ByVal Codigo As String, Descripcion As String)
On Error GoTo handle
    strCodigoProducto = Codigo
    strDescripcionProducto = Descripcion
    lblCodigo.Caption = strCodigoProducto
    lblDescripcion.Caption = strDescripcionProducto
        
    objVenta.ProductoTransferido = strCodigoProducto
    ctlGrillaArray1.Array1 = objVenta.Distribucion
    Busca
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub Cerrar()
    Unload Me
End Sub




