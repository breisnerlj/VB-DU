VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm_VTA_Stock_Total 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Stock"
   ClientHeight    =   6810
   ClientLeft      =   1410
   ClientTop       =   285
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H007D004F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Datos del producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   6255
      Begin VB.TextBox txtEstAbast 
         Appearance      =   0  'Flat
         ForeColor       =   &H007D004F&
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   3555
      End
      Begin VB.TextBox lblCodigo 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Abastecimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   893
         Width           =   1275
      End
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
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   4035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Busqueda de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   60
      TabIndex        =   6
      Top             =   1500
      Width           =   6255
      Begin vbp_Ventas.ctlGrilla ctlGrilla1 
         Height          =   3555
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   6271
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Locales Cercanos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Visible         =   0   'False
         Width           =   1695
      End
      Begin vbp_Ventas.ctlDataCombo cboDistrito 
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Toda la Cadena"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1380
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   9
         Top             =   1500
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   8
         Top             =   1500
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Lima"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   7
         Top             =   1500
         Visible         =   0   'False
         Width           =   855
      End
      Begin vbp_Ventas.ctlDataCombo cboZona 
         Height          =   315
         Left            =   4200
         TabIndex        =   12
         Top             =   1500
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   11
         Top             =   1320
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3405
      Picture         =   "frm_VTA_Stock_Total.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2025
      Picture         =   "frm_VTA_Stock_Total.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   6120
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
            Picture         =   "frm_VTA_Stock_Total.frx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Stock_Total.frx":10AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Stock_Total.frx":1648
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_VTA_Stock_Total"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodigoProducto As String
Public strDescripcionProducto As String
Public strTranferencias As Boolean
Public G_LOCAL_ASIGNADO As String
Public strStockLocal As String
Public strCantPedida As String
Public strFalta As String
Public strCodZona As String
Public intCantFraccionamiento As Integer
'Public intCantProd As Integer
'Public intCantProd1 As Integer
'Public intCantProdFrac As Integer
'Public intCantProdFrac1 As Integer
Public Dif As Integer
Public FlgFraccionado As Boolean
Dim objProducto  As New clsProducto
Dim objZona As New clsZona
Dim Falta As Integer
Dim Ingresa As Integer
Dim resto As Double
Dim division As Double
Dim Disponible As New TrueDBGrid70.Style
Dim NoDisponible As New TrueDBGrid70.Style


Private Sub cboDistrito_Change()
   On Error GoTo Control

    Busca

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Private Sub cboZona_Click(Area As Integer)
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
                            vbNullString, _
                            vbNullString
    Next

    frm_DLV_Ruteo.grdTransferencia.Rebind
    Unload Me
    
    Exit Sub
    
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Sub Busca()
    If Not strCodigoProducto = "" Then
        If optFiltro(0).Value = True Then
            Set ctlGrilla1.DataSource = objProducto.ListaStockTotalVta(objUsuario.CodigoEmpresa, strCodigoProducto, , , "001")
        ElseIf optFiltro(1).Value = True Then
            Set ctlGrilla1.DataSource = objProducto.ListaStockTotalVta(objUsuario.CodigoEmpresa, strCodigoProducto, "1", , "001")
        ElseIf optFiltro(2).Value = True Then
            Set ctlGrilla1.DataSource = objProducto.ListaStockTotalVta(objUsuario.CodigoEmpresa, strCodigoProducto, "0", , "001")
        ElseIf optFiltro(3).Value = True Then
            Set ctlGrilla1.DataSource = objProducto.ListaStockTotalVta(objUsuario.CodigoEmpresa, strCodigoProducto, "", cboZona.BoundText, "001")
        ElseIf optFiltro(4).Value = True Then
            Set ctlGrilla1.DataSource = objProducto.ListaStockTotalVta(objUsuario.CodigoEmpresa, strCodigoProducto, "", cboZona.BoundText, "001", "", objUsuario.CodigoLocal)
        ElseIf optFiltro(5).Value = True Then
            Set ctlGrilla1.DataSource = objProducto.ListaStockTotalVta(objUsuario.CodigoEmpresa, strCodigoProducto, "", cboZona.BoundText, "001", "", "", Mid(objUsuario.UbigeoLocal, 1, 4) & cboDistrito.BoundText)
        End If
    End If
End Sub

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

    
    
    Dim intUnidad, intFraccion As Double
    
    '''intFraccion = Mid(IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value), InStr(1, IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value), "F") + 1, Len(IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value)))
    '''intUnidad = Val(Mid(IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value), 1, IIf(InStr(1, IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value), "F") = 0, 1, InStr(1, IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value), "F")))) * intCantFraccionamiento
    
    
    
    Dim c As Integer
    Dim s As String
    
    s = IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value)
    c = InStr(s, "F")
    
    
    
    If c > 0 Then
        intFraccion = Val(Mid(s, c + 1, Len(s)))
        intUnidad = Val(Mid(s, 1, c - 1)) * intCantFraccionamiento
    Else
        intFraccion = 0
        intUnidad = Val(Mid(s, 1, Len(s))) * intCantFraccionamiento
    End If
    
    
    
    
    
    If Ingresa > Val(intFraccion + intUnidad) Then
        MsgBox "La cantidad es mayor que el stock", vbInformation + vbOKOnly, App.ProductName
        Exit Sub
    End If
    

    
    CantidadFaltante
    If Ingresa > Falta Then
        'MsgBox "La cantidad a Transferir es Mayor al Faltante", vbInformation + vbOKOnly, App.ProductName
        If MsgBox("La cantidad a Transferir es Mayor al Faltante" & Chr(13) & "Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirme") = vbNo Then
            Exit Sub
        End If
    End If
    '''ctlGrillaArray1.Array1 = objVenta.AgregaDistribucion(objUsuario.CodigoLocal, strCodigoProducto, lblDescripcion.Caption, Val(txtCantidad.Text), chkFraccionamiento.Value, 2, Pedido_DLV, ctlGrilla1.Columns(0), G_LOCAL_ASIGNADO)
    
    
    ''''objVenta.ProductoTransferido = strCodigoProducto
    SendKeys "{TAB}"

    CantidadFaltante
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub CantidadFaltante()
    Dim intCant, i As Integer
    
    division = 0: resto = 0: intCant = 0
    If Dif < 0 Then
        For i = 0 To objVenta.Transferencia.UpperBound(1)
            If objVenta.Transferencia(i, 3) = "F" Then
                intCant = intCant + Val(objVenta.Transferencia(i, 4))
            Else
                intCant = intCant + (Val(objVenta.Transferencia(i, 4)) * intCantFraccionamiento)
            End If
        Next
        
        If intCant > 0 And intCant > Abs(Dif) Then
'            lblFalta.Caption = ""
        Else
            Falta = Abs(Dif + intCant)
            division = Int(Abs(Dif + intCant) / intCantFraccionamiento)
            resto = Abs(Dif + intCant) Mod intCantFraccionamiento
'            lblFalta.Caption = IIf(division = 0, "", division) & IIf(resto = 0, "", "F" & resto)
        End If
    End If
End Sub

'Private Sub ctlGrilla1_DblClick()
'    If strTranferencias = True Then
'        Exit Sub
'    End If
'    On Error GoTo CtrlErr
'    If ctlGrilla1.ApproxCount = 0 Then Exit Sub
'    If ctlGrilla1.ApproxCount > 0 And ctlGrilla1.Columns(0).Value <> "" Then
'        'CAMBIAR POR LA NUEVA FUNCION
'        If ctlGrilla1.Columns(3).Value = "0" Then
'            MsgBox "El producto no tiene Stock", vbCritical + vbOKOnly, "Atención"
'            ctlGrilla1.SetFocus
'            Exit Sub
'        End If
'        '------------------------------
'        Dim strIdFrac  As String
'        Dim strIndicadorReceta  As String
'        Dim strIndicador As String
'
'        strIdFrac = objProducto.ListaDevFracciona(strCodigoProducto)
'        strIndicador = objProducto.CodIndicadorReceta(strCodigoProducto)
'
'        strIndicadorReceta = objProducto.IndicadorReceta(strIndicador)
'        'frm_VTA_CantidadProducto.flgEspecieValorada = grdProductos.DataSource("FLG_ESP_VAL").Value
'        frm_VTA_CantidadProducto.subDatos strCodigoProducto, strDescripcionProducto, "001", "Label1(0).Caption", Producto_Normal, strIdFrac, strIndicadorReceta, strIndicador
'
'        Unload Me
'    End If
'Exit Sub
'CtrlErr:
'    MsgBox Err.Description, vbOKOnly + vbInformation, Err.Number
'End Sub

Private Sub ctlGrilla1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
    If ctlGrilla1.Columns("stock").CellValue(Bookmark) = "Disponible" Then
         
            RowStyle = Disponible
Else
        RowStyle = NoDisponible
    End If
End Sub

'Private Sub ctlGrilla1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        ctlGrilla1_DblClick
'    End If
'End Sub

Private Sub ctlGrilla1_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If ctlGrilla1.ApproxCount <> 0 Then
        If division > 0 Then
            'txtCantidad.Text = Val(IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value))
'            chkFraccionamiento.Value = 0
        End If
        If resto > 0 Then
'            txtCantidad.Text = Mid(IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value), InStr(1, IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value), "F") + 1, Len(IIf(IsNull(ctlGrilla1.Columns(2).Value), 0, ctlGrilla1.Columns(2).Value)))
'            chkFraccionamiento.Value = 1
        End If
    End If
End Sub

'Private Sub ctlGrillaArray1_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyDelete And ctlGrillaArray1.ApproxCount <> 0
'            ctlGrillaArray1.Delete
'            ctlGrillaArray1.Rebind
'            CantidadFaltante
'        Case vbKeyReturn
'            SendKeys "{TAB}"
'    End Select
'End Sub


Private Sub Form_Load()
Dim i As Integer
On Error GoTo Control

lblCodigo.Text = strCodigoProducto
lblDescripcion.Caption = strDescripcionProducto
'lblLocalAsignado.Caption = G_LOCAL_ASIGNADO
'lblStockLocal.Caption = strStockLocal
'lblCantPedido.Caption = strCantPedida
'lblFalta.Caption = strFalta
objVenta.Transferencia.ReDim 0, -1, 0, 15
'optFiltro(3).Value = True

frm_VTA_CantidadProducto.flgStockTotal = True

txtEstAbast.Text = gclsOracle.FN_Valor("CMR.FN_COM_EST_ABASTE", strCodigoProducto, "D")

'''''''''''''''
If strTranferencias = True Then
'    Label3.Visible = True
'    txtCantidad.Visible = True
'    chkFraccionamiento.Visible = True
'    cmdTransferir.Visible = True
'    Label2.Visible = True
'    Label4.Visible = True
'    Label7.Visible = True
'    Label9.Visible = True
'    lblLocalAsignado.Visible = True
'    lblStockLocal.Visible = True
'    lblCantPedido.Visible = True
'    lblFalta.Visible = True
    cmdAceptar.Enabled = True
    cmdCancelar.Picture = ImageList1.ListImages(2).Picture
    cmdCancelar.Caption = "&Cancelar"
    Busca
Else
'    Label3.Visible = False
'    txtCantidad.Visible = False
'    chkFraccionamiento.Visible = False
'    cmdTransferir.Visible = False
'    Label2.Visible = False
'    Label4.Visible = False
'    Label7.Visible = False
'    Label9.Visible = False
'    lblLocalAsignado.Visible = False
'    lblStockLocal.Visible = False
'    lblCantPedido.Visible = False
'    lblFalta.Visible = False
    cmdAceptar.Enabled = False
    cmdCancelar.Picture = ImageList1.ListImages(3).Picture
    cmdCancelar.Caption = "&Salir"
End If

    Set cboZona.RowSource = objZona.Lista
    cboZona.ListField = "DES_ZONA"
    cboZona.BoundColumn = "COD_ZONA"
    
    Dim objDistrito As New clsUbigeo
    If Not objUsuario.UbigeoLocal = "" Then
        Set cboDistrito.RowSource = objDistrito.ListaDistrito(Mid(objUsuario.UbigeoLocal, 1, 2), Mid(objUsuario.UbigeoLocal, 3, 2))
        cboDistrito.ListField = "DESCRIPCION"
        cboDistrito.BoundColumn = "CODIGO"
    End If
    
    
    
    
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Local", "Codigo", "Descripción", "U/F", "Ctd Prod.", "Precio", "xTipoventa", "frac", "Origen", "Destino")
    arrAncho = Array(0, 0, 0, 700, 1000, 0, 0, 0, 2500, 0)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter)
    
    
    
'    ctlGrillaArray1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
'    ctlGrillaArray1.Columns(0).Visible = False
'    ctlGrillaArray1.Columns(1).Visible = False
'    ctlGrillaArray1.Columns(2).Visible = False
'    ctlGrillaArray1.Columns(5).Visible = False
'    ctlGrillaArray1.Columns(6).Visible = False
'    ctlGrillaArray1.Columns(7).Visible = False
'    ctlGrillaArray1.Columns(9).Visible = False
    
''    ctlGrillaArray1.Array1 = objVenta.Transferencia
    
'    For i = 0 To objVenta.Distribucion.UpperBound(1)
'        If objVenta.Distribucion(i, 1) = strCodigoProducto Then
'            ctlGrillaArray1.Array1 = objVenta.AgregaTransferencia(objVenta.Distribucion(i, 0), _
'                            objVenta.Distribucion(i, 1), _
'                            objVenta.Distribucion(i, 2), _
'                            objVenta.Distribucion(i, 4), _
'                            objVenta.Distribucion(i, 7), _
'                            objVenta.Distribucion(i, 5), _
'                            objVenta.Distribucion(i, 6), _
'                            objVenta.Distribucion(i, 8), _
'                            objVenta.Distribucion(i, 9))
'        End If
'    Next
    
    arrCampos = Array("cod_local", "des_local", _
                      "stock", "ST", "precio", _
                      "flg_fraccionamiento", "cod_indicador_receta", "EST_ABASTE")
                      
    arrCaption = Array("C. Local", "Local", _
                       "Stock", "S. Trans", "Precio", _
                       "", "", "E. Abast.")
                       
    arrAncho = Array(800, 2500, _
                     1800, 800, 500, _
                     0, 0, 0)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    ctlGrilla1.FetchRowStyle = True
    
    ctlGrilla1.Columns(4).Visible = False
    ctlGrilla1.Columns(5).Visible = False
    ctlGrilla1.Columns(6).Visible = False

    ctlGrilla1.Columns(7).Visible = False
    ctlGrilla1.MarqueeStyle = dbgNoMarquee
    

   CantidadFaltante
   ''aca determino los permisos
   '' If objUsuario.EsQuimico(objUsuario.Codigo) = True Then
        optFiltro(4).Visible = True
        optFiltro(5).Visible = True
        optFiltro(4).Value = True
   '' ElseIf objUsuario.EsTecnico(objUsuario.Codigo) = True Then
   ''     optFiltro(4).Visible = True
   ''     optFiltro(4).Value = True
   '' End If
    
Set Disponible = ctlGrilla1.Styles.Add("Disponible")
    Disponible.Font.Bold = True
    Disponible.Font.Size = 10
    
    Disponible.ForeColor = RGB(79, 0, 125)
    
Set NoDisponible = ctlGrilla1.Styles.Add("NoDisponible")
    'NoDisponible.Font.Bold = True
    'NoDisponible.ForeColor = vbRed
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frm_VTA_CantidadProducto.flgStockTotal = False
    Set objProducto = Nothing
End Sub

Private Sub optFiltro_Click(Index As Integer)
    Select Case Index
        Case 3
            cboZona.Visible = True
        Case 5
            cboDistrito.Visible = True
        Case Else
            cboZona.Visible = False
            cboDistrito.Visible = False
                Busca
        End Select
        
End Sub

'''''''Public Sub Datos(ByVal Codigo As String, Descripcion As String)
'''''''    strCodigoProducto = Codigo
'''''''    strDescripcionProducto = Descripcion
'''''''    lblCodigo.Text = strCodigoProducto
'''''''    lblDescripcion.Caption = strDescripcionProducto
'''''''
'''''''    objVenta.ProductoTransferido = strCodigoProducto
''''''''    ctlGrillaArray1.Array1 = objVenta.Distribucion
'''''''
'''''''    Busca
'''''''End Sub

Public Sub Cerrar()
    Unload Me
End Sub


