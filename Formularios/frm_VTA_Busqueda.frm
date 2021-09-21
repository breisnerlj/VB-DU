VERSION 5.00
Begin VB.Form frm_VTA_Busqueda 
   BorderStyle     =   0  'None
   Caption         =   "---------------------------"
   ClientHeight    =   8505
   ClientLeft      =   840
   ClientTop       =   750
   ClientWidth     =   13095
   Icon            =   "frm_VTA_Busqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTipoBusqueda 
      Height          =   315
      ItemData        =   "frm_VTA_Busqueda.frx":030A
      Left            =   4800
      List            =   "frm_VTA_Busqueda.frx":031D
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   360
      Width           =   2415
   End
   Begin vbp_Ventas.ctlGrilla grdProductos 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   4683
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame fraComplementarios 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   0
      TabIndex        =   16
      Top             =   6885
      Width           =   12690
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   6
         Left            =   2280
         TabIndex        =   18
         Top             =   570
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Activar / Desactivar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2700
         TabIndex        =   17
         Top             =   570
         Width           =   3240
      End
   End
   Begin vbp_Ventas.ctlGrilla grdComplementarios 
      Height          =   1470
      Left            =   0
      TabIndex        =   13
      Top             =   6945
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2593
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtBuscar 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      Tipo            =   8
      TABAuto         =   0   'False
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   9240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraAlternativos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   0
      TabIndex        =   10
      Top             =   4560
      Width           =   12690
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Activar / Desactivar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2700
         TabIndex        =   12
         Top             =   630
         Width           =   3480
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   4
         Left            =   2280
         TabIndex        =   11
         Top             =   630
         Width           =   255
      End
   End
   Begin vbp_Ventas.ctlGrilla grdAlternativos 
      Height          =   1590
      Left            =   0
      TabIndex        =   3
      Top             =   4650
      Visible         =   0   'False
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2805
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmd_salir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   10560
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Ver Stock"
      Height          =   255
      Left            =   4560
      TabIndex        =   30
      Top             =   870
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "F11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMsgComercialAlt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   60
      TabIndex        =   28
      Top             =   6240
      Visible         =   0   'False
      Width           =   12435
   End
   Begin VB.Label labelProductosEnPromocion 
      AutoSize        =   -1  'True
      Caption         =   "Ctrl+P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   5800
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label labelProductosEnPromocionDesc 
      AutoSize        =   -1  'True
      Caption         =   "Ver Promociones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   6520
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ver Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   3000
      TabIndex        =   25
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ctrl+I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   0
      Left            =   2400
      TabIndex        =   24
      Top             =   840
      Width           =   510
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   7
      Left            =   4800
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblTipoBusqueda 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   0
      Left            =   7380
      TabIndex        =   22
      Top             =   360
      Width           =   2580
   End
   Begin VB.Label LblMsgComercial 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   20
      Top             =   3900
      Visible         =   0   'False
      Width           =   12435
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F9"
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
      Left            =   60
      TabIndex        =   15
      Top             =   6690
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Productos Complementarios :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   14
      Top             =   6690
      Width           =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Productos Alternativos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   4350
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Productos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   1710
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F3"
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
      Left            =   60
      TabIndex        =   7
      Top             =   4350
      Width           =   225
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   60
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_Busqueda.frx":037C
      Top             =   60
      Width           =   240
   End
   Begin VB.Image imgComunicandonos 
      Height          =   375
      Left            =   7920
      MouseIcon       =   "frm_VTA_Busqueda.frx":0906
      MousePointer    =   99  'Custom
      Picture         =   "frm_VTA_Busqueda.frx":0A58
      Stretch         =   -1  'True
      ToolTipText     =   "Comunicándonos"
      Top             =   2520
      Width           =   1170
   End
End
Attribute VB_Name = "frm_VTA_Busqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Dim objDocumentoPago As New clsDocumentoPago
Dim objLocal As New clsLocal
Dim strTipoPrecio As String
Private lblnModo As Boolean
Dim strIdFrac As String
Dim strIndicadorReceta As String
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim rsClonBusqueda As oraDynaset
Dim strgrdFoco As String
Dim nVeces As String
Dim strTarjetaCMR As String
Dim esTarjeta As Boolean
Dim productosenpromocion As String

Private Sub cboTipoBusqueda_Click()
    Debug.Print cboTipoBusqueda.ListIndex
    
    
    
    'cboTipoBusqueda.List (cboTipoBusqueda.ListIndex)
'    If Not cboTipoBusqueda.ListIndex = 0 Then
'    frm_VTA_TipoBusqueda.Show vbModal
'    lblTipoBusqueda(0).Caption = "Alergia/Anthistaminico"
'    Else
'    lblTipoBusqueda(0).Caption = ""
'    End If
    cmdBuscar_Click
End Sub

Public Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'I.ECASTILLO 27.10.2020
    If objVenta.isLocalDcCappa = "1" And objVenta.ptmModalidad = Venta_Convenio Then
        MsgBox "La modalidad de venta convenio no es soportada por locales Cappa", vbOKOnly + vbInformation, App.ProductName
        Unload Me
    End If
    'F.ECASTILLO 27.10.2020
End Sub

Private Sub Form_Load()
    On Error GoTo Control
    esTarjeta = False
    nVeces = 0
    'ecastillo 30/09/2019
    cboTipoBusqueda.Clear
    cboTipoBusqueda.AddItem "POR PRODUCTO"
    cboTipoBusqueda.AddItem "POR PRINCIPIO ACTIVO"
    cboTipoBusqueda.AddItem "POR LABORATORIO"
    cboTipoBusqueda.ListIndex = 0
    setteaFormulario Me
    Format_Grilla
    mostrarAlternativos
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Public Sub cmdBuscar_Click()
Dim LocalDespacho As String
Dim strCiaDespacho As String
On Error GoTo Control

    gstrCodTarjetaFid = ""
    gstrCodTarjetaMon = ""
    'gintFidelizado = 0
  
    If txtBuscar.Text = "" Then Exit Sub

    If objUsuario.EsDelivery And objVenta.CodModalidadVenta <> "002" Then
        If Trim(mdiPrincipal.ctlCliente1.LocalAsignado) = "" Then MsgBox "Debe Ingresar un Cliente", vbInformation: Unload frm_VTA_Busqueda: mdiPrincipal.txtDLVTelefono.SetFocus: Exit Sub
        If Trim(mdiPrincipal.txtDLVTelefono.Text) = "" Then MsgBox "Debe Ingresar un Número de Teléfono", vbInformation: Unload frm_VTA_Busqueda: mdiPrincipal.txtDLVTelefono.SetFocus: Exit Sub
    End If

    If Len(Trim(txtBuscar.Text)) < 3 Then
        MsgBox "Ingresar como minimo 3 digitos", vbInformation + vbOKOnly, App.ProductName
        txtBuscar.SetFocus
        Exit Sub
    End If

    If frm_VTA_RecetarioM.pstrFlgRM = "" Then
        Format_Grilla
        Dim strCodigoLocal As String
        Dim tipoBusquedaStr As String
        Dim tipoBusquedaVal As Integer
        strCodigoLocal = mdiPrincipal.ctlCliente1.LocalAsignado
        If strCodigoLocal = "" Then strCodigoLocal = objUsuario.CodigoLocal
        ''agregado por pherrera 230108 no funcaba el local de despacho para sacar el stock
        LocalDespacho = objUsuario.CodigoLocal
        ''If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = "0" Then
        If objUsuario.EsDelivery Then
            LocalDespacho = mdiPrincipal.ctlCliente1.LocalDespacho
            strCiaDespacho = mdiPrincipal.ctlCliente1.sCia
        End If
        If strCiaDespacho = "" Then strCiaDespacho = objUsuario.CodigoEmpresa
        '++++++++ b jct, parametros para llamar a funcion para llenar grid
''        MsgBox "CIA : " + mdiPrincipal.ctlCliente1.sCia
''        Dim s As String
''        s = "mdiPrincipal.ctlCliente1.sCia, strCodigoLocal, strTipoPrecio, Trim(txtBuscar.Text), "", "", LocalDespacho, objVenta.CodModalidadVenta, objVenta.CodigoConvenio===> " _
''        + mdiPrincipal.ctlCliente1.sCia + "," + strCodigoLocal + "," + strTipoPrecio + "," + Trim(txtBuscar.Text) + "," + "null" + "," + "null" + "," + LocalDespacho + "," + objVenta.CodModalidadVenta + "," + objVenta.CodigoConvenio
''        MsgBox "Data Call: " + s
        
        '++++++++ e jct
'        tipoBusquedaStr = cboTipoBusqueda.List(cboTipoBusqueda.ListIndex)
'        If tipoBusquedaStr = "POR PRODUCTO" Then
'            tipoBusquedaVal = 0
'        ElseIf tipoBusquedaStr = "POR PRINCIPIO ACTIVO" Then
'            tipoBusquedaVal = 1
'        End If
        tipoBusquedaVal = cboTipoBusqueda.ListIndex
        
        Dim od As oraDynaset
        Set od = objProducto.Lista(objUsuario.CodigoEmpresa, strCodigoLocal, strTipoPrecio, Trim(txtBuscar.Text), "", "", LocalDespacho, objVenta.CodModalidadVenta, objVenta.CodigoConvenio, strCiaDespacho, tipoBusquedaVal)
        
        grdProductos.Limpiar
        LblMsgComercial(0).Visible = False
        lblMsgComercialAlt.Visible = False
        labelProductosEnPromocion.Item(1).Visible = False
        labelProductosEnPromocionDesc.Item(1).Visible = False
        If od(0) <> -1 Then
            Set grdProductos.DataSource = od 'objProducto.Lista(objUsuario.CodigoEmpresa, strCodigoLocal, strTipoPrecio, Trim(txtBuscar.Text), "", "", LocalDespacho, objVenta.CodModalidadVenta)
            If grdProductos.ApproxCount = 0 Then
                grdProductos_RegistroSeleccionado ""
                txtBuscar.Focus
            Else
                objVenta.busquedaNVeces = 1
                'I.CVIERA 21.12.2020 | 05.01.2021.REVISAR
'                If frm_VTA_MetodosSegmentos.strPrecioTipo <> "" Then
'                    If od(0) = "09938" Then
'                        grdProductos.Columns("PRECIO").Value = Format(frm_VTA_MetodosSegmentos.strPrecioTipo, "###,###.00")
'                        grdProductos.Columns("FLG_SEG").Value = 1
'                    End If
'                End If
                'F.CVIERA 21.12.2020
                'I.ECASTILLO 05.01.2021 | 06.01.2021
                Dim flg_ruteoA_cnv
                flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
                If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
                    GoTo cnvNoRuteaAuto
                End If
                Dim flg_2e_reserva
                flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
                Dim sCia As String
                Dim rsCia As oraDynaset
                Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, mdiPrincipal.ctlCliente1.LocalDespacho)
                If (rsCia.RecordCount > 0) Then
                  sCia = CStr(rsCia(1))
                End If
                Set rsCia = Nothing
                objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, mdiPrincipal.ctlCliente1.LocalDespacho, "RESERVA_STOCK_2DA")
                
                If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
                Else
                    grdProductos.DataSource.MoveFirst
                    While Not grdProductos.DataSource.EOF
                        If grdProductos.Columns("COD_PRODUCTO").Value = "09938" _
                            And frm_VTA_MetodosSegmentos.strPrecioTipo <> "" _
                        Then
                            grdProductos.Columns("PRECIO").Value = Format(frm_VTA_MetodosSegmentos.strPrecioTipo, "###,###.00")
                            grdProductos.Columns("FLG_SEG").Value = 1
                            GoTo exitWhile
                        End If
                        grdProductos.DataSource.MoveNext 'debido a que no se puede actualizar el datasource aqui se cae
                        'por lo que al actualizar precio se forzara la salida del while
                    Wend
                    grdProductos.DataSource.MoveFirst
                End If
exitWhile:
cnvNoRuteaAuto:
                'F.ECASTILLO 05.01.2021
                SendKeys "{TAB}"
            End If
        Else
            'Autor: Juan Arturo Escate Espichan
            'Fecha: 08/06/2011
            
            ''' <12-MAY-14  TCT Aqui Determina  si el codigo ingresado es un CUPON >
            
            Select Case objDocumentoPago.buscaFun(txtBuscar.Text)
                Case "CUPON"
                    AgregaCupon (CStr(txtBuscar.Text))
                Case "CONVENIO"
                    AgregaConvenio (CStr(txtBuscar.Text))
                Case "FIDELIZADO"
                    AgregaFidelizado (CStr(txtBuscar.Text))
                Case "MONEDERO"
                    If objVenta.EscaneoTarjeta = True Then
                        AgregaMonedero (CStr(txtBuscar.Text))
                    End If
                Case Else
                    If objVenta.ValidaTarjetaCMR(Trim(txtBuscar.Text)) > 0 Then
                        AgregaCMR (CStr(txtBuscar.Text))
                        'AgregaCMR (Replace(FN_TARJETA_CMR(Trim(txtBuscar.Text)), "%", ""))
                    End If
            End Select
        End If
    ElseIf frm_VTA_RecetarioM.pstrFlgRM = "1" Then
        Format_Grilla_RM
        If frm_VTA_RecetarioM.pstrRucProv = "" Then MsgBox "Seleccione a un proveedor", vbCritical, Caption: Exit Sub
        Set grdProductos.DataSource = objProducto.ListaRegMagistral(objUsuario.CodigoEmpresa, _
                                                                    objUsuario.CodigoLocal, _
                                                                    Trim(txtBuscar.Text), _
                                                                    frm_VTA_RecetarioM.pstrRucProv, _
                                                                    grdProductos.Columns("COD_TIPO_INSUMO").Value)
        If grdProductos.ApproxCount <= 0 Then
            ''MsgBox "El producto no esta asociado a este proveedor de Ruc Nº " & frm_VTA_RecetarioM.pstrRucProv & " ", vbCritical, Caption
            MsgBox "Producto no es un Insumo del RECETARIO MAGISTRAL", vbCritical, "Recetario Magistral"
            txtBuscar.selection
            txtBuscar.SetFocus
        Else
            SendKeys "{TAB}"
            grdProductos.Columns("COD_PRODUCTO").Visible = False
            ''grdProductos.Columns("DES_UND_CAPACIDAD_ABREV").Visible = False
            grdProductos.Columns("DES_CLASE_COM").Visible = False
            grdProductos.Columns("DES_CATEGORIA_COM").Visible = False
            grdProductos.Columns("STOCK").Visible = False
            grdProductos.Columns("PCT_MARGEN").Visible = False
            grdProductos.Columns("IMP_COSTO_UNI").Visible = False
        End If
    End If
    'I.CVIERA 21.12.2020 | 06.01.2021
'    grdProductos.Refresh
    'F.CVIERA 21.12.2020
    
Exit Sub
Control:
    Me.txtBuscar.Text = ""
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error_busqueda"
    Me.txtBuscar.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim tempCtrl  As Boolean
    Dim ShiftDown  As Boolean
    
    Call frmPedido.FocusOptiones(KeyCode, Shift)

    ShiftDown = (Shift And vbShiftMask) > 0
    tempCtrl = (Shift And vbCtrlMask) > 0

    Select Case KeyCode
        Case ShiftDown And vbKey6
            Exit Sub
        Case tempCtrl And vbKeyP
            If grdProductos.ApproxCount > 0 And productosenpromocion <> "0" Then
                frm_DLV_PromocionesXproductos.Show vbModal
                grdProductos.SetFocus
            End If
        Case tempCtrl And vbKeyI
            Select Case strgrdFoco
                Case "Producto"
                    If grdProductos.ApproxCount > 0 And grdProductos.Columns("COD_PROD_INK").Value <> "" Then
                        mostrarImagen grdProductos.Columns("COD_PROD_INK").Value, grdProductos.Columns("DES_PRODUCTO_2"), grdProductos.Columns("DESC_LABORATORIO"), grdProductos.Columns("DES_GRUPO"), grdProductos.Columns("DES_CLASE")
                        grdProductos.SetFocus
                    End If
                Case "Alternativo"
                    If grdAlternativos.ApproxCount > 0 And grdAlternativos.Columns("COD_PROD_INK").Value <> "" Then
                        mostrarImagen grdAlternativos.Columns("COD_PROD_INK").Value, grdAlternativos.Columns("DES_PRODUCTO_2"), grdAlternativos.Columns("DESC_LABORATORIO"), grdAlternativos.Columns("DES_GRUPO"), grdAlternativos.Columns("DES_CLASE")
                        grdAlternativos.SetFocus
                    End If
                Case "Complementario"
                    If grdComplementarios.ApproxCount > 0 And grdComplementarios.Columns("COD_PROD_INK").Value <> "" Then
                        mostrarImagen grdComplementarios.Columns("COD_PROD_INK").Value, grdComplementarios.Columns("DES_PRODUCTO_2"), grdComplementarios.Columns("DESC_LABORATORIO"), grdComplementarios.Columns("DES_GRUPO"), grdComplementarios.Columns("DES_CLASE")
                        grdComplementarios.SetFocus
                    End If
            End Select
    End Select

    If ShiftDown And KeyCode = 53 Then Exit Sub

    psub_KeyDownAplicacion KeyCode, Shift
    If KeyCode = 118 Then Exit Sub
    grdProductos.FetchRowStyle = True

    Select Case KeyCode
        Case vbKeyF1
            frm_VTA_Busqueda.txtBuscar.SetFocus
        Case vbKeyF2
            frm_VTA_Busqueda.grdProductos.SetFocus
        Case vbKeyF3
            If frm_VTA_Busqueda.grdAlternativos.Visible Then frm_VTA_Busqueda.grdAlternativos.SetFocus
        Case vbKeyF4
            mostrarAlternativos
        Case vbKeyF9
            If frm_VTA_Busqueda.grdComplementarios.Visible Then frm_VTA_Busqueda.grdComplementarios.SetFocus
        Case vbKeyF11
''''''''''''''''''            'If (objUsuario.EsDelivery = True Or objUsuario.EsQuimico(objUsuario.Codigo) = True) And grdProductos.ApproxCount > 0 Then
''''''''''''''''''            If objUsuario.EsDelivery = True And grdProductos.ApproxCount > 0 Then
''''''''''''''''''                On Error GoTo CtrlErr1
''''''''''''''''''                frm_DLV_Stock_Total.Datos grdProductos.Columns(0).Value, grdProductos.Columns(1).Value
''''''''''''''''''                frm_DLV_Stock_Total.Show vbModal
''''''''''''''''''            Exit Sub
''''''''''''''''''CtrlErr1:
''''''''''''''''''                Err.Raise Err.Number, "", App.FileDescription
''''''''''''''''''            End If
'''''''''''''''''''AGREDADO PARA QUE SEA LA PANTALLA PARA EL LOCAL
''''''''''''''''''            If (objUsuario.EsQuimico(objUsuario.Codigo) = True Or objUsuario.EsTecnico(objUsuario.Codigo) = True) And grdProductos.ApproxCount > 0 Then
''''''''''''''''''                On Error GoTo CtrlErr1
''''''''''''''''''                frm_VTA_Stock_Total.strCodigoProducto = grdProductos.Columns(0).Value & ""
''''''''''''''''''                frm_VTA_Stock_Total.strDescripcionProducto = grdProductos.Columns(1).Value & ""
''''''''''''''''''                frm_VTA_Stock_Total.Show vbModal
''''''''''''''''''            Exit Sub
''''''''''''''''''CtrlErr12:
''''''''''''''''''                Err.Raise Err.Number, "", App.FileDescription
''''''''''''''''''            End If
        Case vbKeyF12
            On Error GoTo CtrlErr
            If frmPedido.grdPedido.ApproxCount > 0 Then
              frmPedido.grdPedido.SetFocus
            End If
            Exit Sub
CtrlErr:
            MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
        Case Else
            
            If Shift = 2 Or Shift = 1 Then KeyCode = 0
    End Select
End Sub

Private Sub Form_LostFocus()
    Set objProducto = Nothing
    frm_VTA_RecetarioM.pstrFlgRM = ""
End Sub

Private Sub grdAlternativos_DblClick()
    If grdProductos.ApproxCount = 0 Then
        Exit Sub
    End If
    If grdAlternativos.ApproxCount = 0 Then
        Exit Sub
    End If
    Dim Precio As Double
    If objUsuario.EsDelivery Then
        Precio = grdAlternativos.Columns(9).Value
    Else
        Precio = grdAlternativos.Columns(3).Value
    End If
    If grdAlternativos.ApproxCount > 0 And Precio > 0 Then
        ''CAMBIAR POR LA NUEVA FUNCION
        If grdAlternativos.Columns(6).Value = "0" Then
            MsgBox "El producto no tiene Stock", vbCritical + vbOKOnly, App.ProductName
            grdAlternativos.SetFocus
            Exit Sub
        End If
        '------------------------------
        strIdFrac = objProducto.ListaDevFracciona(grdAlternativos.Columns(0).Value, objUsuario.CodigoLocal, objVenta.CodModalidadVenta)
        strIndicadorReceta = objProducto.IndicadorReceta(grdAlternativos.Columns(5).Value)
        If objVenta.bk_ServiceType = "RET" And grdAlternativos.Columns(0).Value = "09938" Then
        Else
            frm_VTA_CantidadProducto.subDatos grdAlternativos.Columns(0).Value, grdAlternativos.Columns(1).Value, strTipoPrecio, Label1(2).Caption, Producto_Normal, strIdFrac, strIndicadorReceta, "", "", "", "", "", grdAlternativos.DataSource("DESC_LABORATORIO").Value
        End If
    End If
End Sub

Private Sub grdAlternativos_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
Dim n As Double
Dim s As Double
Dim f As Integer
Dim e As Integer
    
    If Condition = 0 Then
        Select Case Col
               Case 1, 7, 8
                    n = val(grdAlternativos.Columns(7).CellText(Bookmark))
                    s = val(grdAlternativos.Columns(6).CellText(Bookmark))
                    f = InStr(grdAlternativos.Columns(6).CellText(Bookmark), "F")
                    If n > 0 And (s > 0 Or f > 0) Then
                        CellStyle.ForeColor = vbBlue 'vbRed JESCATE
                        CellStyle.Font.Bold = True
                    End If
            Case 2
                    e = val(grdAlternativos.Columns(10).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbRed 'vbBlue JESCATE
                        CellStyle.Font.Bold = True
                    End If
        End Select
    End If
    
    If Condition = 2 Or Condition = 3 Then
        Select Case Col
               Case 1, 7, 8
                    n = val(grdAlternativos.Columns(7).CellText(Bookmark))
                    s = val(grdAlternativos.Columns(6).CellText(Bookmark))
                    f = InStr(grdAlternativos.Columns(6).CellText(Bookmark), "F")
                    If n > 0 And (s > 0 Or f > 0) Then
                        CellStyle.ForeColor = RGB(140, 224, 0) 'vbYellow JESCATE
                        CellStyle.Font.Bold = True
                    End If
               Case 2
                    e = val(grdAlternativos.Columns(10).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = &HFF00&
                        CellStyle.Font.Bold = True
                    End If
        End Select
    End If
    
    If Condition = 2 Or Condition = 3 Then
        Select Case Col
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                If Not objUsuario.EsDelivery Then
                e = val(grdAlternativos.Columns(10).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbYellow 'RGB(140, 224, 0) JESCATE
                        CellStyle.Font.Bold = True
                    End If
                Else
                    e = val(grdAlternativos.Columns(12).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbYellow 'RGB(140, 224, 0) JESCATE
                        CellStyle.Font.Bold = True
                    End If
                End If
        End Select
        
    ElseIf Condition = 0 Then
        Select Case Col
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                If Not objUsuario.EsDelivery Then
                e = val(grdAlternativos.Columns(10).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbRed 'vbBlue JESCATE
                        CellStyle.Font.Bold = True
                    End If
                Else
                    e = val(grdAlternativos.Columns(12).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbRed 'vbBlue JESCATE
                        CellStyle.Font.Bold = True
                    End If
                End If
        End Select
    End If
    
    Dim Valor As String
    Valor = grdAlternativos.Columns(8).CellText(Bookmark)

    If (Valor = "GG") Then
        If Col = 16 Then
            CellStyle.BackColor = RGB(255, 255, 0)
        End If
        CellStyle.ForeColor = vbRed
    ElseIf (Valor = "G") Then
        If Col = 16 Then
            CellStyle.BackColor = RGB(255, 255, 0)
        End If
        CellStyle.ForeColor = vbBlue
    ElseIf (Valor = "3G") Then
    'Verde 87, 166, 57
        If Col = 16 Then
            CellStyle.BackColor = RGB(255, 255, 0)
        End If
        CellStyle.ForeColor = RGB(87, 166, 57)
    
    Else
    CellStyle.ForeColor = vbBlack
    CellStyle.Font.Bold = False
    End If
    grdAlternativos.Styles(5).BackColor = RGB(162, 181, 205)
    
    
End Sub

Private Sub grdAlternativos_GotFocus()
    strgrdFoco = "Alternativo"
End Sub

Private Sub grdAlternativos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim tempCtrl  As Boolean, tempAlt As Boolean
    tempCtrl = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyReturn
            grdAlternativos_DblClick
        Case vbKeyF11
                        If objUsuario.EsDelivery = True And grdAlternativos.ApproxCount > 0 Then
                            'AE 18/08/2014
                            objVenta.LocalAtencion = mdiPrincipal.ctlCliente1.LocalAsignado
                            objVenta.LocalDespacho = mdiPrincipal.ctlCliente1.LocalDespacho
                            frm_DLV_Stock_Total.Datos grdAlternativos.Columns(0).Value, grdAlternativos.Columns(1).Value
                            frm_DLV_Stock_Total.Show vbModal
                        Else
                            If (objUsuario.EsQuimico(objUsuario.Codigo) = True Or objUsuario.EsTecnico(objUsuario.Codigo) = True) And grdProductos.ApproxCount > 0 Then
                                frm_VTA_Stock_Total.strCodigoProducto = grdAlternativos.Columns(0).Value & ""
                                frm_VTA_Stock_Total.strDescripcionProducto = grdAlternativos.Columns(1).Value & ""
                                frm_VTA_Stock_Total.Show vbModal
                            End If
                        End If
'         Case tempCtrl And vbKeyI
'            If grdAlternativos.ApproxCount > 0 And grdAlternativos.Columns("COD_PROD_INK").Value <> "" Then
'                mostrarImagen grdAlternativos.Columns("COD_PROD_INK").Value, grdAlternativos.Columns("DES_PRODUCTO_2"), grdAlternativos.Columns("DESC_LABORATORIO"), grdAlternativos.Columns("DES_GRUPO"), grdAlternativos.Columns("DES_CLASE")
'                grdAlternativos.SetFocus
'            End If
    
    End Select
End Sub
Private Sub showImageWebOrDrive(codProducto As String, Producto As String, Laboratorio As String, grupo As String, clase As String)
    With frm_VTA_ImagenProducto
        .flgWeb = True
        .cMsg = 0
        Let .codigoProducto = codProducto
        Let .Producto = Producto
        Let .Laboratorio = Laboratorio
        Let .grupo = grupo
        Let .clase = clase
        .Show vbModal
    End With
End Sub
Private Sub mostrarImagen(codProducto As String, Producto As String, Laboratorio As String, grupo As String, clase As String)
    Dim imagenFileUrl As String
    Dim urlFolderImagen As String
    MousePointer = vbHourglass
    If codProducto = "" Then
        GoTo salir
    End If
    showImageWebOrDrive codProducto, Producto, Laboratorio, grupo, clase
    GoTo salir
    'el resto de codigo ya no va
    'gstrUriFolderImagen = "V:" '"\\10.100.50.41\e_commerce\MAESTRO DE FOTOS EDITADAS\"
'    If gstrUriFolderImagen = "" Then
'        gstrUriFolderImagen = CStr(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FILE_IMAGEN_PRODUCTO"))
'    End If
'
'    If Dir(gstrUriFolderImagen, vbDirectory) = "" Then
'        MsgBox "La carpeta compartida de imagenes no existe o no es accesible" & vbCrLf & _
'               gstrUriFolderImagen, vbCritical + vbOKOnly, App.ProductName
'    End If
'
'    imagenFileUrl = gstrUriFolderImagen & codProducto & "\"
'
'    If Dir(imagenFileUrl, vbDirectory) <> "" Then
'        If Dir(imagenFileUrl & "\*.png") <> "" Then
'            With frm_VTA_ImagenProducto
'                Let .codigoproducto = codProducto
'                Let .imagePath = imagenFileUrl
'                Let .imagenFileUrl = imagenFileUrl
'                Let .Producto = Producto
'                Let .Laboratorio = Laboratorio
'                Let .grupo = grupo
'                Let .clase = clase
'                .Show vbModal
'            End With
'            'Exit Sub
'            GoTo salir
'        Else
'            GoTo noExiste
''            MsgBox "No se encontraron imagenes para el producto seleccionado" & vbCrLf & _
''                   imagenFileUrl, vbCritical + vbOKOnly, App.ProductName
'        End If
'    Else
'        imagenFileUrl = gstrUriFolderImagen & codProducto & "*.png"
'        If Dir(imagenFileUrl) <> "" Then
'            With frm_VTA_ImagenProducto
'                Let .codigoproducto = codProducto
'                Let .imagePath = gstrUriFolderImagen
'                Let .imagenFileUrl = imagenFileUrl
'                Let .Producto = Producto
'                Let .Laboratorio = Laboratorio
'                Let .grupo = grupo
'                Let .clase = clase
'                .Show vbModal
'            End With
'            'Exit Sub
'            GoTo salir
'        Else
'            GoTo noExiste
''            MsgBox "No se encontraron imagenes para el producto seleccionado" & vbCrLf & _
''                   imagenFileUrl, vbCritical + vbOKOnly, App.ProductName
'        End If
'    End If
salir:
    MousePointer = vbDefault
    Exit Sub
noExiste:
    MousePointer = vbDefault
    MsgBox "No se encontraron imagenes para el producto seleccionado" & vbCrLf & _
                   imagenFileUrl, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub grdAlternativos_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo Control
    
    If strgrdFoco = "Alternativo" Then
        If grdAlternativos.ApproxCount > 0 Then
            If objUsuario.EsDelivery Then
                If Not fraComplementarios.Visible Then Set grdComplementarios.DataSource = objProducto.ProdComplementario(objUsuario.CodigoEmpresa, _
                                                                                                                           mdiPrincipal.ctlCliente1.LocalAsignado, _
                                                                                                                           strTipoPrecio, _
                                                                                                                           grdAlternativos.Columns("COD_PRODUCTO").Value, _
                                                                                                                           objUsuario.CodigoLocal, objVenta.CodigoConvenio)
                grdComplementarios.Refresh
            Else
                If Not fraComplementarios.Visible Then Set grdComplementarios.DataSource = objProducto.ProdComplementario(objUsuario.CodigoEmpresa, _
                                                                                                                           objUsuario.CodigoLocal, _
                                                                                                                           strTipoPrecio, _
                                                                                                                           grdAlternativos.Columns("COD_PRODUCTO").Value, _
                                                                                                                           objUsuario.CodigoLocal, objVenta.CodigoConvenio)
                grdComplementarios.Refresh
            End If
    
            ''26/08/07 agregado por pherrera se cae cuando es recetario magistral
            If objVenta.ptmModalidad <> Recetario Then
                If grdAlternativos.Columns("MSJ_ARG").Value <> "" Then
                    lblMsgComercialAlt.Visible = True
                    ''LblMsgConercial(0).BackColor = &HFF8080
                    lblMsgComercialAlt.Caption = grdAlternativos.Columns("MSJ_ARG").Value
                    'lblMsgComercialAlt.Caption = "[PROD]-> Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged."
                Else
                    lblMsgComercialAlt.Visible = False
                    lblMsgComercialAlt.Caption = ""
                End If
            End If
        Else
            If grdComplementarios.ApproxCount <= 0 Then
                grdComplementarios.Limpiar
            End If
        End If
        
        
        
    End If
    Exit Sub
Control:
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub


Private Sub grdComplementarios_DblClick()
    If grdProductos.ApproxCount = 0 Then
        Exit Sub
    End If
     If grdComplementarios.ApproxCount > 0 And grdComplementarios.Columns(3).Value > 0 Then
        ''CAMBIAR POR LA NUEVA FUNCION
        If grdComplementarios.Columns(6).Value = "0" Then
            MsgBox "El producto no tiene Stock", vbCritical + vbOKOnly, App.ProductName
            grdComplementarios.SetFocus
            Exit Sub
        End If
        '------------------------------
        strIdFrac = objProducto.ListaDevFracciona(grdComplementarios.Columns(0).Value, objUsuario.CodigoLocal, objVenta.CodModalidadVenta)
        strIndicadorReceta = objProducto.IndicadorReceta(grdComplementarios.Columns(5).Value)
        If objVenta.bk_ServiceType = "RET" And grdComplementarios.Columns(0).Value = "09938" Then
        Else
            frm_VTA_CantidadProducto.subDatos grdComplementarios.Columns(0).Value, grdComplementarios.Columns(1).Value, strTipoPrecio, Label1(4).Caption, Producto_Normal, strIdFrac, strIndicadorReceta, "", "", "", "", "", grdComplementarios.DataSource("DESC_LABORATORIO").Value
        End If
    End If
End Sub

Private Sub grdComplementarios_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
Dim n As Double
Dim s As Double
Dim f As Integer
Dim e As Integer
    If Condition = 0 Then
        Select Case Col
            Case 1, 7, 8
               n = val(grdComplementarios.Columns(7).CellText(Bookmark))
               s = val(grdComplementarios.Columns(6).CellText(Bookmark))
               f = InStr(grdComplementarios.Columns(6).CellText(Bookmark), "F")
               If n > 0 And (s > 0 Or f > 0) Then
                CellStyle.ForeColor = vbBlue 'vbRed JESCATE
                CellStyle.Font.Bold = True
               End If
            Case 2
                e = val(grdComplementarios.Columns(10).CellText(Bookmark))
                If e > 0 Then
                    CellStyle.ForeColor = vbRed ' vbBlue JESCATE
                    CellStyle.Font.Bold = True
                End If
        End Select
    End If
    
    If Condition = 2 Or Condition = 3 Then
        Select Case Col
            Case 1, 7, 8
               n = val(grdComplementarios.Columns(7).CellText(Bookmark))
               s = val(grdComplementarios.Columns(6).CellText(Bookmark))
               f = InStr(grdComplementarios.Columns(6).CellText(Bookmark), "F")
               If n > 0 And (s > 0 Or f > 0) Then
                CellStyle.ForeColor = vbYellow
                CellStyle.Font.Bold = True
               End If
            Case 2
                e = val(grdComplementarios.Columns(10).CellText(Bookmark))
                If e > 0 Then
                    CellStyle.ForeColor = &HFF00&
                    CellStyle.Font.Bold = True
                End If
        End Select
    End If
    
    Dim Valor As String
    Valor = grdAlternativos.Columns(8).CellText(Bookmark)
    
    If (Valor = "GG") Then
    CellStyle.ForeColor = vbRed
    ElseIf (Valor = "G") Then
    CellStyle.ForeColor = vbBlue
    ElseIf (Valor = "3G") Then
    'Naranja
    'Verde 87, 166, 57
    CellStyle.ForeColor = RGB(87, 166, 57)
    Else
    CellStyle.ForeColor = vbBlack
    CellStyle.Font.Bold = False
    End If

    grdComplementarios.Styles(5).BackColor = RGB(162, 181, 205)
    
End Sub

Private Sub grdComplementarios_GotFocus()
    strgrdFoco = "Complementario"
End Sub

Private Sub grdComplementarios_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim tempCtrl  As Boolean
    tempCtrl = (Shift And vbCtrlMask) > 0
    Select Case KeyCode
        Case vbKeyReturn
            grdComplementarios_DblClick
        
'         Case tempCtrl And vbKeyI
'            If grdComplementarios.ApproxCount > 0 And grdComplementarios.Columns("COD_PROD_INK").Value <> "" Then
'                mostrarImagen grdComplementarios.Columns("COD_PROD_INK").Value, grdComplementarios.Columns("DES_PRODUCTO_2"), grdComplementarios.Columns("DESC_LABORATORIO"), grdComplementarios.Columns("DES_GRUPO"), grdComplementarios.Columns("DES_CLASE")
'                grdComplementarios.SetFocus
'            End If
        
    End Select
End Sub

Private Sub grdProductos_DblClick()
    On Error GoTo CtrlErr
    If grdProductos.ApproxCount = 0 Then
        Exit Sub
    End If
    If LenB(grdProductos.DataSource("COD_CODIGO_SAP").Value & "") = 0 Then
        MsgBox "El producto (" & grdProductos.Columns("COD_PRODUCTO").Value & ") no se encuentra homologado", _
                vbCritical + vbOKOnly, App.ProductName
        grdProductos.SetFocus
        Exit Sub
    End If
    If grdProductos.ApproxCount > 0 And grdProductos.Columns("COD_PRODUCTO").Value <> "" Then
        ''CAMBIAR POR LA NUEVA FUNCION
        ''IF grdProductos.Columns(3).Value > 0
        If objVenta.ptmModalidad <> Guias_Remision And objUsuario.TipoMaquina <> objUsuario.TipoMaquinaCabina Then
          If Trim(grdProductos.Columns(3).Value) = 0 Then Exit Sub
            If grdProductos.Columns(6).Value = "0" Then
                If objProducto.FnEsModalidad_Recetario(grdProductos.Columns("COD_PRODUCTO").Value) > 0 Then
                Else
                    MsgBox "El producto (" & grdProductos.Columns("COD_PRODUCTO").Value & ") no tiene Stock", vbCritical + vbOKOnly, App.ProductName
                    grdProductos.SetFocus
                    Exit Sub
                End If
            End If
          ''End If
        
        End If
        '------------------------------
        
        If objProducto.FnEsModalidad_Recetario(grdProductos.Columns("COD_PRODUCTO").Value) > 0 Then
            '''frm_VTA_RecetarioM.pstrFlgRM = "1"
            frm_VTA_RecetarioM.Show
            frm_VTA_RecetarioM.SetFocus
        Else
            strIdFrac = objProducto.ListaDevFracciona(grdProductos.Columns("COD_PRODUCTO").Value, objUsuario.CodigoLocal, objVenta.CodModalidadVenta)
            strIndicadorReceta = objProducto.IndicadorReceta(grdProductos.Columns(5).Value)
            frm_VTA_CantidadProducto.flgEspecieValorada = grdProductos.DataSource("FLG_ESP_VAL").Value
            If objVenta.bk_ServiceType = "RET" And grdProductos.Columns(0).Value = "09938" Then
            Else
                frm_VTA_CantidadProducto.subDatos "" & grdProductos.Columns("COD_PRODUCTO").Value, _
                                                  "" & grdProductos.Columns(1).Value, _
                                                  strTipoPrecio, _
                                                  "" & Label1(0).Caption, _
                                                  Producto_Normal, _
                                                  strIdFrac, _
                                                  strIndicadorReceta, _
                                                  "" & grdProductos.Columns(5).Value, _
                                                  "" & grdProductos.DataSource("STOCK").Value, _
                                                  "", _
                                                  "", _
                                                  "" & grdProductos.DataSource("COD_CODIGO_SAP").Value, _
                                                  "" & grdProductos.DataSource("DESC_LABORATORIO").Value
            End If
            txtBuscar.SetFocus
        End If
        '------------------------------

    End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
    grdProductos.SetFocus
End Sub

Private Sub grdProductos_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
Dim n As Double
Dim s As Double
Dim f As Integer
Dim e As Integer
    If Condition = 0 Then
        Select Case Col
            Case 1, 7, 8
               ''If Not objUsuario.EsDelivery Then
                    n = val(grdProductos.Columns(7).CellText(Bookmark))
                    If n > 0 Then
                        CellStyle.ForeColor = vbBlue 'vbRed JESCATE
                        CellStyle.Font.Bold = True
                    End If

''''''                    'Se comento por que no se debe validar si tiene o no stock
''''''                    'Jahzeel López
''''''                    '12/11/2007
''''''                    'N = Val(grdProductos.Columns(7).CellText(bookmark))
''''''                    's = Val(grdProductos.Columns(6).CellText(bookmark))
''''''                    'f = InStr(grdProductos.Columns(6).CellText(bookmark), "F")
''''''                    'If N > 0 And (s > 0 Or f > 0) Then
''''''                    ' CellStyle.ForeColor = vbRed
''''''                    ' CellStyle.Font.Bold = True
''''''                    'End If

               ''End If
                    
            Case 2
                If Not objUsuario.EsDelivery Then
                    e = val(grdProductos.Columns(10).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbBlue 'vbRed JESCATE
                        CellStyle.Font.Bold = True
                    End If
                Else
                    e = val(grdProductos.Columns(12).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbBlue 'vbRed JESCATE
                        CellStyle.Font.Bold = True
                    End If
                End If
        End Select

    End If
    
''    If Condition = 1 Then
''        Select Case Col
''            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
''                CellStyle.ForeColor = vbGreen
''            End Select
''    Else
''         ' CellStyle.ForeColor = vbBlue
''    End If
        
    If Condition = 2 Or Condition = 3 Then
        Select Case Col
            Case 1, 7, 8
               ''If Not objUsuario.EsDelivery Then
                    n = val(grdProductos.Columns(7).CellText(Bookmark))
                    If n > 0 Then
                        CellStyle.ForeColor = RGB(140, 224, 0) 'vbYellow JESCATE
                        CellStyle.Font.Bold = True
                    End If
                    
                    ''Se comento por que no se debe validar si tiene o no stock
                    ''Jahzeel López
                    ''12/11/2007
                    ''N = Val(grdProductos.Columns(7).CellText(bookmark))
                    ''s = Val(grdProductos.Columns(6).CellText(bookmark))
                    ''f = InStr(grdProductos.Columns(6).CellText(bookmark), "F")
                    ''If N > 0 And (s > 0 Or f > 0) Then
                    '' CellStyle.ForeColor = vbYellow
                    '' CellStyle.Font.Bold = True
                    ''End If
               ''End If
''            Case 1, 2, 9
''                If Not objUsuario.EsDelivery Then
''                e = Val(grdProductos.Columns(10).CellText(bookmark))
''                    If e > 0 Then
''                        CellStyle.ForeColor = &HFF00&
''                        CellStyle.Font.Bold = True
''                    End If
''                Else
''                    e = Val(grdProductos.Columns(12).CellText(bookmark))
''                    If e > 0 Then
''                        CellStyle.ForeColor = &HFF00&
''                        CellStyle.Font.Bold = True
''                    End If
''                End If
        End Select
    End If
    
    If Condition = 2 Or Condition = 3 Then
        Select Case Col
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9
                If Not objUsuario.EsDelivery Then
                    e = val(grdProductos.Columns(10).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbYellow 'RGB(140, 224, 0) JESCATE
                        CellStyle.Font.Bold = True
                    End If
                Else
                    e = val(grdProductos.Columns(12).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbYellow ' RGB(140, 224, 0) JESCATE
                        CellStyle.Font.Bold = True
                    End If
                End If
        End Select
        
    ElseIf Condition = 0 Then
        Select Case Col
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9
                If Not objUsuario.EsDelivery Then
                e = val(grdProductos.Columns(10).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbRed 'vbBlue JESCATE
                        CellStyle.Font.Bold = True
                    End If
                Else
                    e = val(grdProductos.Columns(12).CellText(Bookmark))
                    If e > 0 Then
                        CellStyle.ForeColor = vbRed 'vbBlue JESCATE
                        CellStyle.Font.Bold = True
                    End If
                End If
        End Select
    End If

Dim Valor As String
Valor = grdProductos.Columns(8).CellText(Bookmark)
If (Valor = "GG") Then
    If Col = 8 Then
        CellStyle.BackColor = RGB(255, 255, 0)
    End If
    CellStyle.ForeColor = vbRed
ElseIf (Valor = "G") Then
    If Col = 8 Then
        CellStyle.BackColor = RGB(255, 255, 0)
    End If
    CellStyle.ForeColor = vbBlue
ElseIf (Valor = "3G") Then
    'Naranja 244, 70, 17
    'Verde 87, 166, 57
    If Col = 8 Then
        CellStyle.BackColor = RGB(255, 255, 0)
    End If
    CellStyle.ForeColor = RGB(87, 166, 57)
Else
    CellStyle.ForeColor = vbBlack
    CellStyle.Font.Bold = False
End If

If (grdProductos.Columns("PROM").CellText(Bookmark) = "P") Then
    If Col = 21 Then
        CellStyle.BackColor = RGB(255, 255, 0)
    End If
End If

grdProductos.Styles(5).BackColor = RGB(162, 181, 205)


End Sub

Private Sub grdProductos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
On Error GoTo handle
''                If Val(grdPedidoStock.Columns(6).CellText(bookmark)) < 0 Then
''                    RowStyle.BackColor = RGB(240, 128, 128)
''                End If
        ''If Split = 0 Then
''            If (Val(grdProductos.Columns(7).CellText(bookmark)) < 0) Or (Val(grdProductos.Columns(8).CellText(bookmark)) < 0) Or (Val(grdProductos.Columns(10).CellText(bookmark)) < 0) Then
''                RowStyle.ForeColor = RGB(0, 213, 0)
''              Else
''                RowStyle.ForeColor = RGB(77, 77, 196)
''            End If
''        'End If
        
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdProductos_GotFocus()
    strgrdFoco = "Producto"
End Sub

''Private Sub grdProductos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
''Dim s As Double
''Dim f As Integer
''
''On Error GoTo handle
''                s = Val(grdProductos.Columns(6).CellText(Bookmark))
''                f = InStr(grdProductos.Columns(6).CellText(Bookmark), "F")
''               If s = 0 And f = 0 Then
''                'RowStyle.ForeColor = &HC0E0FF
''                'RowStyle.BackColor = &HE0E0E0
''                'RowStyle.Font.Bold = True
''               End If
''Exit Sub
''handle:
''    MsgBox Err.Description, vbCritical, App.ProductName
''End Sub

Private Sub grdProductos_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim ShiftDown  As Boolean
    Dim tempCtrl  As Boolean
    
    tempCtrl = (Shift And vbCtrlMask) > 0
    
    If grdProductos.ApproxCount = 0 Then
        Exit Sub
    End If
  
    Let frm_DLV_PromocionesXproductos.codigoProducto = grdProductos.Columns("COD_PRODUCTO").Value
   
    productosenpromocion = frm_DLV_PromocionesXproductos.GetValidaCantidadPromocion
     
    If productosenpromocion = "0" Then
        labelProductosEnPromocion.Item(1).Visible = False
        labelProductosEnPromocionDesc.Item(1).Visible = False
    Else
        labelProductosEnPromocion.Item(1).Visible = True
        labelProductosEnPromocionDesc.Item(1).Visible = True
    End If

    Select Case KeyCode
        Case vbKeyReturn
            grdProductos_DblClick
                       
        Case vbKeyF10
            InsertarVentaPerdida
            frm_VTA_Busqueda.grdProductos.SetFocus
        
        Case vbKeyF11
            If objUsuario.EsDelivery = True And grdProductos.ApproxCount > 0 Then
                ''AE 18/08/2014
                objVenta.LocalAtencion = mdiPrincipal.ctlCliente1.LocalAsignado
                objVenta.LocalDespacho = mdiPrincipal.ctlCliente1.LocalDespacho
                frm_DLV_Stock_Total.Datos grdProductos.Columns("COD_PRODUCTO").Value, grdProductos.Columns(1).Value
                frm_DLV_Stock_Total.Show vbModal
            Else
                ''If (objUsuario.EsQuimico(objUsuario.Codigo) = True Or objUsuario.EsTecnico(objUsuario.Codigo) = True) And grdProductos.ApproxCount > 0 Then
                If objUsuario.PuedeVerStock(objUsuario.Codigo) = True And grdProductos.ApproxCount > 0 Then
                    frm_VTA_Stock_Total.strCodigoProducto = grdProductos.Columns("COD_PRODUCTO").Value & ""
                    frm_VTA_Stock_Total.strDescripcionProducto = grdProductos.Columns(1).Value & ""
                    frm_VTA_Stock_Total.Show vbModal
                End If
            End If
                        
'        Case tempCtrl And vbKeyI
'            If grdProductos.ApproxCount > 0 And grdProductos.Columns("COD_PROD_INK").Value <> "" Then
'                mostrarImagen grdProductos.Columns("COD_PROD_INK").Value, grdProductos.Columns("DES_PRODUCTO_2"), grdProductos.Columns("DESC_LABORATORIO"), grdProductos.Columns("DES_GRUPO"), grdProductos.Columns("DES_CLASE")
'                grdProductos.SetFocus
'            End If
    
    End Select
End Sub

Private Sub grdProductos_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo Control

    If grdProductos.ApproxCount > 0 Then
        If objUsuario.EsDelivery Then
        
            Let frm_DLV_PromocionesXproductos.codigoProducto = grdProductos.Columns("COD_PRODUCTO").Value
   
            productosenpromocion = frm_DLV_PromocionesXproductos.GetValidaCantidadPromocion
     
            If productosenpromocion = "0" Then
     
                labelProductosEnPromocion.Item(1).Visible = False
                labelProductosEnPromocionDesc.Item(1).Visible = False
     
            End If
     
            If productosenpromocion <> "0" Then
       
                labelProductosEnPromocion.Item(1).Visible = True
                labelProductosEnPromocionDesc.Item(1).Visible = True
     
            End If
        
            'If Not fraAlternativos.Visible Then Set grdAlternativos.DataSource = objProducto.Lista(objUsuario.CodigoEmpresa, mdiPrincipal.ctlCliente1.LocalAsignado, strTipoPrecio, txtBuscar.Text, grdProductos.Columns(4).Value, grdProductos.Columns("COD_PRODUCTO").Value, objUsuario.CodigoLocal)
            
            ' JCT 09-ABR-12, si es Mi Farma en local no usar DLV, si no local despacho
            'MsgBox "mdiPrincipal.ctlCliente1.LocalDespacho: " + mdiPrincipal.ctlCliente1.LocalDespacho
            If Not fraAlternativos.Visible Then Set grdAlternativos.DataSource = objProducto.Lista(objUsuario.CodigoEmpresa, mdiPrincipal.ctlCliente1.LocalAsignado, strTipoPrecio, txtBuscar.Text, grdProductos.Columns(4).Value, grdProductos.Columns("COD_PRODUCTO").Value, _
                                                                                                mdiPrincipal.ctlCliente1.LocalDespacho, "", "", mdiPrincipal.ctlCliente1.sCia)
            
''''            If Not fraComplementarios.Visible Then Set grdComplementarios.DataSource = objProducto.Lista(objUsuario.CodigoEmpresa, mdiPrincipal.ctlCliente1.LocalAsignado, strTipoPrecio, txtBuscar.Text, grdProductos.Columns(4).Value, "", objUsuario.CodigoLocal)
''''    05/05/09 Modificado x ccieza para ver listado de productos complementarios
            
            If strgrdFoco = "Producto" Then
'''                If Not fraComplementarios.Visible Then Set grdComplementarios.DataSource = objProducto.ProdComplementario(objUsuario.CodigoEmpresa, _
'''                                                                                                                          mdiPrincipal.ctlCliente1.LocalAsignado, _
'''                                                                                                                          strTipoPrecio, _
'''                                                                                                                          grdProductos.Columns("COD_PRODUCTO").Value, _
'''                                                                                                                           objUsuario.CodigoLocal, objVenta.codigoConvenio)
                 ' JCT
                 If Not fraComplementarios.Visible Then Set grdComplementarios.DataSource = objProducto.ProdComplementario(objUsuario.CodigoEmpresa, _
                                                                                                                          mdiPrincipal.ctlCliente1.LocalAsignado, _
                                                                                                                          strTipoPrecio, _
                                                                                                                          grdProductos.Columns("COD_PRODUCTO").Value, _
                                                                                                                          mdiPrincipal.ctlCliente1.LocalDespacho, _
                                                                                                                          objVenta.CodigoConvenio, _
                                                                                                                         mdiPrincipal.ctlCliente1.sCia)
                
                grdComplementarios.Refresh
            End If
        Else
            'If Not fraAlternativos.Visible Then Set grdAlternativos.DataSource = objProducto.Lista(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strTipoPrecio, txtBuscar.Text, grdProductos.Columns(4).Value, grdProductos.Columns("COD_PRODUCTO").Value, objUsuario.CodigoLocal)
            
         'JCT 09-ABR-12, si es Mi Farma en local no usar DLV, si no local despacho
         'MsgBox "mdiPrincipal.ctlCliente1.LocalDespacho: " + mdiPrincipal.ctlCliente1.LocalDespacho
         If Not fraAlternativos.Visible Then Set grdAlternativos.DataSource = objProducto.Lista(mdiPrincipal.ctlCliente1.sCia, objUsuario.CodigoLocal, strTipoPrecio, txtBuscar.Text, grdProductos.Columns(4).Value, grdProductos.Columns("COD_PRODUCTO").Value, objUsuario.CodigoLocal, objUsuario.CodigoEmpresa)

''''            If Not fraComplementarios.Visible Then Set grdComplementarios.DataSource = objProducto.Lista(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strTipoPrecio, txtBuscar.Text, grdProductos.Columns(4).Value, "", objUsuario.CodigoLocal)
''''    05/05/09 Modificado x ccieza para ver listado de productos complementarios
            
            If strgrdFoco = "Producto" Then
                
                If Not fraComplementarios.Visible Then Set grdComplementarios.DataSource = objProducto.ProdComplementario(objUsuario.CodigoEmpresa, _
                                                                                                                          objUsuario.CodigoLocal, _
                                                                                                                          strTipoPrecio, _
                                                                                                                          grdProductos.Columns("COD_PRODUCTO").Value, _
                                                                                                                          objUsuario.CodigoLocal, _
                                                                                                                          objVenta.CodigoConvenio, _
                                                                                                                          objUsuario.CodigoEmpresa)
                grdComplementarios.Refresh
            End If
        End If
        
        ''26/08/07 agregado por pherrera se cae cuando es recetario magistral
        If objVenta.ptmModalidad <> Recetario Then
            If grdProductos.Columns("MSG_COMERCIAL").Value <> "" Then
                LblMsgComercial(0).Visible = True
                ''LblMsgConercial(0).BackColor = &HFF8080
                LblMsgComercial(0).Caption = grdProductos.Columns("MSG_COMERCIAL").Value
              Else
                LblMsgComercial(0).Visible = False
                LblMsgComercial(0).Caption = ""
            End If
        End If
        
        If grdAlternativos.ApproxCount > 0 Then
            grdAlternativos.row = 0
            If grdAlternativos.Columns("MSJ_ARG").Value <> "" Then
                lblMsgComercialAlt.Visible = True
                lblMsgComercialAlt.Caption = grdAlternativos.Columns("MSJ_ARG").Value
            Else
                lblMsgComercialAlt.Visible = False
                lblMsgComercialAlt.Caption = ""
            End If
        Else
            lblMsgComercialAlt.Visible = False
            lblMsgComercialAlt.Caption = ""
        End If
    Else
        grdAlternativos.Limpiar
        grdComplementarios.Limpiar
    End If

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub imgComunicandonos_Click()
On Error GoTo Control
If Trim(pstrPaginaComunicandonos) = "" Then
    MsgBox "No se ha cargado ningún valor en la constante.", vbOKOnly + vbExclamation, "Advertencia"
    Exit Sub
End If
objUsuario.GrabaVisita Me.name, "092"
ShellExecute Me.hwnd, "open", pstrPaginaComunicandonos, vbNullString, "c:\", ByVal 1&

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Private Sub txtBuscar_Change()
    If Len(Trim(txtBuscar.Text)) <= 0 Then LblMsgComercial(0).Visible = False
    objVenta.strBusqueda = txtBuscar.Text
End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
    If EsEscaneado(KeyCode, Shift) = True Then
        objVenta.EscaneoTarjeta = True
    Else
        objVenta.EscaneoTarjeta = False
    End If
End Sub

Private Sub txtBuscar_LostFocus()
    txtBuscar.TABAuto = False
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    On Error GoTo CtrlErr
    If KeyAscii = 13 Then
        nVeces = nVeces + 1
    End If
    
    If nVeces = 1 Or nVeces = 2 Then
        'If FN_TARJETA_CMR(Trim(txtBuscar.Text)) <> "" And esTarjeta = False And KeyAscii = 13 Then
        If FN_TARJETA_CMR(Trim(txtBuscar.Text)) <> "" And KeyAscii = 13 Then
            txtBuscar.Text = Replace(FN_TARJETA_CMR(Trim(txtBuscar.Text)), "%", "")
            cmdBuscar_Click
        Else
            If KeyAscii = 13 Then
                cmdBuscar_Click
            End If
            nVeces = 0
        End If
    Else
        If KeyAscii = 13 Then
            cmdBuscar_Click
        End If
    End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
End Sub

Private Function FN_TARJETA_CMR(Texto As String)
'If InStr(Texto, "&") Then
'    FN_TARJETA_CMR = Replace(fncGuion(Texto, 0, "&"), "%", "")
'End If

Dim NumTarjeta As String
NumTarjeta = Trim(Texto)
    If Mid(NumTarjeta, 1, 2) = "%B" And Mid(NumTarjeta, 19, 1) = "&" Then
        FN_TARJETA_CMR = Mid(NumTarjeta, 3, 16)
        Exit Function
    End If
    If Mid(NumTarjeta, 1, 1) = "%" And Mid(NumTarjeta, 18, 1) = "&" Then
        FN_TARJETA_CMR = Mid(NumTarjeta, 2, 16)
        Exit Function
    End If
    
    If Mid(NumTarjeta, 1, 1) = "Ñ" And Mid(NumTarjeta, 18, 1) = "¿" Then
        FN_TARJETA_CMR = Mid(NumTarjeta, 2, 16)
        Exit Function
    End If
    FN_TARJETA_CMR = ""
End Function

Public Sub Datos(ByVal pstrTipoPrecio As String)
    strTipoPrecio = pstrTipoPrecio
    txtBuscar.Text = IIf(Len(Trim(objVenta.strBusqueda)) = 0, "", objVenta.strBusqueda)
    cmdBuscar_Click
    frmPedido.ReCalculaPrecio strTipoPrecio
    Me.Show
End Sub

Private Sub Format_Grilla()
    Dim bolMuestraStock As Boolean
    Dim columna As TrueDBGrid70.Column

On Error GoTo CtrlErr

    bolMuestraStock = IIf(objUsuario.MostrarStock(objUsuario.Perfil) = "1", True, False)

    If objUsuario.EsDelivery = False Then

        arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LINEA", "PRECIO", "ASO_SUSTITUTO", "COD_INDICADOR_RECETA", "STOCK", "IMP_COMI_UND", "IMP_COMI_FRA", "PRECIO", "FLG_EXCLUSIVO", "MSG_COMERCIAL", "FLG_FARMACO")
        arrCaption = Array("Codigo", "Producto", "Linea", "Precio", "Sustituto", "CodIndicador", "Stock", "Comisi", "Com.F", "Precio Final", "Exclusivo", "Mensaje", "Farmaco")
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgRight, dbgGeneral, dbgGeneral, dbgGeneral, dbgRight, dbgRight, dbgRight, dbgRight, dbgLeft, dbgLeft, dbgCenter)
        If Not bolMuestraStock Then
            'arrAncho = Array(0, 4100, 800, 0, 0, 0, 550, 550, 550, 900, 0, 2000, 500)
            arrAncho = Array(0, 6100, 800, 0, 0, 0, 550, 550, 550, 900, 0, 2000, 500)
        Else
            'arrAncho = Array(0, 3250, 800, 0, 0, 0, 850, 550, 550, 900, 0, 2000, 500)
            arrAncho = Array(0, 4250, 800, 0, 0, 0, 850, 550, 550, 900, 0, 2000, 500)
        End If
        grdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdProductos.Columns(0).Visible = False
        grdProductos.Columns(0).AllowSizing = False
        grdProductos.Columns(1).WrapText = True
        grdProductos.Columns(3).Visible = False
        grdProductos.Columns(4).Visible = False
        grdProductos.Columns(5).Visible = False
        grdProductos.Columns(6).Visible = bolMuestraStock
        grdProductos.Columns(10).Visible = False
        grdProductos.Columns(12).Visible = False
        grdProductos.Columns("MSG_COMERCIAL").Visible = False
        grdProductos.Columns(4).AllowSizing = False
        grdProductos.Columns(0).FetchStyle = True
        grdProductos.Columns(1).FetchStyle = True
        grdProductos.Columns(2).FetchStyle = True
        grdProductos.Columns(3).FetchStyle = True
        grdProductos.Columns(6).FetchStyle = True
        grdProductos.Columns(7).FetchStyle = True
        grdProductos.Columns(8).FetchStyle = True
        grdProductos.Columns(9).FetchStyle = True
        grdProductos.Columns(10).FetchStyle = True
        grdProductos.Columns(11).FetchStyle = True
        ''grdProductos.Columns(12).FetchStyle = True
        grdProductos.Columns(3).NumberFormat = "##,##0.00"
        grdProductos.Columns(7).NumberFormat = "##0.00"
        grdProductos.Columns(8).NumberFormat = "##0.00"
        grdProductos.RowDividerStyle = 0
         
        For Each columna In grdProductos.Columns
            columna.AllowSizing = False
        Next
            
        arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LINEA", "PRECIO", "ASO_SUSTITUTO", "COD_INDICADOR_RECETA", "STOCK", "IMP_COMI_UND", "IMP_COMI_FRA", "PRECIO", "FLG_EXCLUSIVO", "MSG_COMERCIAL", "FLG_FARMACO")
        arrCaption = Array("Codigo", "Producto", "Linea", "Precio", "Sustituto", "CodIndicador", "Stock", "Comisi", "Com.F", "Precio Final", "Exclusivo", "Mensaje", "Farmaco")
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgRight, dbgGeneral, dbgGeneral, dbgGeneral, dbgRight, dbgRight, dbgRight, dbgRight, dbgLeft, dbgCenter)
        If Not bolMuestraStock Then
            arrAncho = Array(0, 4100, 800, 0, 0, 0, 550, 550, 550, 900, 0, 2000, 500)
        Else
            arrAncho = Array(0, 3250, 800, 0, 0, 0, 850, 550, 550, 900, 0, 2000, 500)
        End If

        grdAlternativos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdAlternativos.Columns(0).Visible = False
        grdAlternativos.Columns(0).AllowSizing = False
        grdAlternativos.Columns(1).WrapText = True
        grdAlternativos.Columns(3).Visible = False
        grdAlternativos.Columns(4).Visible = False
        grdAlternativos.Columns(5).Visible = False
        grdAlternativos.Columns(6).Visible = bolMuestraStock
        grdAlternativos.Columns(10).Visible = False
        grdAlternativos.Columns(11).Visible = False
        grdAlternativos.Columns(12).Visible = False
        grdAlternativos.Columns(4).AllowSizing = False
        grdAlternativos.Columns(0).FetchStyle = True
        grdAlternativos.Columns(1).FetchStyle = True
        grdAlternativos.Columns(2).FetchStyle = True
        grdAlternativos.Columns(3).FetchStyle = True
        grdAlternativos.Columns(6).FetchStyle = True
        grdAlternativos.Columns(7).FetchStyle = True
        grdAlternativos.Columns(8).FetchStyle = True
        grdAlternativos.Columns(9).FetchStyle = True
        grdAlternativos.Columns(10).FetchStyle = True
        grdAlternativos.Columns(11).FetchStyle = True
        grdAlternativos.Columns(3).NumberFormat = "##,##0.00"
        grdAlternativos.Columns(7).NumberFormat = "##0.00"
        grdAlternativos.Columns(8).NumberFormat = "##0.00"
        grdAlternativos.RowDividerStyle = 0
        
        For Each columna In grdAlternativos.Columns
            columna.AllowSizing = False
        Next

        grdComplementarios.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdComplementarios.Columns(0).Visible = False
        grdComplementarios.Columns(0).AllowSizing = False
        grdComplementarios.Columns(1).WrapText = True
        grdComplementarios.Columns(3).Visible = False
        grdComplementarios.Columns(4).Visible = False
        grdComplementarios.Columns(5).Visible = False
        grdComplementarios.Columns(6).Visible = bolMuestraStock
        grdComplementarios.Columns(10).Visible = False
        grdComplementarios.Columns(11).Visible = False
        grdComplementarios.Columns(12).Visible = False
        grdComplementarios.Columns(4).AllowSizing = False
        grdComplementarios.Columns(0).FetchStyle = True
        grdComplementarios.Columns(1).FetchStyle = True
        grdComplementarios.Columns(2).FetchStyle = True
        grdComplementarios.Columns(3).FetchStyle = True
        grdComplementarios.Columns(6).FetchStyle = True
        grdComplementarios.Columns(7).FetchStyle = True
        grdComplementarios.Columns(8).FetchStyle = True
        grdComplementarios.Columns(3).NumberFormat = "##,##0.00"
        grdComplementarios.Columns(7).NumberFormat = "##0.00"
        grdComplementarios.Columns(8).NumberFormat = "##0.00"
        grdComplementarios.RowDividerStyle = 0
            
        For Each columna In grdComplementarios.Columns
            columna.AllowSizing = False
        Next
    Else
        ''##agregado por pherrera 25 10 07
        grdProductos.RowHeight = 0
        grdProductos.RowHeight = grdProductos.RowHeight * 1.8
        grdProductos.AlternatingRowStyle = True
        grdProductos.Styles(6).BackColor = &HF1F1F1
        ''##

        ''si es delivery
        'I.CVIERA 31.12.2020
        'arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO_2", "DES_LINEA", "PRC_KAIROS", "ASO_SUSTITUTO", "COD_INDICADOR_RECETA", "STOCK", "IMP_COMI_UND", "IMP_COMI_FRA", "PRECIO", "EST_PRODUCTO", "STOCK", "FLG_EXCLUSIVO", "MSG_COMERCIAL", "STOCK_CAD", "COD_CODIGO_SAP", "PCT_COMI", "COD_PROD_INK", "DES_GRUPO", "DES_CLASE", "DESC_LABORATORIO", "PROM")
        'arrCaption = Array("Codigo", "Producto", "Linea", "Kairos", "Sustituto", "CodIndicador", "Stock", "Estado", "Com.G.", "Publico", "Est", "Local", "flg_exclusivo", "Mensaje", "Cadena", "Codigo SAP", "Comision", "COD_PROD_INK", "Grupo", "Clase", "Laboratorio", "Prom")
        'arrAncho = Array(0, 5500, 2000, 0, 0, 0, 0, 800, 700, 700, 0, 800, 0, 0, 800, 0, 600, 0, 0, 0, 0, 500)
        'arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgRight, dbgGeneral, dbgGeneral, dbgGeneral, dbgCenter, dbgCenter, dbgRight, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgRight, dbgGeneral, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
        arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO_2", "DES_LINEA", "PRC_KAIROS", "ASO_SUSTITUTO", "COD_INDICADOR_RECETA", "STOCK", "IMP_COMI_UND", "IMP_COMI_FRA", "PRECIO", "EST_PRODUCTO", "STOCK", "FLG_EXCLUSIVO", "MSG_COMERCIAL", "STOCK_CAD", "COD_CODIGO_SAP", "PCT_COMI", "COD_PROD_INK", "DES_GRUPO", "DES_CLASE", "DESC_LABORATORIO", "PROM", "FLG_SEG")
        arrCaption = Array("Codigo", "Producto", "Linea", "Kairos", "Sustituto", "CodIndicador", "Stock", "Estado", "Com.G.", "Publico", "Est", "Local", "flg_exclusivo", "Mensaje", "Cadena", "Codigo SAP", "Comision", "COD_PROD_INK", "Grupo", "Clase", "Laboratorio", "Prom", "Flg_Seg")
        'arrAncho = Array(550, 3150, 790, 0, 0, 0, 0, 600, 0, 700, 0, 550, 400, 2000, 600, 0)
        arrAncho = Array(0, 5500, 2000, 0, 0, 0, 0, 800, 700, 700, 0, 800, 0, 0, 800, 0, 600, 0, 0, 0, 0, 500, 500)
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgRight, dbgGeneral, dbgGeneral, dbgGeneral, dbgCenter, dbgCenter, dbgRight, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgRight, dbgGeneral, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
        'F.CVIERA 31.12.2020
        grdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdProductos.Columns(0).Visible = True
        grdProductos.Columns(0).AllowSizing = False
        grdProductos.Columns(1).WrapText = True
        grdProductos.Columns(3).Visible = False
        grdProductos.Columns(4).Visible = False
        grdProductos.Columns(5).Visible = False
        grdProductos.Columns(6).Visible = False
        grdProductos.Columns(7).Visible = True
        grdProductos.Columns(8).Visible = True
        grdProductos.Columns(10).Visible = False
        grdProductos.Columns(12).Visible = False
        grdProductos.Columns(13).Visible = False
        grdProductos.Columns(15).Visible = False
        grdProductos.Columns(17).Visible = False
        grdProductos.Columns("DES_GRUPO").Visible = False
        grdProductos.Columns("DES_CLASE").Visible = False
        grdProductos.Columns("DESC_LABORATORIO").Visible = False
        grdProductos.Columns("PROM").Visible = True
        grdProductos.Columns("FLG_SEG").Visible = False 'CVIERA 31.12.2020
        grdProductos.Columns(4).AllowSizing = False
        grdProductos.Columns(0).FetchStyle = True
        grdProductos.Columns(1).FetchStyle = True
        grdProductos.Columns(2).FetchStyle = True
        grdProductos.Columns(3).FetchStyle = True
        grdProductos.Columns(6).FetchStyle = True
        grdProductos.Columns(7).FetchStyle = True
        grdProductos.Columns(8).FetchStyle = True
        grdProductos.Columns(12).FetchStyle = True
        grdProductos.Columns("PROM").FetchStyle = True
        grdProductos.Columns(3).NumberFormat = "##,##0.00"
        grdProductos.Columns(7).NumberFormat = "##0.00"
        grdProductos.Columns(8).NumberFormat = "##0.00"
        
        grdProductos.FetchRowStyle = True

        For Each columna In grdProductos.Columns
            columna.AllowSizing = False
        Next
        
        arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO_2", "DES_LINEA", "IMP_PRODUCTO", "ASO_SUSTITUTO", "COD_INDICADOR_RECETA", "STOCK", "IMP_COMI_UND", "IMP_COMI_FRA", "PRECIO", "EST_PRODUCTO", "STOCK", "FLG_EXCLUSIVO", "MSG_COMERCIAL", "STOCK_CAD", "PCT_COMI", "IMP_COMI_FRA", "COD_PROD_INK", "DES_GRUPO", "DES_CLASE", "DESC_LABORATORIO", "MSJ_ARG")
        arrCaption = Array("Codigo", "Producto", "Linea", "Kairos", "Sustituto", "CodIndicador", "Stock", "Comisi", "Com.F", "Publico", "Est", "Local", "flg_exclusivo", "Mensaje", "Cadena", "Comi", "Com. Gar.", "COD_PROD_INK", "Grupo", "Clase", "Laboratorio", "Argumento")
        'arrAncho = Array(550, 3360, 790, 600, 0, 0, 0, 0, 0, 700, 390, 550, 400, 2000, 550)
        arrAncho = Array(0, 5500, 2000, 0, 0, 0, 0, 0, 0, 700, 800, 750, 0, 0, 750, 600, 700, 0, 5000, 0, 0, 0)
        arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgRight, dbgGeneral, dbgGeneral, dbgGeneral, dbgRight, dbgRight, dbgRight, dbgCenter, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgGeneral, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
        
        grdAlternativos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdAlternativos.Columns(0).Visible = True
        grdAlternativos.Columns(0).AllowSizing = False
        grdAlternativos.Columns(1).WrapText = True
        grdAlternativos.Columns(3).Visible = False
        grdAlternativos.Columns(4).Visible = False
        grdAlternativos.Columns(5).Visible = False
        grdAlternativos.Columns(6).Visible = False
        grdAlternativos.Columns(7).Visible = False
        grdAlternativos.Columns(8).Visible = False
        grdAlternativos.Columns(12).Visible = False
        grdAlternativos.Columns(13).Visible = False
        grdAlternativos.Columns(15).Visible = True
        grdAlternativos.Columns(17).Visible = False
        grdAlternativos.Columns("DES_GRUPO").Visible = False
        grdAlternativos.Columns("DES_CLASE").Visible = False
        grdAlternativos.Columns("DESC_LABORATORIO").Visible = False
        grdAlternativos.Columns("MSJ_ARG").Visible = False
        grdAlternativos.Columns(4).AllowSizing = False
        grdAlternativos.Columns(0).FetchStyle = True
        grdAlternativos.Columns(1).FetchStyle = True
        grdAlternativos.Columns(2).FetchStyle = True
        grdAlternativos.Columns(3).FetchStyle = True
        grdAlternativos.Columns(6).FetchStyle = True
        grdAlternativos.Columns(7).FetchStyle = True
        grdAlternativos.Columns(8).FetchStyle = True
        grdAlternativos.Columns(12).FetchStyle = True
        grdAlternativos.Columns(16).FetchStyle = True
        grdAlternativos.Columns(3).NumberFormat = "##,##0.00"
        grdAlternativos.Columns(7).NumberFormat = "##0.00"
        grdAlternativos.Columns(8).NumberFormat = "##0.00"

        For Each columna In grdAlternativos.Columns
            columna.AllowSizing = False
        Next

        grdComplementarios.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdComplementarios.Columns(0).Visible = True
        grdComplementarios.Columns(0).AllowSizing = False
        grdComplementarios.Columns(1).WrapText = True
        grdComplementarios.Columns(3).Visible = False
        grdComplementarios.Columns(4).Visible = False
        grdComplementarios.Columns(5).Visible = False
        grdComplementarios.Columns(6).Visible = False
        grdComplementarios.Columns(7).Visible = False
        grdComplementarios.Columns(8).Visible = False
        grdComplementarios.Columns(12).Visible = False
        grdComplementarios.Columns(13).Visible = False
        grdComplementarios.Columns(17).Visible = False
        grdComplementarios.Columns("DES_GRUPO").Visible = False
        grdComplementarios.Columns("DES_CLASE").Visible = False
        grdComplementarios.Columns("DESC_LABORATORIO").Visible = False
        grdComplementarios.Columns("MSJ_ARG").Visible = False
        grdComplementarios.Columns(4).AllowSizing = False
        grdComplementarios.Columns(0).FetchStyle = False
        grdComplementarios.Columns(1).FetchStyle = False
        grdComplementarios.Columns(2).FetchStyle = True
        grdComplementarios.Columns(3).FetchStyle = False
        grdComplementarios.Columns(6).FetchStyle = False
        grdComplementarios.Columns(7).FetchStyle = False
        grdComplementarios.Columns(8).FetchStyle = False
        grdComplementarios.Columns(12).FetchStyle = False
        grdComplementarios.Columns(3).NumberFormat = "##,##0.00"
        grdComplementarios.Columns(7).NumberFormat = "##0.00"
        grdComplementarios.Columns(8).NumberFormat = "##0.00"

        For Each columna In grdComplementarios.Columns
            columna.AllowSizing = False
        Next
End If

Exit Sub

CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
End Sub

Sub Format_Grilla_RM()
           arrCampos = Array("COD_TIPO_INSUMO", "DES_TIPO_INSUMO", _
                             "COD_PRODUCTO", "DES_PRODUCTO", _
                             "DES_UND_CAPACIDAD_ABREV", "IMP_PRECIO_VTA", _
                             "DES_CLASE_COM", "DES_CATEGORIA_COM", _
                             "STOCK", "PCT_MARGEN", _
                             "IMP_COSTO_UNI", "COD_UND_CAPACIDAD")
                              
           arrCaption = Array("Codigo", "T Insumo", _
                              "Codigo", "Descripción", _
                              "Medida", "Precio Vta", _
                              "Clase", "Categoria", _
                              "Stock", "Pct Margen", _
                              "Pre Uni", "Cod Und")
                               
           arrAncho = Array(900, 1000, _
                            900, 2500, _
                            1000, 900, _
                            1500, 800, _
                            900, 900, _
                            900, 800)
                             
           arrAlineacion = Array(dbgCenter, dbgLeft, _
                                 dbgCenter, dbgLeft, _
                                 dbgLeft, dbgRight, _
                                 dbgLeft, dbgLeft, _
                                 dbgRight, dbgRight, _
                                 dbgRight, dbgCenter)
                                  
           grdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
           grdProductos.Columns(0).Visible = False
           grdProductos.Columns(0).AllowSizing = False
           grdProductos.Columns(1).WrapText = True
           grdProductos.Columns(2).Visible = False
           grdProductos.Columns(3).WrapText = True
           grdProductos.Columns(4).NumberFormat = "##,##0.00"
           grdProductos.Columns(6).Visible = False
           grdProductos.Columns(7).Visible = False
           grdProductos.Columns(8).Visible = False
           grdProductos.Columns(9).Visible = False
           grdProductos.Columns(10).Visible = False
           grdProductos.Columns(11).Visible = True
           grdProductos.RowDividerStyle = 1
End Sub

Sub mostrarAlternativos()

        'Dim CodUsu, nomPc, nomLocal, nomUsu, CodPerfil, CodLocal As String
        'CodUsu = objUsuario.Codigo
        'nomUsu = objUsuario.Nombre
        'nomPc = objUsuario.NombrePC
        'CodLocal = objUsuario.CodigoLocal
        'nomLocal = objUsuario.NombreLocal
        'CodPerfil = objUsuario.Perfil
                     
        fraAlternativos.Visible = lblnModo
        grdAlternativos.Visible = Not lblnModo

        fraComplementarios.Visible = lblnModo
        grdComplementarios.Visible = Not lblnModo

        lblnModo = Not lblnModo

        If Me.txtBuscar.Text <> "" Then
            If lblnModo Then
                objUsuario.GrabaLogBusqueda 1, Me.txtBuscar.Text
            Else
                objUsuario.GrabaLogBusqueda 0, Me.txtBuscar.Text
            End If
        End If
        
        grdProductos_RegistroSeleccionado ""
End Sub

Sub AgregaCupon(ByVal Texto As String)
Dim objDocDest As New clsDocumentoPago
Dim rsDocDest As oraDynaset
Dim rs As oraDynaset
Dim objFP As New clsFormaPago

'''<12-MAY-14    TCT Determina si Cupon Ingresado es Valido >

Set rs = objFP.ListaCupon(Texto)
Dim NumDocPago, CodDocPago As String
NumDocPago = "" & rs("NUM_DOCUMENTO_PAGO")
CodDocPago = "" & rs("COD_DOCUMENTO_PAGO")
Set rsDocDest = objDocDest.DetalleDocumento(NumDocPago, CodDocPago)
If rsDocDest.RecordCount = 0 Then
  MsgBox "El número de documento ingresado no existe o ya fue usado.", vbOKOnly + vbInformation, "Advertencia"
  Set objDocDest = Nothing
  'KeyAscii = 0
  Me.txtBuscar.SetFocus 'txtNroDoc.SetFocus
  Exit Sub
End If
                

Dim Fecha, Cliente, DNI As String
Dim i As Integer
Fecha = rs("FCH_EMI")
Fecha = left$(Fecha, 10)


Cliente = "" & rs("DES_CLIENTE")
DNI = "" & rs("NUM_DOCUMENTO_ID")

If Cliente = "" Then
        
    'frm_VTA_FormaPagoDD.txtNroDoc.Text = "" & rs("NUM_DOCUMENTO_PAGO")
    'frm_VTA_FormaPagoDD.mskFecEmi.Text = "" & Fecha
    'frm_VTA_FormaPagoDD.pstrDato = "" & rs("COD_FORMA_PAGO")
    'frm_VTA_FormaPagoDD.pstrDatoDes = "" & rs("DES_FORMA_PAGO")
    'frm_VTA_FormaPagoDD.od.FindFirst "COD_HIJO=" & "'" & rs("COD_HIJO") & "'"
    'frm_VTA_FormaPagoDD.CargarValorFP
    'frm_VTA_FormaPagoDD.Show
    'frm_VTA_FormaPagoDD.SetFocus
    
    If (frmPedido.pstrDniCli <> "") Then
    
        'frm_VTA_FormaPagoDD.txtNombre.Text = frmPedido.pstrNomcli
        'frm_VTA_FormaPagoDD.txtDNI.Text = frmPedido.pstrDniCli
                    
        objVenta.AgregaFormaPago "" & rs("COD_FORMA_PAGO"), _
                                 "" & rs("DES_FORMA_PAGO"), _
                                 "" & rs("COD_HIJO"), _
                                 "" & rs("DES_HIJO"), _
                                 "" & rs("IMP_VALOR"), _
                                 "", rs("COD_MONEDA"), _
                                 "", "", _
                                 "", "", _
                                 objUsuario.TipoCambio, "", _
                                 "", "", _
                                 "", rs("NUM_DOCUMENTO_PAGO"), _
                                 "", "", _
                                 "", "", _
                                 "" & Fecha, "", _
                                 "", frmPedido.pstrNomcli, _
                                 frmPedido.pstrDniCli, rs("COD_HIJO")
        
        MsgBox "Se ha agregado el cupon Nº " & Texto & Chr(13) & "" & rs("DES_TITULO") & " " & rs("DES_SUBTITULO"), vbOKOnly + vbInformation, App.ProductName
        Me.txtBuscar.Text = ""
        Me.txtBuscar.SetFocus
        
        
        'frm_VTA_FormaPagoDD.txtDNI.Text = frmPedido.pstrDniCli
        'frm_VTA_FormaPagoDD.TxtNombre.Text = frmPedido.pstrNomcli
        'frm_VTA_FormaPagoDD.Aceptar
        frmPedido.Cal_Montos
    Else
        MsgBox "Se ha agregado el cupon Nº " & Texto & Chr(13) & "" & rs("DES_TITULO") & " " & rs("DES_SUBTITULO") & "." & Chr(13) & " Para hacer efectivo el descuento debe Fidelizar al Cliente", vbOKOnly + vbInformation, App.ProductName
        frmPedido_Busca_Cli.Show vbModal
        
        If Not (frmPedido.pstrNomcli = "" And frmPedido.pstrDniCli = "") Then
              objVenta.AgregaFormaPago "" & rs("COD_FORMA_PAGO"), _
                                 "" & rs("DES_FORMA_PAGO"), _
                                 "" & rs("COD_HIJO"), _
                                 "" & rs("DES_HIJO"), _
                                 "" & rs("IMP_VALOR"), _
                                 "", rs("COD_MONEDA"), _
                                 "", "", _
                                 "", "", _
                                 objUsuario.TipoCambio, "", _
                                 "", "", _
                                 "", rs("NUM_DOCUMENTO_PAGO"), _
                                 "", "", _
                                 "", "", _
                                 "" & Fecha, "", _
                                 "", frmPedido.pstrNomcli, _
                                 frmPedido.pstrDniCli, rs("COD_HIJO")
        
        'MsgBox "Se ha agregado el cupon Nº " & Texto & Chr(13) & "" & rs("DES_TITULO") & " " & rs("DES_SUBTITULO"), vbOKOnly + vbInformation, App.ProductName
        Me.txtBuscar.Text = ""
        Me.txtBuscar.SetFocus
        
        
        'frm_VTA_FormaPagoDD.txtDNI.Text = frmPedido.pstrDniCli
        'frm_VTA_FormaPagoDD.TxtNombre.Text = frmPedido.pstrNomcli
        'frm_VTA_FormaPagoDD.Aceptar
        frmPedido.Cal_Montos

        'frm_VTA_FormaPagoDD.txtNombre.Focus
        End If
    End If
    
Else
    
    objVenta.AgregaFormaPago rs("COD_FORMA_PAGO"), _
                             rs("DES_FORMA_PAGO"), _
                             rs("COD_HIJO"), _
                             rs("DES_HIJO"), _
                             rs("IMP_VALOR"), _
                             "", rs("COD_MONEDA"), _
                             "", "", _
                             "", "", _
                             objUsuario.TipoCambio, "", _
                             "", "", _
                             "", rs("NUM_DOCUMENTO_PAGO"), _
                             "", "", _
                             "", "", _
                             "" & Fecha, "", _
                             "", "" & rs("DES_CLIENTE"), _
                             "" & rs("NUM_DOCUMENTO_ID"), rs("COD_HIJO")
    
    MsgBox "Se ha agregado el cupon Nº " & Texto & Chr(13) & "" & rs("DES_TITULO") & " " & rs("DES_SUBTITULO"), vbOKOnly + vbInformation, App.ProductName
    Me.txtBuscar.Text = ""
    Me.txtBuscar.SetFocus
    
    
    'frm_VTA_FormaPagoDD.Show
    'frm_VTA_FormaPagoDD.SetFocus
    'frm_VTA_FormaPagoDD.txtNroDoc.Text = "" & rs("NUM_DOCUMENTO_PAGO")
    'frm_VTA_FormaPagoDD.mskFecEmi.Text = "" & Fecha
    'frm_VTA_FormaPagoDD.pstrDato = rs("COD_FORMA_PAGO")
    'frm_VTA_FormaPagoDD.pstrDatoDes = rs("DES_FORMA_PAGO")
    'frm_VTA_FormaPagoDD.TxtNombre.Focus
    'frm_VTA_FormaPagoDD.od.FindFirst "COD_HIJO=" & "'" & rs("COD_HIJO") & "'"
    'frm_VTA_FormaPagoDD.txtDNI.Text = "" & rs("DES_CLIENTE")
    'frm_VTA_FormaPagoDD.TxtNombre.Text = "" & rs("NUM_DOCUMENTO_ID")
    'frm_VTA_FormaPagoDD.Aceptar
End If
frmPedido.Cal_Promo
End Sub

Sub AgregaConvenio(ByVal Texto As String)
Dim objDocDest As New clsDocumentoPago
Dim rsDocDest As oraDynaset
Dim rs As oraDynaset
Dim objBC As New clsConvenio
Dim strCodConvenioBarra, strCodClienteBarra, strDesClienteBarra, strEstadoCliente, strEstadoBenef As String

Set rs = objBC.ListaBarraConvenio(Texto)



If rs.RecordCount > 0 Then
    strCodConvenioBarra = "" & rs("cod_convenio")
    strCodClienteBarra = "" & rs("cod_cliente")
    strDesClienteBarra = "" & rs("des_cliente")
    strEstadoCliente = "" & rs("estadoCli")
    strEstadoBenef = "" & rs("estadoBen")
    If Not (strCodConvenioBarra = "" Or strCodClienteBarra = "") Then
       If strEstadoCliente = "1" Then
            If strEstadoBenef = "1" Then
                objVenta.LimpiaServicio
                'frm_VTA_Modalidad.LimpiarSiSalgodeGuia
                If objVenta.ptmModalidad = Guias_Remision Then
                    mdiPrincipal.subNuevo
                End If
                frmPedido.grdPedido.Rebind
                Unload Me
                ptmTipoPrecio = Convenio
                objVenta.CodigoTipoVenta = Venta_Convenio
                frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
                objVenta.ptmModalidad = Venta_Convenio
                frmPedido.Label4.Visible = True
                frmPedido.lblPctCopago.Visible = True
                frmPedido.Label8.Visible = True
                frmPedido.lblcopago.Visible = True
                frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
                frm_VTA_RecetarioM.pstrFlgRM = ""
                mdiPrincipal.cmdGrabaVenta.Enabled = True
                frm_VTA_Convenio.Show
                frm_VTA_Convenio.Carga_Convenio strCodConvenioBarra
                
                 With frm_VTA_ListaBeneficiario
                    .strCriterio = ""
                    .strCodConvenio = ""
                    .strCriterio = strDesClienteBarra
                    .strCodConvenio = strCodConvenioBarra
                    .output_Codigo_Beneficiario = ""
                    .output_Nombre_Beneficiario = ""
                    .Consulta
                    .grdBeneficiario_DblClick
                    frm_VTA_Convenio.txtBeneficiario.Text = ""
                    frm_VTA_Convenio.txtCodigoBeneficiario = .output_Codigo_Beneficiario
                    frm_VTA_Convenio.lblBeneficiario.Caption = .output_Nombre_Beneficiario
                    frm_VTA_Convenio.lblConsumo.Caption = Format(objVenta.Consumo, "###,###.00")
                    frm_VTA_Convenio.LblSaldo.Caption = Format(objVenta.LineaCred, "###,###.00")
                    Dim objConv As New clsConvenio
                    If objConv.EsRimac(strCodConvenioBarra) = True Then
                           On Error GoTo f
                        Dim t As Integer
                        While t <= UBound(frm_VTA_ListaBeneficiario.arrDatos)
                            frm_VTA_Convenio.pxdbDatos.Value(t, 5) = frm_VTA_ListaBeneficiario.arrDatos(t)
                            If t = 4 Then
                                frm_VTA_Convenio.txtBeneficiario.Text = frm_VTA_Convenio.pxdbDatos(t, 5)
                                frm_VTA_Convenio.lblBeneficiario.Caption = frm_VTA_Convenio.pxdbDatos(t, 5)
                            End If
                            
                         t = t + 1
                        Wend
                    frm_VTA_Convenio.ctlGrillaArray1.Rebind
f:
            
                   End If
                    'frm_VTA_Convenio.cboTipoCopago.SetFocus
                End With
            Else
                MsgBox "El Beneficiario se encuentra Inactivo", vbOKOnly + vbInformation, "Advertencia"
                Me.txtBuscar.SetFocus
            End If
        Else
            MsgBox "El Cliente se encuentra Inactivo", vbOKOnly + vbInformation, "Advertencia"
            Me.txtBuscar.SetFocus
        End If
    Else
        MsgBox "La Barra no Tiene Configurado el Convenio o el Beneficiario", vbOKOnly + vbInformation, "Advertencia"
        Me.txtBuscar.SetFocus
    End If
Else
    MsgBox "El Codigo de Barra no Existe", vbOKOnly + vbInformation, "Advertencia"
    Me.txtBuscar.SetFocus
End If

End Sub


Sub InsertarVentaPerdida()
Dim a As String

    If grdProductos.Columns(6).Value = "0" Then

        a = objVenta.GrabaVentaPerdida(objUsuario.CodigoLocal, _
                                       grdProductos.Columns("COD_PRODUCTO").Value, _
                                       objUsuario.Codigo)
         
        If a = "" Then
            MsgBox "Se ha Registrado el Producto para reposición", vbInformation + vbOKOnly, App.ProductName
        Else
            MsgBox a, vbCritical + vbOKOnly, "Atención"
        End If
    End If
End Sub

Sub AgregaFidelizado(ByVal Texto As String)
    Dim rs As oraDynaset
    Dim objCli As New clsClienteD

    On Error GoTo ErrControl
    
    Set rs = objCli.ListaClienteBarra(Texto)
    If rs.RecordCount > 0 Then
        gstrCodTarjetaFid = Texto
        If "" & rs("dni_cli") = "" Then
            frmPedido_Busca_Cli.Show vbModal
        Else
            frmPedido_Busca_Cli.ctlTxtDNI.Text = rs("dni_cli")
            frmPedido_Busca_Cli.ctlTxtDNI_KeyPress (13)
        End If
    Else
        MsgBox "No Existe la Tarjeta", vbCritical + vbOKOnly, "Atención"
        ''MsgBox "Desea Agregar la Tarjeta : " & Texto, vbOKOnly, "Aviso"
    End If
    Me.txtBuscar.Text = ""
    Me.txtBuscar.SetFocus
    
    Exit Sub
ErrControl:
    Me.txtBuscar.Text = ""
    Me.txtBuscar.SetFocus
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub AgregaMonedero(ByVal vNumeroTarjeta As String)
    Dim oFPC As New clsFPConstante
    Dim oFP As New clsFarmaPuntos
    Dim vAfiliado As clsAfiliado
    Dim rs As oraDynaset
    Dim vEstado As String, vDniCliente As String, vDniTemp As String, temp() As String
    Dim vUbigeo As String, vCodCliente As String, vTelefono As String, vCelular As String
    Dim vPuntosAcumulados As Double, vNombreAfiliado As String
    
    On Error GoTo Control
    Screen.MousePointer = vbHourglass
    
    gstrCodTarjetaMon = vNumeroTarjeta
    
    ' ir a orbis por el estado de la tarjeta
    temp = Split(oFP.GetEstadoTarjeta(vNumeroTarjeta, objUsuario.Codigo), "@")
    
    ' Si es una tarjeta no valida
    If temp(0) = oFPC.EstadoTarjeta.INVALIDA Then
        Err.Raise vbObjectError, "AgregaMonedero", "Tarjeta Invalida"
    End If
    
    ' si no lo encuentra en orbis buscar en el local
    If temp(0) = oFPC.EstadoTarjeta.SIN_ESTADO Or temp(0) = oFPC.EstadoTarjeta.INACTIVA Then
        temp = Split(objVenta.fnGetEstadoTarjetaMonedero(vNumeroTarjeta), "@")
        vEstado = temp(0)
        vDniCliente = temp(1)
        vPuntosAcumulados = CDbl(val(temp(2)))
        
        ' Si no encuentra los datos de la tarjeta en el local
        ' Afiliar por ser tarjeta nueva
        If vEstado = oFPC.EstadoTarjeta.SIN_ESTADO Then
                
            MsgBox "Programa Monedero: TARJETA NUEVA." & vbCrLf & vbCrLf, _
                   vbOKOnly + vbInformation, _
                   "Programa Monedero del Ahorro"
            
            frmPedido_Busca_Cli.Caption = "Programa Monedero del Ahorro"
            frmPedido_Busca_Cli.ctlTxtDNI.Text = vDniCliente
            frmPedido_Busca_Cli.b_afiliar = True
            frmPedido_Busca_Cli.b_monedero = True
            frmPedido_Busca_Cli.Show vbModal
            vCodCliente = frmPedido_Busca_Cli.v_delfrm
            vNombreAfiliado = frmPedido.lbl_Cliente.Caption
            vDniCliente = frmPedido.pstrDniCli
                
        ' Si lo encuentra en el local
        Else
        
            ' Solicitar DNI no escaneado y comparar con el obtenido
            vDniTemp = FrmPedido_Ingre_trj.ObtenerTarjeta("Validación Tarjeta Programa Monedero del Ahorro", _
                                                          eeDniCliente, _
                                                          False)
            
            ' Si es el mismo cargar datos del cliente
            If vDniTemp = vDniCliente Then
                Set rs = objVenta.fnListaAfiliadoMonedero(IIf(Len(vDniCliente) = 8, "002", "004"), vDniCliente)
                vCodCliente = "" & rs("COD_CLIENTE")
                vNombreAfiliado = "" & rs("DES_NOM_CLIENTE") & ", " & _
                                       rs("DES_APE_CLIENTE") & " " & _
                                       rs("DES_APE2_CLIENTE")
            
                MsgBox "Bienvenido " & vNombreAfiliado & vbCrLf & _
                       "DNI/CE: " & vDniCliente & vbCrLf & _
                       "Puntos: " & CStr(val(vPuntosAcumulados)), vbInformation + vbOKOnly, _
                       "Programa Monedero del Ahorro"
            ' Sino mostrar error y salir
            Else
                Err.Raise vbObjectError, "", "DNI/CE ingresado no corresponde a la tarjeta"
            End If
        End If
    ' si la encuenta en orbis
    Else
        vEstado = temp(0)
        vDniCliente = temp(1)
        vPuntosAcumulados = CDbl(val(temp(2)))
    
        ' Buscar datos del afiliado (DNI) en orbis
        Set vAfiliado = oFP.ObtenerDatosAfiliado(vDniCliente, objUsuario.Codigo)
        
        ' Si lo encuentra actualizar datos de BBDD
        If Not vAfiliado Is Nothing And vAfiliado.DNI <> "" Then
            
            vUbigeo = vAfiliado.Departamento & vAfiliado.Provincia & vAfiliado.Distrito
            vUbigeo = IIf(Len(vUbigeo) < 6, "", vUbigeo)
            vTelefono = IIf(IsNumeric(vAfiliado.Telefono), vAfiliado.Telefono, "")
            vCelular = IIf(IsNumeric(vAfiliado.Celular), vAfiliado.Celular, "")
            
            Set rs = objVenta.fnListaAfiliadoMonedero(vAfiliado.TipoDni, vAfiliado.DNI)
            vCodCliente = "" & rs("COD_CLIENTE")
            vCodCliente = objVenta.fnGrabaAfiliadoMonedero(vCodCliente, _
                                                           vAfiliado.TipoDni, _
                                                           vAfiliado.DNI, _
                                                           vAfiliado.Nombre, _
                                                           vAfiliado.ApParterno, _
                                                           vAfiliado.ApMarterno, _
                                                           vAfiliado.Email, _
                                                           vAfiliado.Genero, _
                                                           vAfiliado.FechaNacimiento, _
                                                           vUbigeo, _
                                                           vNumeroTarjeta, _
                                                           vTelefono, _
                                                           vCelular, _
                                                           vAfiliado.TipoLugar, _
                                                           vAfiliado.Direccion, _
                                                           vAfiliado.TipoDireccion, _
                                                           vAfiliado.Referencias, _
                                                           "S", _
                                                           "S")
            vNombreAfiliado = vAfiliado.Nombre & ", " & vAfiliado.ApParterno & " " & vAfiliado.ApMarterno
        
            MsgBox "Bienvenido " & vNombreAfiliado & vbCrLf & _
                   "DNI/CE: " & vDniCliente & vbCrLf & _
                   "Puntos: " & CStr(val(vPuntosAcumulados)), vbInformation + vbOKOnly, _
                   "Programa Monedero del Ahorro"
        End If
    End If
    
    If vCodCliente <> "" Then
        objVenta.EsVentaMonedero = True
        objVenta.NumeroTarjetaMonedero = vNumeroTarjeta
        objVenta.PuntosTarjetaMonedero = vPuntosAcumulados
        objVenta.CodigoCliente = vCodCliente
        
        frmPedido_Busca_Cli.v_delfrm = vCodCliente
        frmPedido.lbl_Cliente.Caption = vNombreAfiliado
        frmPedido.pstrDniCli = vDniCliente
        frmPedido.pstrNomcli = frmPedido.lbl_Cliente.Caption
        frmPedido.Cal_Promo
        frmPedido.Cal_Montos
        frmPedido.grdPedido.Refresh
        frmPedido.loadOptions
    End If
    
    Me.txtBuscar.Text = ""
    Me.txtBuscar.SetFocus
    
    Screen.MousePointer = vbDefault
    Exit Sub
Control:
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Sub AgregaCMR(ByVal Texto As String)
    Dim rs As oraDynaset
    Dim objCliCMR As New clsClienteD
    
    Set rs = objCliCMR.ListaClienteCMR(Texto)
    
    frmPedido.pstrDniCli = ""
    If rs.RecordCount > 0 Then
        frmPedido_Busca_Cli.ctlTxtDNI.Text = rs("NUM_DOCUMENTO_ID")
        frmPedido_Busca_Cli.ctlTxtDNI_KeyPress (13)
    Else
        frmPedido_Busca_Cli.ctlTxtDNI.Text = ""
        frmPedido_Busca_Cli.Show vbModal
        If frmPedido.pstrDniCli <> "" Then
            Dim strMensaje As String
            strMensaje = objCliCMR.GrabarTarjetaCMR(objVenta.CodigoCliente, Texto, objUsuario.Codigo)
            If strMensaje = "" Then
            Else
                MsgBox strMensaje, vbCritical, App.ProductName
            End If
        End If
    End If
    
    If (frmPedido.pstrDniCli <> "") Then
    
        frmPedido.optCredito.Value = True
        frmPedido.flgF6 = 0
        
        objVenta.AgregaFormaPago "002", _
                                 "TARJETA", _
                                 "032", _
                                 "TARJ-CMRFALABELLA", _
                                 "0", _
                                 "", "01", _
                                 "", "", _
                                 "", "", _
                                 objUsuario.TipoCambio, Texto, _
                                 "0", "", _
                                 "1", "", _
                                 "", "", _
                                 "", "", _
                                 "" & objUsuario.sysdate, "", _
                                 "", frmPedido.pstrNomcli, _
                                 frmPedido.pstrDniCli, ""
        
                                 
        MsgBox "Se ha agregado la Forma de Pago CMR", vbOKOnly + vbInformation, App.ProductName
        Me.txtBuscar.Text = ""
        Me.txtBuscar.SetFocus
        gintFidelizado = 1
        
        frmPedido.Cal_Montos
        frmPedido.Cal_Promo
   
    'Else
    '    MsgBox "Se ha agregado la Forma de Pago CMR", vbOKOnly + vbInformation, App.ProductName
    '    frmPedido_Busca_Cli.Show vbModal
    '
    '    If Not (frmPedido.pstrNomcli = "" And frmPedido.pstrDniCli = "") Then
    '        frmPedido.optCredito.Value = True
    '        'frmPedido.flgF6 = 1
    '
    '        objVenta.AgregaFormaPago "002", _
    '                             "TARJETA", _
    '                             "032", _
    '                             "TARJ-CMRFALABELLA", _
    '                             "0", _
    '                             "", "01", _
    '                             "", "", _
    '                             "", "", _
    '                             objUsuario.TipoCambio, Texto, _
    '                              "0", "", _
    '                             "1", "", _
    '                             "", "", _
    '                             "", "", _
    '                             "" & objUsuario.sysdate, "", _
    '                             "", frmPedido.pstrNomcli, _
    '                             frmPedido.pstrDniCli, ""
    '        Me.txtBuscar.Text = ""
    '        Me.txtBuscar.SetFocus
    '
    '        frmPedido.Cal_Montos
    '    End If
    End If

    'frmPedido.Cal_Promo

End Sub


