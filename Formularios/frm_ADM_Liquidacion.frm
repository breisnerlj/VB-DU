VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_ADM_Liquidacion 
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   7125
   Begin VB.TextBox txtLiquidacion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4920
      TabIndex        =   17
      Top             =   105
      Width           =   2055
   End
   Begin vbp_Ventas.ctlToolBar cToolBar 
      Height          =   600
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1058
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tip. Doc."
      TabPicture(0)   =   "frm_ADM_Liquidacion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdTipDoc"
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "F. Pag."
      TabPicture(1)   =   "frm_ADM_Liquidacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdForPag"
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Mod. Vta."
      TabPicture(2)   =   "frm_ADM_Liquidacion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdModVta"
      Tab(2).Control(1)=   "Label1(8)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "T. Cred."
      TabPicture(3)   =   "frm_ADM_Liquidacion.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdTarCre"
      Tab(3).Control(1)=   "Label1(2)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Scotiabank"
      TabPicture(4)   =   "frm_ADM_Liquidacion.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(6)"
      Tab(4).Control(1)=   "grdScotia"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Remesas"
      TabPicture(5)   =   "frm_ADM_Liquidacion.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdRemesa"
      Tab(5).Control(1)=   "Label1(5)"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Diferencias"
      TabPicture(6)   =   "frm_ADM_Liquidacion.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "grdDife"
      Tab(6).Control(1)=   "Label1(4)"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Errores"
      TabPicture(7)   =   "frm_ADM_Liquidacion.frx":00C4
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Label1(3)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "grdError"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      Begin vbp_Ventas.ctlGrilla grdModVta 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   2
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin vbp_Ventas.ctlGrilla grdForPag 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   18
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin vbp_Ventas.ctlGrilla grdTipDoc 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   19
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin vbp_Ventas.ctlGrilla grdTarCre 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   22
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin vbp_Ventas.ctlGrilla grdError 
         Height          =   3975
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7011
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin vbp_Ventas.ctlGrilla grdDife 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   25
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin vbp_Ventas.ctlGrilla grdRemesa 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   27
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin vbp_Ventas.ctlGrilla grdScotia 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   29
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6376
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
         MultiSelect     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3.1.- Notas de credito en positivo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "5.- Transacciones de Scotiabank"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   6
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "6.- Remesas"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   -74760
         TabIndex        =   28
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "7.- Notas de crédito en positivo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   26
         Top             =   840
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "4.- Documentos con Tarjetas de Crédito"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   23
         Top             =   840
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2.- Formas de Pago"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   21
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1.- Detalle de Venta por Tipo de Documentos"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   20
         Top             =   840
         Width           =   3210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3.- Ventas x Modalidad de Venta"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   8
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   2310
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   7095
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   16
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtPerfil 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtMaquina 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Index           =   5
         Left            =   2880
         TabIndex        =   9
         Top             =   1110
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Perfil:"
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
         Index           =   4
         Left            =   2880
         TabIndex        =   8
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
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
         Index           =   3
         Left            =   2880
         TabIndex        =   7
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Máquina:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1110
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Local:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   795
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Liquidación:"
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
      Index           =   6
      Left            =   3480
      TabIndex        =   10
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "frm_ADM_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pCodLiq As String
Private objLiq As clsLiquidacion

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cToolBar_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    On Error GoTo Control
    
    Select Case boton
        Case tlbTipoBoton.tb_Actualizar
            Call LlenarGrids
        Case tlbTipoBoton.Imprimir
            Call ImprimirGrid
        Case tlbTipoBoton.tb_Excel
            Call ExportarGrid
        Case tlbTipoBoton.tb_email
            Call MailingGrid
        Case tlbTipoBoton.salir
            Unload Me
    End Select
    
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Load()
    On Error GoTo Control
    
    Set objLiq = New clsLiquidacion
    
    setteaFormulario Me
    
    SSTab1.Tab = 0

    txtLiquidacion.Text = pCodLiq
    
    Call FormatToolBar
    Call FormatGrids
    Call LlenarCabecera
    Call LlenarGrids
    
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Control
    
    pCodLiq = vbNullString
    Set grdTipDoc.DataSource = Nothing
    Set grdForPag.DataSource = Nothing
    Set grdModVta.DataSource = Nothing
    Set grdTarCre.DataSource = Nothing
    Set grdScotia.DataSource = Nothing
    Set grdRemesa.DataSource = Nothing
    Set grdDife.DataSource = Nothing
    Set grdError.DataSource = Nothing
    Set objLiq = Nothing

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub FormatToolBar()
    Dim i As Byte
    
    On Error GoTo Control
    
    cToolBar.Width = 0
    For i = 1 To 15
        Select Case i
            Case 6, 7, 8, 15
                cToolBar.Buttons(i).Visible = True
                cToolBar.Width = cToolBar.Width + cToolBar.Buttons(i).Width
            Case Else
                cToolBar.Buttons(i).Visible = False
        End Select
    Next
    
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub LlenarCabecera()
    Dim rs As oraDynaset

    On Error GoTo Control
    
    Set rs = objLiq.RepoListaDatos(pCodLiq)
    txtEmpresa.Text = rs.Fields("CIA")
    txtLocal.Text = rs.Fields("COD_LOCAL")
    txtMaquina.Text = rs.Fields("COD_MAQUINA")
    txtUsuario.Text = rs.Fields("COD_USUARIO") & " - " & _
                      rs.Fields("APE_PAT_USUARIO") & " " & _
                      rs.Fields("APE_MAT_USUARIO") & " " & _
                      rs.Fields("DES_NOMBRE")
    txtPerfil.Text = rs.Fields("DES_PERFIL")
    txtEstado.Text = rs.Fields("FLG_ESTADO_CAJA")
    Set rs = Nothing

    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub LlenarGrids()
    On Error GoTo Control
    
    Set grdTipDoc.DataSource = objLiq.RepoListaXTipDoc(pCodLiq)
    Set grdForPag.DataSource = objLiq.RepoListaXForPag(pCodLiq)
    Set grdModVta.DataSource = objLiq.RepoListaXModVta(pCodLiq)
    Set grdTarCre.DataSource = objLiq.RepoListaDocTarjeta(pCodLiq)
    Set grdScotia.DataSource = objLiq.RepoListaDocCajeExpress(pCodLiq)
    Set grdRemesa.DataSource = objLiq.RepoListaRemesas(pCodLiq)
    Set grdDife.DataSource = objLiq.RepoListaNCMalgen(pCodLiq)
    Set grdError.DataSource = objLiq.RepoListaErrores(pCodLiq)

    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ExportarGrid()
    On Error GoTo Control
    
    Select Case SSTab1.Tab
        Case 0
            grdTipDoc.MostrarExcel
        Case 1
            grdForPag.MostrarExcel
        Case 2
            grdModVta.MostrarExcel
        Case 3
            grdTarCre.MostrarExcel
        Case 4
            grdScotia.MostrarExcel
        Case 5
            grdRemesa.MostrarExcel
        Case 6
            grdDife.MostrarExcel
        Case 7
            grdError.MostrarExcel
    End Select

    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ImprimirGrid()
    On Error GoTo Control
    
    Select Case SSTab1.Tab
        Case 0
            grdTipDoc.MostrarImprimir
        Case 1
            grdForPag.MostrarImprimir
        Case 2
            grdModVta.MostrarImprimir
        Case 3
            grdTarCre.MostrarImprimir
        Case 4
            grdScotia.MostrarImprimir
        Case 5
            grdRemesa.MostrarImprimir
        Case 6
            grdDife.MostrarImprimir
        Case 7
            grdError.MostrarImprimir
    End Select

    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub MailingGrid()
    On Error GoTo Control
    
    Select Case SSTab1.Tab
        Case 0
            grdTipDoc.MostrarEmail
        Case 1
            grdForPag.MostrarEmail
        Case 2
            grdModVta.MostrarEmail
        Case 3
            grdTarCre.MostrarEmail
        Case 4
            grdScotia.MostrarEmail
        Case 5
            grdRemesa.MostrarEmail
        Case 6
            grdDife.MostrarEmail
        Case 7
            grdError.MostrarEmail
    End Select

    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub FormatGrids()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim i As Integer
  
    On Error GoTo Control
    
    'Por tipo de documento
    arrCampos = Array("FCH_EMISION", "COD_TIPO_DOCUMENTO", "COD_ESTADO", "CTD", "MIN_NUM_DOCUMENTO", _
                      "MAX_NUM_DOCUMENTO", "MTO_TOTAL")
    arrCaption = Array("Fecha", "Tip_Doc", "Est_Doc", "#Docs", "Minimo", _
                       "Maximo", "Importe_Soles")
    arrAncho = Array(1100, 900, 900, 900, 1300, _
                     1300, 1300)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgRight)
    
    grdTipDoc.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdTipDoc.Columns.Count - 1
        grdTipDoc.Columns(i).AllowSizing = False
        grdTipDoc.Columns(i).WrapText = True
    Next
    
    'Por forma de pago
    arrCampos = Array("DES_HIJO", "RETIRO", "MTO_TOTAL")
    arrCaption = Array("FormaCobroRegistradas", "Retiro", "Importe_Soles(Aprox)")
    arrAncho = Array(1700, 1200, 2000)
    arrAlineacion = Array(dbgLeft, dbgCenter, dbgRight)
    
    grdForPag.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdForPag.Columns.Count - 1
        grdForPag.Columns(i).AllowSizing = False
        grdForPag.Columns(i).WrapText = True
    Next
    
    'Por modalidad de venta
    arrCampos = Array("COD_MODALIDAD_VENTA", "DES_MODALIDAD", "COD_TIPO_DOCUMENTO", "MTO_TOTAL")
    arrCaption = Array("C. Modalidad", "Modalidad", "T. Documento", "Importe")
    arrAncho = Array(1100, 1600, 1200, 1200)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgRight)
    
    grdModVta.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdModVta.Columns.Count - 1
        grdModVta.Columns(i).AllowSizing = False
        grdModVta.Columns(i).WrapText = True
    Next
    
    'Por tarjetas de credito
    arrCampos = Array("COD_TIPO_DOCUMENTO", "NUM_DOCUMENTO", "DES_HIJO", "RET_EFEC", "IMP_MONEDA_NAC")
    arrCaption = Array("Doc_con_Tarjeta", "Num_Documento", "Forma_Pago", "Retiro", "Importe_Soles(Solo de la Parte de Tjt)")
    arrAncho = Array(1200, 1200, 1700, 1000, 1500)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgRight)
    
    grdTarCre.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdTarCre.Columns.Count - 1
        grdTarCre.Columns(i).AllowSizing = False
        grdTarCre.Columns(i).WrapText = True
    Next
    
    'Scotiabank
    arrCampos = Array("CIA", "COD_LOCAL", "COD_TIPO_DOCUMENTO", "NUM_DOCUMENTO", "MTO_TOTAL", _
                      "IMP_MONEDA_NAC", "XNUM_TARJETA", "NUM_AUTORIZACION")
    arrCaption = Array("CIA", "Local", "TD", "Nº Documento", "Importe", _
                       "Importe_Soles", "Nº Tarjeta", "Nº Autoriza.")
    arrAncho = Array(1000, 1000, 1000, 1000, 1200, _
                     1200, 1200, 1200)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgRight, dbgRight)
    
    grdScotia.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdScotia.Columns.Count - 1
        grdScotia.Columns(i).AllowSizing = False
        grdScotia.Columns(i).WrapText = True
    Next
    
    'Remesas
    arrCampos = Array("COD_REMESA", "COD_MAQUINA", "IMP_TOTAL", "FCH_DEPOSITO", "EST_REMESA")
    arrCaption = Array("C. Remesa", "Maquina", "Imp Total", "Fecha", "Importe")
    arrAncho = Array(1000, 1000, 1000, 1000, 1200)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    
    grdRemesa.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdRemesa.Columns.Count - 1
        grdRemesa.Columns(i).AllowSizing = False
        grdRemesa.Columns(i).WrapText = True
    Next
    
    'Diferencias
    arrCampos = Array("CIA", "COD_TIPO_DOCUMENTO", "NUM_DOCUMENTO", "SEC_FORPAG_DOC", "IMP_SIN_REDONDEO")
    arrCaption = Array("CIA", "TD", "N. Documento", "", "Importe")
    arrAncho = Array(1000, 800, 800, 800, 1200)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    
    grdDife.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdDife.Columns.Count - 1
        grdDife.Columns(i).AllowSizing = False
        grdDife.Columns(i).WrapText = True
    Next
    
    'Errores
    arrCampos = Array("SEC_ERROR", "ITEM_ERROR", "DES_ERROR", "FCH_REGISTRA", "MACHINE", _
                      "COD_USUARIO", "TIP_ERROR", "ERROR_ORACLE", "ANIO", "MES", _
                      "DIA", "VERSION", "SEC_ARCHIVO", "COD_LIQUIDACION")
    arrCaption = Array("Sec.", "Item", "Descripción", "Fecha", "Maquina", _
                       "Usuario", "T.Error", "Err.Oracle", "Año", "Mes", _
                       "Día", "Versión", "Sec.Arch.", "Liquidación")
    arrAncho = Array(800, 800, 5000, 1000, 1200, _
                     1000, 1000, 1000, 800, 800, _
                     800, 1000, 1000, 1200)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    
    grdError.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For i = 0 To grdError.Columns.Count - 1
        grdError.Columns(i).WrapText = True
    Next
    
    grdError.Columns(5).Visible = False
    grdError.Columns(12).Visible = False
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
