VERSION 5.00
Begin VB.Form frm_VTA_Servicios 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrillaArray grdServicios 
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5106
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "TV por Cable"
      Height          =   735
      Index           =   5
      Left            =   360
      Picture         =   "frm_VTA_Servicios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "Internet"
      Height          =   735
      Index           =   4
      Left            =   5160
      Picture         =   "frm_VTA_Servicios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "Tef. Celular"
      Height          =   735
      Index           =   1
      Left            =   1560
      Picture         =   "frm_VTA_Servicios.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "SOAT"
      Height          =   735
      Index           =   6
      Left            =   1560
      Picture         =   "frm_VTA_Servicios.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_Servicios.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_Servicios.frx":1692
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "Sv. Eléctricos"
      Height          =   735
      Index           =   3
      Left            =   3960
      Picture         =   "frm_VTA_Servicios.frx":1C1C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "Agua Potable"
      Height          =   735
      Index           =   2
      Left            =   2760
      Picture         =   "frm_VTA_Servicios.frx":205E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "Telefonía Fija"
      Height          =   735
      Index           =   0
      Left            =   360
      Picture         =   "frm_VTA_Servicios.frx":24A0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presione F1 para seleccionar el servicio a pagar"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   3390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presione F2 para seleccionar el servicio pagado"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   3390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F6"
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
      Index           =   9
      Left            =   840
      TabIndex        =   19
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F5"
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
      Index           =   8
      Left            =   5880
      TabIndex        =   17
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Index           =   7
      Left            =   2160
      TabIndex        =   15
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F7"
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
      Index           =   0
      Left            =   2175
      TabIndex        =   13
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
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
      Index           =   5
      Left            =   6097
      TabIndex        =   11
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
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
      Left            =   4380
      TabIndex        =   10
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   6255
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   6255
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios"
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
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   60
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frm_VTA_Servicios.frx":28E2
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F3"
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
      Left            =   3345
      TabIndex        =   3
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   4575
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Left            =   780
      TabIndex        =   1
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "frm_VTA_Servicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Dim objProducto As New clsProducto
Dim oraDato As oraDynaset
Dim dblTotal As Double

Private Sub ctlGrilla1_DblClick()
On Error GoTo handle
    Select Case ctlGrilla1.Columns(0).Value
       Case "001", "002", "005", "006", "007"  'Telefonica Fijo
            frm_VTA_ServicioTelefonica.strCodigoPadre = ctlGrilla1.Columns(0).Value
            frm_VTA_ServicioTelefonica.strDescripcionPadre = ctlGrilla1.Columns(1).Value
            frm_VTA_ServicioTelefonica.Show
        Case "003"  'Agua Potable
            frm_VTA_ServicioSedapal.strCodigoPadre = ctlGrilla1.Columns(0).Value
            frm_VTA_ServicioSedapal.strDescripcionPadre = ctlGrilla1.Columns(1).Value
            frm_VTA_ServicioSedapal.Show
        Case "004" 'Servicios Electricos
            frm_VTA_ServicioLuzDelSur.strCodigoPadre = ctlGrilla1.Columns(0).Value
            frm_VTA_ServicioLuzDelSur.strDescripcionPadre = ctlGrilla1.Columns(1).Value
            frm_VTA_ServicioLuzDelSur.Show
        Case "008"
            frm_VTA_RecargaVirtual.strCodigoPadre = ctlGrilla1.Columns(0).Value
           frm_VTA_RecargaVirtual.strDescripcionPadre = ctlGrilla1.Columns(1).Value
            frm_VTA_RecargaVirtual.flgRecarga = False
            frm_VTA_RecargaVirtual.Show
        Case "009"
            frm_VTA_RecargaVirtual.strCodigoPadre = ctlGrilla1.Columns(0).Value
            frm_VTA_RecargaVirtual.strDescripcionPadre = ctlGrilla1.Columns(1).Value
            frm_VTA_RecargaVirtual.flgRecarga = True
            frm_VTA_RecargaVirtual.Show
        Case "010"
            frm_VTA_ServicioSeguros.strCodigoPadre = ctlGrilla1.Columns(0).Value
            frm_VTA_ServicioSeguros.strDescripcionPadre = ctlGrilla1.Columns(1).Value
            frm_VTA_ServicioSeguros.Show
    End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub




Private Sub ctlGrilla1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handle


    If KeyCode = vbKeyReturn And Shift = 1 Then
        cmdAceptar_Click
        Exit Sub
    End If

    Select Case KeyCode
        Case vbKeyReturn
            ctlGrilla1_DblClick
    End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub


Private Sub Form_Load()
On Error GoTo handle
    setteaFormulario Me
    Dim arrCampos, arrCaption, arrWidth, arrAlineacion As Variant
    Dim columna As TrueDBGrid70.Column
    arrCampos = Array("COD_TIPO_SERVICIO", "DES_TIPO_SERVICIO")
    arrCaption = Array("Codigo", "Servicio")
    arrWidth = Array(800, 2500)
    arrAlineacion = Array(dbgLeft, dbgLeft)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrWidth, arrAlineacion
    Set ctlGrilla1.DataSource = objServicio.ListaTipo
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Código", "Descripción", "Tipo", "DescripcionTipo", "Operacion", "TipoVenta", "Monto", "TipoCambio", "Importe", "producto")
    arrWidth = Array(600, 2000, 0, 0, 1000, 0, 800, 800, 800, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgRight, dbgCenter, dbgCenter, dbgRight, dbgRight, dbgRight, dbgGeneral)
    
    grdServicios.FormatoGrilla arrCampos, arrCaption, arrWidth, arrAlineacion
    grdServicios.Columns(2).Visible = False
    grdServicios.Columns(3).Visible = False
    grdServicios.Columns(5).Visible = False
    grdServicios.Columns(9).Visible = False
    'grdServicios.ColumnFooter = True
    'grdServicios.Columns(7).FooterText = "Total S/."
    
    
    For Each columna In grdServicios.Columns
            columna.AllowSizing = False
    Next
    

    grdServicios.Array1 = objVenta.Servicios
    grdServicios.Rebind

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdAceptar_Click()
If Not objVenta.CodModalidadVenta = Servicio Then
    MsgBox "La modalidad no es Servicio, por favor seleccione la Modalidad adecuada", vbCritical, App.ProductName
    Exit Sub
End If
On Error GoTo handle
Dim i As Integer
Dim Y As Integer
Dim X As Integer
Dim strCodProducto As String
Dim strDesProducto As String
Dim intCant As Integer
Dim indicador As String
Dim PctComi As Double
If objVenta.Servicios.UpperBound(1) < 0 Then MsgBox "Debe selecionar algun servicio", vbCritical, App.ProductName: Exit Sub
    objVenta.Servicios.QuickSort 0, objVenta.Servicios.UpperBound(1), 9, XORDER_ASCEND, XTYPE_STRING
    Y = 0
    Do While True
        strCodProducto = objVenta.Servicios(Y, 9)
        strDesProducto = objVenta.Servicios(Y, 3)
        intCant = 0
        For i = Y To objVenta.Servicios.UpperBound(1)
            If strCodProducto <> objVenta.Servicios(i, 9) Then Exit For
            'intCant = intCant + 1
            intCant = intCant + Val(objVenta.Servicios(i, 13))
        Next i
        
        indicador = objProducto.CodIndicadorReceta(strCodProducto)
        PctComi = objProducto.pctComision(strCodProducto, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
        If objUsuario.CodLocalCallCenter = "1DLV" Then 'ECASTILLO 22.06.2020
            Set oraDato = objProducto.ListaDato("94", objUsuario.CodigoLocal, Format(Regular, "000"), strCodProducto, intCant, "0", "", objUsuario.CodLocalCallCenter)
        Else
            Set oraDato = objProducto.ListaDato(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Format(Regular, "000"), strCodProducto, intCant, "0", "", objUsuario.CodLocalCallCenter)
        End If
        
        objVenta.AgregaProducto strCodProducto, strDesProducto, intCant, "0", IIf(IsNull(oraDato(4).Value), "0", oraDato(4).Value), objVenta.CodigoTipoVenta, Producto_Normal, , , , , , indicador, PctComi
        
        If i > objVenta.Servicios.UpperBound(1) Then Exit Do
        Y = i
    Loop


'    For i = 0 To objVenta.Servicios.UpperBound(1)
'            Set oraDato = objProducto.ListaDato(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Format(Regular, "000"), objVenta.Servicios(i, 9), dblCant, "0", "")
'            objVenta.AgregaProducto objVenta.Servicios(i, 9), objVenta.Servicios(i, 3), 1, "0", IIf(IsNull(oraDato(4).Value), "0", oraDato(4).Value), objVenta.CodigoTipoVenta, Producto_Normal
'        End If
'    Next i
'''
    
    
    
    
    frmPedido.Cal_Montos
    '
    frmPedido.Cal_Promo
    frmPedido.grdPedido.Rebind
    'Unload Me

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click
    On Error GoTo Control
        psub_KeyDownAplicacion KeyCode, Shift
        Select Case KeyCode
            
            Case vbKeyF1
                ctlGrilla1.SetFocus
                
'                cmdTelefonica_Click
            Case vbKeyF2
                grdServicios.SetFocus
'                cmdSedapal_Click
            Case vbKeyF3
'                cmdLuzDelSur_Click
            Case vbKeyF4
'                cmdAlquiler_Click
            Case vbKeyF5
'                CmdKeito_Click

            Case vbKeyEscape
                Unload Me
        End Select
    
    Exit Sub
Control:
        MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
    'objVenta.CancelarVenta
End Sub

Private Sub grdServicios_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handle
If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click

Select Case KeyCode
    Case vbKeyDelete
        grdServicios.Delete
           
End Select
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub


Private Sub CalFooter()
    Dim k%
    Dim intCant As Integer
    
    intCant = 0: dblTotal = 0
    
    For k = 0 To objVenta.Servicios.UpperBound(1)
    
        If objVenta.CobroResponsabilidad(k, 8) > 0 Or objVenta.CobroResponsabilidad(k, 8) <> "" Then
            dblTotal = dblTotal + Val(objVenta.CobroResponsabilidad(k, 8))
        End If
    Next k
    grdServicios.Columns(8).FooterText = Format(dblTotal, "#,###,##0.00")
    
    
End Sub

