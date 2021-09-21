VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_VTA_ConsultaDoc 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRecuperar 
      Caption         =   "Recu&perar"
      Height          =   615
      Left            =   4770
      Picture         =   "frm_VTA_ConsultaDoc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6360
      Width           =   950
   End
   Begin vbp_Ventas.ctlDataCombo dbcTipoDoc 
      Height          =   315
      Left            =   1920
      TabIndex        =   27
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Detalle"
      Height          =   615
      Left            =   3870
      Picture         =   "frm_VTA_ConsultaDoc.frx":0414
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6360
      Width           =   800
   End
   Begin VB.CommandButton cmdFormaPago 
      Caption         =   "&F. Pago"
      Height          =   615
      Left            =   2970
      Picture         =   "frm_VTA_ConsultaDoc.frx":099E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Width           =   800
   End
   Begin VB.OptionButton optTipoBusqueda 
      Caption         =   "&Fecha"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton optTipoBusqueda 
      Caption         =   "&Documento"
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
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdReimprimir 
      Caption         =   "Re &Imprimir"
      Height          =   615
      Left            =   1970
      Picture         =   "frm_VTA_ConsultaDoc.frx":0F28
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   900
   End
   Begin VB.CommandButton cmdAnulacion 
      Caption         =   "An&ular"
      Height          =   615
      Left            =   1070
      Picture         =   "frm_VTA_ConsultaDoc.frx":14B2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   800
   End
   Begin VB.CommandButton cmdAsignarUsuario 
      Caption         =   "&Asignar"
      Height          =   615
      Left            =   240
      Picture         =   "frm_VTA_ConsultaDoc.frx":1A3C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1860
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdReactivar 
      Caption         =   "&Reactivar"
      Height          =   615
      Left            =   120
      Picture         =   "frm_VTA_ConsultaDoc.frx":1FC6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   850
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   5820
      Picture         =   "frm_VTA_ConsultaDoc.frx":2550
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      Width           =   800
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   555
      Left            =   4920
      Picture         =   "frm_VTA_ConsultaDoc.frx":2ADA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaFin 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115933185
      CurrentDate     =   39001
   End
   Begin MSComCtl2.DTPicker dtpFechaIni 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115933185
      CurrentDate     =   39001
   End
   Begin vbp_Ventas.ctlTextBox txtNumDocFin 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Tipo            =   7
      MaxLength       =   11
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
   Begin vbp_Ventas.ctlTextBox txtNumDocIni 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Tipo            =   7
      MaxLength       =   11
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
   Begin vbp_Ventas.ctlGrilla grdDocumento 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4471
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   6
      Left            =   4080
      TabIndex        =   26
      Top             =   6960
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   3260
      TabIndex        =   25
      Top             =   6960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   5
      Left            =   2320
      TabIndex        =   24
      Top             =   6960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   4
      Left            =   1350
      TabIndex        =   23
      Top             =   6960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   3
      Left            =   600
      TabIndex        =   22
      Top             =   2460
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   21
      Top             =   6960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   11
      Left            =   6070
      TabIndex        =   20
      Top             =   6960
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
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
      Left            =   4605
      TabIndex        =   19
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
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
      Left            =   2385
      TabIndex        =   18
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Documento Final"
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
      Left            =   4380
      TabIndex        =   17
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Documento Inicial"
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
      Left            =   2175
      TabIndex        =   16
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Documento :"
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
      Left            =   120
      TabIndex        =   15
      Top             =   180
      Width           =   1800
   End
End
Attribute VB_Name = "frm_VTA_ConsultaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento
Dim objImpresion As New clsImpresiones
Public pblnFpago As Boolean


'Private Sub cmdAceptar_KeyDown(KeyCode As Integer, Shift As Integer)
'    MsgBox KeyCode
'End Sub

Private Sub cmdAnulacion_Click()
Dim Bookmark As Variant
Dim strTipoVenta As String
    
    On Error GoTo CtrlErr
    If grdDocumento.ApproxCount = 0 Then
        Exit Sub
    End If

    strTipoVenta = grdDocumento.Columns(8).Text
    If MsgBox("Desea anular la " + dbcTipoDoc.Text + " Nº " + grdDocumento.Columns(0).Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        grdDocumento.SetFocus
        Exit Sub
    End If
    
    If strTipoVenta = "001" Then
        If MsgBox("El documento que desea anular ha sido generado por un pedido de Delivery." & Chr(13) & "Esta seguro de anularlo?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirme") = vbNo Then
            Exit Sub
        End If
    End If


    Bookmark = grdDocumento.Bookmark


    Dim gvarError As String
    Dim ValorRet As String
    gvarError = objDocumento.Anula(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, dbcTipoDoc.BoundText, Replace(grdDocumento.Columns(0).Value, "-", ""), grdDocumento.Columns(4).Value, objUsuario.Codigo, ValorRet)
    
    If gvarError = "" Then
        cmdBuscar_Click
        MsgBox "Se anulo los siguientes documentos:" + Chr(13) + ValorRet, vbInformation, App.ProductName
        grdDocumento.Bookmark = Bookmark
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    grdDocumento.SetFocus
End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdAsignarUsuario_Click()

On Error GoTo CtrlErr

    If grdDocumento.ApproxCount = 0 Then
        Exit Sub
    End If
'    fraAsignar.Visible = True
'    lblCodigo.Visible = True
'    optDoc(0).Value = True
        
    frm_VTA_AsignarUsuarioDocumento.datos "Asignar Usuarios", grdDocumento.DataSource, Replace(grdDocumento.Columns(0).Value, "-", ""), left(grdDocumento.Columns(5).Value, 5), dbcTipoDoc.BoundText
        
Exit Sub

CtrlErr:
        MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdBuscar_Click()

    
   On Error GoTo Control
   If dbcTipoDoc.BoundText = "" Or dbcTipoDoc.BoundText = "*" Then
    MsgBox "Debe seleccionar el tipo de documento.", vbOKOnly + vbExclamation, "Error"
    dbcTipoDoc.SetFocus
    Exit Sub
   End If

    Set grdDocumento.DataSource = objDocumento.lista(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, dbcTipoDoc.BoundText, txtNumDocIni.Text, txtNumDocFin.Text, dtpFechaIni.Value, dtpFechaFin.Value + 1)
    
    grdDocumento.SetFocus
    

   Exit Sub

Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

''Private Sub cmdCambioCorrelativo_Click()
''    frm_VTA_Correccion.Show
''End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdDetalle_Click()
    grdDocumento_DblClick
End Sub

Private Sub cmdFormaPago_Click()
On Error GoTo Control

If grdDocumento.ApproxCount = 0 Then Exit Sub
frm_VTA_FormaPago.Param_Tipo_Documento = dbcTipoDoc.BoundText
frm_VTA_FormaPago.Param_Numero_Documento = Replace(grdDocumento.Columns(0).Value, "-", "")
frm_VTA_FormaPago.Modificacion = True
pblnFpago = True
frm_VTA_FormaPago.Show

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdReactivar_Click()

Dim Bookmark As Variant
    
    On Error GoTo CtrlErr
        If grdDocumento.ApproxCount = 0 Then
                    Exit Sub
        End If

    If MsgBox("Desea Reactivar la " + dbcTipoDoc.Text + " Nº " + grdDocumento.Columns(0).Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        grdDocumento.SetFocus
        Exit Sub
    End If


    Bookmark = grdDocumento.Bookmark
Dim gvarError As String
Dim ValorRet As String
    gvarError = objDocumento.Reactiva(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, dbcTipoDoc.BoundText, Replace(grdDocumento.Columns(0).Value, "-", ""), grdDocumento.Columns(4).Value, objUsuario.Codigo, ValorRet)
    
    If gvarError = "" Then
        cmdBuscar_Click
        MsgBox "Se reactivo los siguientes documentos" + Chr(13) + ValorRet, vbInformation, App.ProductName
        grdDocumento.Bookmark = Bookmark
        
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    End If
    grdDocumento.SetFocus
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdRecuperar_Click()
    Dim oFPC As New clsFPConstante
    Dim oFP As New clsFarmaPuntos
    Dim vEstadoTarjeta As clsTarjetaBean
    Dim vNroTarjeta As String, res As String, vNroDocumento As String
    Dim oDP As New clsDocumentoPago
    
    On Error GoTo CtrlErr
    
    If grdDocumento.ApproxCount = 0 Then
        Exit Sub
    End If
    
    If Not (dbcTipoDoc.BoundText = COD_TIPO_GRL) Then
        
        vNroDocumento = Replace(grdDocumento.Columns(0).Value, "-", "")
        
        'Solicitar y Validar tarjeta
        vNroTarjeta = FrmPedido_Ingre_trj.ObtenerTarjeta("Recuperación de Puntos")
        
        Screen.MousePointer = vbHourglass
        
        If oDP.buscaFun(vNroTarjeta) = "MONEDERO" Then
            
            Set vEstadoTarjeta = oFP.ValidarTarjetaAsociada(vNroTarjeta, objUsuario.Codigo)
            
            If vEstadoTarjeta.EstadoTarjeta <> oFPC.EstadoTarjeta.ACTIVA Then
                GoTo CtrlErrTrj
            End If
        
        Else
            GoTo CtrlErrTrj
        End If
        
        'Validar si se puede recuperar puntos del comprobante
        res = objVenta.RecuperarPuntosMonedero(objUsuario.CodigoEmpresa, _
                                               dbcTipoDoc.BoundText, _
                                               vNroDocumento, _
                                               objUsuario.CodigoLocal, _
                                               objUsuario.Codigo, _
                                               vEstadoTarjeta.NumeroTarjeta, _
                                               vEstadoTarjeta.PuntosTotalAcumulados, _
                                               "N")
        
        Screen.MousePointer = vbDefault
        
        MsgBox "La recuperación se verá reflejada en el transcurso del día." & vbCrLf & _
               "Porque no hay conexión con el servicio de puntos.", _
               vbInformation + vbOKOnly, App.ProductName
        
        FnImprimirVoucher dbcTipoDoc.BoundText, vNroDocumento, CDbl(Val(res)), vEstadoTarjeta
        
    End If
    Exit Sub
CtrlErrTrj:
    Screen.MousePointer = vbDefault
    MsgBox "La Tarjeta NO es válida", vbCritical + vbOKOnly, App.ProductName
    Exit Sub
CtrlErr:
    'MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
    Screen.MousePointer = vbDefault
    Debug.Print Err.Description
    MsgBox "No se pudo recuperar los puntos del comprobante." & vbCrLf & _
           Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub FnImprimirVoucher(ByVal vTipoDocumento As String, _
                              ByVal vNroDocumento As String, _
                              ByVal vCantidadPuntos As Double, _
                              ByVal vTarjeta As clsTarjetaBean)
    Dim strMensaje As String
    Dim intConstante As Integer
    
    intConstante = 184
    
    Screen.MousePointer = vbHourglass
    
    If SetearImpresoraCupon = False Then
        Screen.MousePointer = vbDefault
        MsgBox "No tiene la impresora de cupones Instalada o tiene diferente nombre", vbCritical, App.ProductName
        Exit Sub
    End If
    
    'Logo
    If objVenta.EsMFA(objUsuario.CodigoLocal) = True Then
        mdiPrincipal.imgLogoBtl.Picture = mdiPrincipal.ImageList1.ListImages(3).Picture
        Printer.PaintPicture mdiPrincipal.imgLogoBtl, 570, 20, 2950, 850
    Else
        mdiPrincipal.imgLogoBtl.Picture = mdiPrincipal.ImageList1.ListImages(1).Picture
        Printer.PaintPicture mdiPrincipal.imgLogoBtl, 1000, 20, 2150, 950
    End If
    Printer.CurrentY = 1154
    
    'Titulo
    Printer.FontName = "Arial"
    Printer.FontSize = 9
    Printer.Font.Bold = False
    centra_printer "Constancia de Recuperación de Puntos"
    
    'Nombre cliente
    Printer.FontName = "Printer FontB 10cpi Tall"
    centra_printer vTarjeta.DNI
    centra_printer vTarjeta.NombreCompleto
    
    'Numero de tarjeta
    Printer.FontName = "Printer FontB 11cpi Tall"
    centra_printer EncriptarNumeroTarjeta(vTarjeta.NumeroTarjeta)

    Printer.CurrentX = 20
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.Print
    centra_printer "Fecha: " & Format$(Now, "dd/mm/YYYY") & Space(10) & "Hora: " & Format$(Now, "hh:nn:ss")
    
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = True
    centra_printer "Puntos Acumulados: " & CStr(vCantidadPuntos)
    
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = False
    centra_printer vTipoDocumento & " - " & vNroDocumento
    
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = False
    strMensaje = "Para saber su estado actual de puntos, consultar la pagina www.mifarma.com.pe"
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    
    strMensaje = "La acumulación de puntos se verá reflejada al día siguiente de la misma."
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    
'    Printer.FontSize = 6
'    Printer.CurrentY = Printer.CurrentY + intConstante
'    Printer.Print
'    strMensaje = objUsuario.Codigo & Space(10) & _
'                 App.Major & "." & App.Minor & "." & App.Revision & Space(9) & _
'                 "BTL" & objUsuario.CodigoLocal
'    centra_printer strMensaje
    
    Printer.Print
    Printer.EndDoc
   
    Screen.MousePointer = vbDefault
    MsgBox "Recoger el voucher de la cuponera", vbInformation, App.ProductName
    Exit Sub
Control:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub cmdReimprimir_Click()
    
    
    On Error GoTo CtrlErr
        If grdDocumento.ApproxCount = 0 Then
                    Exit Sub
        End If

    If (dbcTipoDoc.BoundText = COD_TIPO_TKB Or dbcTipoDoc.BoundText = COD_TIPO_TKF) And grdDocumento.ApproxCount > 1 Then
    
        If MsgBox("Desea Re imprimir TODO el rango de documentos ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                            
                With grdDocumento.DataSource
                    .MoveFirst
                    While Not .EOF
                        objDocumento.ImprimirDocumento dbcTipoDoc.BoundText, Replace(grdDocumento.Columns(0).Value, "-", ""), "", grdDocumento.Columns("COD_MODALIDAD_VENTA").Value
                        ''MsgBox "Tipo " + dbcTipoDoc.BoundText + " Nº " + grdDocumento.Columns(0).Value
                        .MoveNext
                    Wend
                End With
        Else
            If MsgBox("Desea Re imprimir SOLO el " + dbcTipoDoc.Text + " Nº " + grdDocumento.Columns(0).Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
                grdDocumento.SetFocus
                Exit Sub
            End If
            objDocumento.ImprimirDocumento dbcTipoDoc.BoundText, Replace(grdDocumento.Columns(0).Value, "-", ""), "", grdDocumento.Columns("COD_MODALIDAD_VENTA").Value
            ''MsgBox "Tipo " + dbcTipoDoc.BoundText + " Nº " + grdDocumento.Columns(0).Value
            grdDocumento.SetFocus
        End If
    


    Else
        If MsgBox("Desea Re imprimir la " + dbcTipoDoc.Text + " Nº " + grdDocumento.Columns(0).Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
            grdDocumento.SetFocus
            Exit Sub
        End If
'        If MsgBox("Desea reimprmir en Dolares", vbYesNo, App.ProductName) = vbYes Then
'            objDocumento.ImprimirDocumento dbcTipoDoc.BoundText, Replace(grdDocumento.Columns(0).Value, "-", ""), "", grdDocumento.Columns("COD_MODALIDAD_VENTA").Value, , , True
'        Else
            objDocumento.ImprimirDocumento dbcTipoDoc.BoundText, Replace(grdDocumento.Columns(0).Value, "-", ""), "", grdDocumento.Columns("COD_MODALIDAD_VENTA").Value
'        End If
        
        grdDocumento.SetFocus
    End If

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    
End Sub

'Private Sub cmdSalir_Click()
''    fraAsignar.Visible = False
'End Sub




Private Sub dbcTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdCancelar_Click
    End Select
End Sub


Private Sub dtpFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdBuscar.SetFocus
End Sub

Private Sub dtpFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpFechaFin.SetFocus
End Sub

Private Sub Form_Load()

If objDocumento.FnValidaReimprimir(objUsuario.Codigo) = 0 Then
    cmdReimprimir.Enabled = False
Else
    cmdReimprimir.Enabled = True
End If

setteaFormulario Me

Me.left = 0
Me.top = 0


        Set dbcTipoDoc.RowSource = objDocumento.ListaTipo(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
        dbcTipoDoc.ListField = "DESCRIPCION"
        dbcTipoDoc.BoundColumn = "CODIGO"
        
        dbcTipoDoc.BoundText = objVenta.CodigoDocumentoVenta


optTipoBusqueda(0).Value = True
dtpFechaIni.Value = objUsuario.sysdate
dtpFechaFin.Value = objUsuario.sysdate
'fraAsignar.Visible = False

Call SetGrid
Me.Caption = "Administrador"


End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
        Case vbKeyEscape
            cmdCancelar_Click
        Case vbKeyF1
            cmdReactivar_Click
        Case vbKeyF2
            'cmdAsignarUsuario_Click
        Case vbKeyF3
            cmdAnulacion_Click
        Case vbKeyF4
            cmdReimprimir_Click
        Case vbKeyF5
            cmdFormaPago_Click
            
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub activaNumeroDoc()
            dtpFechaIni.Enabled = False
            dtpFechaFin.Enabled = False
            txtNumDocIni.Enabled = True
            txtNumDocFin.Enabled = True
End Sub

Private Sub activaFechaDoc()
            dtpFechaIni.Enabled = True
            dtpFechaFin.Enabled = True
            txtNumDocIni.Enabled = False
            txtNumDocFin.Enabled = False
End Sub

Private Sub grdDocumento_DblClick()
   On Error GoTo Control

            If grdDocumento.ApproxCount = 0 Then Exit Sub
                    
                    frm_ADM_PreviewDoc.datos objUsuario.CodigoEmpresa, _
                    objUsuario.CodigoLocal, _
                    dbcTipoDoc.BoundText, _
                    grdDocumento.Columns("NUM_DOCUMENTO").Value, _
                    "", _
                    "" & grdDocumento.DataSource("COD_MODALIDAD_VENTA")
                    

Exit Sub
Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

Private Sub grdDocumento_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)

    Select Case Condition
        Case 0
            Select Case Col
                Case 0, 1, 3, 4
                    If grdDocumento.Columns(4).CellValue(Bookmark) = "ANU" Then
                        CellStyle.ForeColor = vbRed
                    End If
                Case 2
                    If grdDocumento.Columns(4).CellValue(Bookmark) = "ANU" Then
                        CellStyle.ForeColor = vbRed
                        CellStyle.Font.Bold = True
                    End If
            End Select
        Case 1, 2
            Select Case Col
                Case 0, 1, 3, 4
                    If grdDocumento.Columns(4).CellValue(Bookmark) = "ANU" Then
                        CellStyle.ForeColor = vbBlue
                    End If
                Case 2
                    If grdDocumento.Columns(4).CellValue(Bookmark) = "ANU" Then
                        CellStyle.ForeColor = vbBlue
                        CellStyle.Font.Bold = True
                    End If
            End Select
    End Select

End Sub

Private Sub grdDocumento_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
On Error GoTo handle


    If grdDocumento.Columns(4).CellValue(Bookmark) = "ANU" Then
        RowStyle.BackColor = RGB(252, 207, 213)
        RowStyle.ForeColor = RGB(0, 0, 0)
        grdDocumento.CambiaSeleccionado RGB(252, 207, 213)
        Exit Sub
    Else
        grdDocumento.CambiaSeleccionado RGB(222, 235, 254)
    End If
    
    If grdDocumento.Columns(4).CellValue(Bookmark) <> "ANU" Then
        RowStyle.BackColor = RGB(251, 242, 183)
        RowStyle.ForeColor = RGB(0, 0, 0)
        grdDocumento.CambiaSeleccionado RGB(251, 242, 183)
    Else
        grdDocumento.CambiaSeleccionado RGB(222, 235, 254)
    End If



Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName



End Sub

Private Sub grdDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call grdDocumento_DblClick
    End Select
End Sub

Private Sub optTipoBusqueda_Click(Index As Integer)
    If Index = 0 Then
        activaNumeroDoc
    Else
        activaFechaDoc
    End If
    
End Sub

Private Sub optTipoBusqueda_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Select Case Index
            Case 0
                txtNumDocIni.SetFocus
            Case 1
                dtpFechaIni.SetFocus
        End Select
    
    End If
End Sub

Private Sub txtNumDocFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdBuscar.SetFocus
End Sub


Private Sub SetGrid()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
    arrCampos = Array("NUM_DOCUMENTO", "FCH_REGISTRA", "MTO_TOTAL", "DES_RAZON_SOCIAL", "COD_ESTADO", "COD_USUARIO_DEPENDIENTE", "DES_ABREVIATURA", "NUM_AUTORIZA", "COD_TIPO_VENTA", "COD_MODALIDAD_VENTA")
    arrCaption = Array("Documento", "Emisión", "Importe", "Cliente", "Estad", "Usuario", "Modalidad", "Nº Autoriza", "Tipo Venta", "Modalidad")
    arrAncho = Array(1200, 1500, 900, 2400, 450, 3000, 1500, 1000, 1000, 600)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgRight, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgCenter)
    grdDocumento.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDocumento.Columns(0).FetchStyle = True
    grdDocumento.Columns(1).FetchStyle = True
    grdDocumento.Columns(2).FetchStyle = True
    grdDocumento.Columns(3).FetchStyle = True
    grdDocumento.Columns(4).FetchStyle = True
    grdDocumento.Col = 0
End Sub

Private Sub txtNumDocIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        txtNumDocFin.Text = txtNumDocIni.Text
        
    End If
End Sub
