VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_VTA_ConsCobServ 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlDataCombo dbcMoneda 
      Height          =   315
      Left            =   840
      TabIndex        =   17
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdAnulacion 
      Caption         =   "An&ular"
      Height          =   615
      Left            =   1320
      Picture         =   "frm_VTA_ConsCobServ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   1020
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   615
      Left            =   120
      Picture         =   "frm_VTA_ConsCobServ.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "&Buscar"
      Height          =   615
      Left            =   5760
      Picture         =   "frm_VTA_ConsCobServ.frx":060E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFchIni 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   61407233
      CurrentDate     =   39093
   End
   Begin vbp_Ventas.ctlGrilla grdServicios 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7011
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlDataCombo dbcServicios 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4200
      Picture         =   "frm_VTA_ConsCobServ.frx":0B98
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   5520
      Picture         =   "frm_VTA_ConsCobServ.frx":1122
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFchFin 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   61407233
      CurrentDate     =   39093
   End
   Begin vbp_Ventas.ctlDataCombo dbcUsuario 
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Top             =   1200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   900
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   1650
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1650
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   510
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
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
      Index           =   12
      Left            =   4320
      TabIndex        =   4
      Top             =   6840
      Width           =   900
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
      Left            =   6000
      TabIndex        =   3
      Top             =   6840
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Consulta de Cobranza por Servicios"
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
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3840
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frm_VTA_ConsCobServ.frx":16AC
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frm_VTA_ConsCobServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Dim objMoneda As New clsMoneda

Private Sub cmdAnulacion_Click()
Dim objDocumento As New clsDocumento
Dim Bookmark As Variant
    
    On Error GoTo CtrlErr
        If grdServicios.ApproxCount = 0 Then
                    Exit Sub
        End If


    If MsgBox("Desea anular el recibo Nº " + grdServicios.DataSource("NUM_DOCUMENTO").Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        grdServicios.SetFocus
        Exit Sub
    End If
    


    Bookmark = grdServicios.Bookmark

    If objVenta.EsNavsat(grdServicios.DataSource("COD_TIPO_SERVICIO").Value) = "1" Then
        If objVenta.AnulaTransaccionNavsat("" & grdServicios.DataSource("NUM_SUMINISTRO").Value, "" & grdServicios.DataSource("COD_PRODUCTO").Value, "" & grdServicios.DataSource("DOC_ASOC").Value, "" & grdServicios.DataSource("NUM_RECIBO").Value, "" & grdServicios.DataSource("NUM_VOUCH_OPE").Value) = False Then
            MsgBox "No se pudo anular la transaccion de NAVSAT", vbCritical, App.ProductName
            Exit Sub
        End If
    End If

    Dim gvarError As String
    Dim ValorRet As String
    gvarError = objDocumento.Anula(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, "REC", grdServicios.DataSource("NUM_DOCUMENTO").Value, grdServicios.DataSource("COD_ESTADO").Value, objUsuario.Codigo, ValorRet)
    
    If gvarError = "" Then
        CmdConsultar_Click
        MsgBox "Se anulo los siguientes documentos:" + Chr(13) + ValorRet, vbInformation, App.ProductName
        grdServicios.Bookmark = Bookmark
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    grdServicios.SetFocus
    Set objDocumento = Nothing
End If
    Exit Sub
CtrlErr:
     Set objDocumento = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub CmdConsultar_Click()
Dim rstTotLstSerCob As oraDynaset

    If dbcMoneda.BoundText = "" Then
        MsgBox "Debe indicar la moneda de la transacción", vbInformation + vbOKOnly, App.ProductName
        Exit Sub
    End If

    Set grdServicios.DataSource = objServicio.ListaSrvCob(objUsuario.CodigoEmpresa, _
                                                          objUsuario.CodigoLocal, _
                                                          CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                          CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")), _
                                                          dbcServicios.BoundText, _
                                                          dbcUsuario.BoundText, _
                                                          dbcMoneda.BoundText)
                                                          
                                                          
                                                              
    Set rstTotLstSerCob = objServicio.TotListaSrvCob(objUsuario.CodigoEmpresa, _
                                                          objUsuario.CodigoLocal, _
                                                          CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                          CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")), _
                                                          dbcServicios.BoundText, _
                                                          dbcUsuario.BoundText, _
                                                          dbcMoneda.BoundText)
                                                              
    If Not rstTotLstSerCob.EOF Then
        grdServicios.Columns(4).FooterText = IIf(IsNull(rstTotLstSerCob("TOT_IMP").Value), "0.00", rstTotLstSerCob("TOT_IMP").Value)
    End If
                                                          
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdModificar_Click()
On Error GoTo handle
    If grdServicios.ApproxCount = 0 Then Exit Sub
    
    frm_ADM_ModificaServ.strCodigoPadre = "" & grdServicios.DataSource("COD_TIPO_SERVICIO")
    frm_ADM_ModificaServ.strCodigoHijo = "" & grdServicios.DataSource("COD_SERVICIO")
    frm_ADM_ModificaServ.strNumSuministro = "" & grdServicios.DataSource("NUM_VOUCH_OPE")
    frm_ADM_ModificaServ.strTipoDocumento = "" & grdServicios.DataSource("cod_tipo_documento")
    frm_ADM_ModificaServ.strNumeroDocumento = "" & grdServicios.DataSource("Num_documento")
    frm_ADM_ModificaServ.Show
Exit Sub
handle:
MsgBox Err.Description, vbCritical, App.ProductName
    

End Sub

Private Sub Form_Load()
    SetteaFormulario Me
    Set dbcServicios.RowSource = objServicio.ListaTipoCons
    dbcServicios.ListField = "DES_TIPO_SERVICIO"
    dbcServicios.BoundColumn = "COD_TIPO_SERVICIO"
    
    Set dbcUsuario.RowSource = objUsuario.ListaCons("", objUsuario.CodigoLocal)
    dbcUsuario.ListField = "NOM_USUARIO"
    dbcUsuario.BoundColumn = "COD_USUARIO"
    
    
    Set dbcMoneda.RowSource = objMoneda.Lista("")
    dbcMoneda.ListField = "DES_MONEDA"
    dbcMoneda.BoundColumn = "COD_MONEDA"
    
    Call SeteaGrillaServicio
    dtpFchIni.Value = objUsuario.sysdate
    dtpFchFin.Value = objUsuario.sysdate
End Sub


Sub SeteaGrillaServicio()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("USU_COB", "SERVICIO", "NUM_DOC_COB", _
                      "MON", "IMP", _
                      "NUM_VOUCH_OPE", "#RECIBO", "#SUMIN", _
                      "COD_LIQUIDACION")
                      
    arrCaption = Array("Usuario", "Servicio", "#Doc", _
                       "Mon", "Imp.", "#Voucher", _
                       "#Recibo", "#Suministro", "#Liquidación")

    arrAncho = Array(2500, 1100, 1500, _
                     400, 700, 1100, _
                     1100, 1100, 1100)
                     
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, _
                          dbgLeft, dbgRight, dbgLeft, _
                          dbgLeft, dbgLeft, dbgLeft)
                          
    grdServicios.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdServicios.Columns(4).NumberFormat = "##0.#0"
    grdServicios.ColumnFooter = True
    grdServicios.Columns(2).FooterText = "Total"
    grdServicios.Columns(2).FooterDivider = False
    grdServicios.Columns(3).FooterText = "->"
    grdServicios.Columns(2).FooterAlignment = dbgRight
    grdServicios.Columns(4).FooterBackColor = &HC0E0FF
    
    
End Sub


