VERSION 5.00
Begin VB.Form frm_VTA_ListaConvenio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Convenios"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   6840
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla grdConvenio 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8493
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_ListaConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCriterio As String
Dim objConvenio As New clsConvenio
Public out_DocumentoBeneficiario As String
Public out_DocumentoEmpresa As String
Public out_FlgBeneficiarios As String
Public out_PctBeneficiario As String
Public out_FlgMedico As String
Public out_FlgRepartidor As String
Public out_FlgPolitica As String
Public out_CodigoConvenio As String
Public out_NombreConvenio As String
Public out_FlgTipoConvenio As String
Public out_DireccionSocial As String
Public out_flg_valida_lincre As String
'Public out_flg_plan_vital As String  '-- Variable que indentifica si es plan vital 25/01/2010
Private Sub Form_Load()
On Error GoTo handle
    out_CodigoConvenio = ""
    Set grdConvenio.DataSource = objConvenio.ListaXLocal(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strCriterio)
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("COD_CONVENIO", "COD_CLIENTE", "DES_CONVENIO", "FLG_BENEFICIARIOS", "PCT_BENEFICIARIO", "FLG_REPARTIDOR", "FLG_MEDICO", "FLG_POLITICA", "PCT_BENEFICIARIO", "PCT_EMPRESA", "IMP_LINEA_CREDITO", "FLG_TIPO_CONVENIO", "FLG_ATENCION_TODOS_LOCALES", "FLG_PERIODO_VALIDEZ", "FCH_INICIO", "FCH_FIN", "DES_DIRECCION_SOCIAL", "FLG_VALIDA_LINCRE_BENEF")
    arrCaption = Array("Convenio", "Cliente", "Nombre", "Beneficiario", "PctBeneficiario", "Repartido", "Medico", "Politica", "%Benef.", "%Empresa", "Crédito", "Tipo", "Aten", "Verif.Fch", "FchIni", "FchFin", "Direccion", "FLG_VALIDA_LINCRE_BENEF")
    arrAncho = Array(1100, 0, 4500, 0, 0, 0, 0, 0, 600, 600, 1000, 600, 0, 0, 0, 0, 3000, 0)
    arrAlineacion = Array(dbgCenter, dbgGeneral, dbgLeft, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgRight, dbgRight, dbgRight, dbgRight, dbgCenter, dbgCenter, dbgRight, dbgRight, dbgLeft, dbgLeft)
    grdConvenio.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdConvenio.Columns(1).Visible = False
    grdConvenio.Columns(3).Visible = False
    grdConvenio.Columns(4).Visible = False
    grdConvenio.Columns(5).Visible = False
    grdConvenio.Columns(6).Visible = False
    grdConvenio.Columns(7).Visible = False
    grdConvenio.Columns(12).Visible = False
    grdConvenio.Columns(13).Visible = False
    grdConvenio.Columns(14).Visible = False
    grdConvenio.Columns(15).Visible = False
    grdConvenio.Columns(16).Visible = True
    grdConvenio.Columns(17).Visible = False
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdAceptar_Click()
'''''Dim rsConvenioXLocal As OraDynaset
'''''
'''''    If grdConvenio.Columns("FLG_ATENCION_TODOS_LOCALES").Value = "0" Then
'''''
'''''        Set objConvenio = New clsConvenio
'''''        Set rsConvenioXLocal = objConvenio.ListaXLocal(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
'''''        Set objConvenio = Nothing
'''''
'''''        If rsConvenioXLocal.EOF Then
'''''            MsgBox "El convenio " & grdConvenio.Columns("DES_CONVENIO").Value & " no puede ser atendido en este Local", vbInformation, App.ProductName
'''''            Exit Sub
'''''        End If
'''''    End If

    If grdConvenio.Columns("FLG_PERIODO_VALIDEZ").Value = "1" Then
        If grdConvenio.Columns("FCH_INICIO").Value < objUsuario.sysdate _
            And grdConvenio.Columns("FCH_FIN").Value < objUsuario.sysdate Then
            MsgBox "El convenio " & grdConvenio.Columns("DES_CONVENIO").Value & " Ya esta vencido", vbInformation, App.ProductName
            Exit Sub
        End If
    End If



    GuardaVariables
    'grdConvenio_DblClick
    Unload Me
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objConvenio = Nothing
End Sub

Private Sub grdConvenio_DblClick()
    cmdAceptar_Click
End Sub

Private Sub grdConvenio_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdConvenio_DblClick
    End Select
End Sub


Private Sub GuardaVariables()

On Error GoTo CtrlErr
    out_FlgBeneficiarios = "" & grdConvenio.Columns("FLG_BENEFICIARIOS").Value
    out_PctBeneficiario = "" & grdConvenio.Columns("PCT_BENEFICIARIO").Value
    out_FlgMedico = "" & grdConvenio.Columns("FLG_MEDICO").Value
    out_FlgRepartidor = "" & grdConvenio.Columns("FLG_REPARTIDOR").Value
    out_FlgPolitica = "" & grdConvenio.Columns("FLG_POLITICA").Value
    out_CodigoConvenio = "" & grdConvenio.Columns("COD_CONVENIO").Value
    out_NombreConvenio = "" & grdConvenio.Columns("DES_CONVENIO").Value
    out_FlgTipoConvenio = "" & grdConvenio.Columns("FLG_TIPO_CONVENIO").Value
    out_DocumentoBeneficiario = "" & grdConvenio.DataSource("COD_TIPDOC_BENEFICIARIO").Value
    out_DocumentoEmpresa = "" & grdConvenio.DataSource("COD_TIPDOC_CLIENTE").Value
    out_DireccionSocial = "" & grdConvenio.DataSource("DES_DIRECCION_SOCIAL").Value
    out_flg_valida_lincre = "" & grdConvenio.DataSource("FLG_VALIDA_LINCRE_BENEF").Value
    
    objVenta.flgBeneficiarios = out_FlgBeneficiarios
    objVenta.FlgValidaLinCre = out_flg_valida_lincre
    objVenta.FlgPlanVital = gclsOracle.FN_Valor("BTLPROD.PKG_CARGA_ARCHIVO_CNV.FN_DEV_CON_ARCHIVO", out_CodigoConvenio)  '-- 25/01/2010 Identifica el control de plan vital - crueda
    objVenta.NumMaximoUnidades = Val("" & grdConvenio.DataSource("CTD_MAX_UND_PRODUCTO").Value)
    
''''''    objVenta.PctBeneficiario
''''''    objVenta.FlgMedico
''''''    objVenta.FlgRepartidor
''''''    objVenta.FlgPolitica
''''''    objVenta.CodigoConvenio
''''''    objVenta.NombreConvenio
''''''    objVenta.FlgTipoConvenio
''''''
''''''
''''''
''''''    If objVenta.FlgTipoConvenio = "1" Then
''''''        'frmPedido.Label6.Visible = False
''''''        'frmPedido.lblTotal.Visible = False
''''''        frmPedido.Label4.Visible = True
''''''        frmPedido.lblPctCopago.Visible = True
''''''        frmPedido.Label8.Visible = True
''''''        frmPedido.lblcopago.Visible = True
''''''    Else
''''''        'frmPedido.Label6.Visible = True
''''''        'frmPedido.lblTotal.Visible = True
''''''        frmPedido.Label4.Visible = False
''''''        frmPedido.lblPctCopago.Visible = False
''''''        frmPedido.Label8.Visible = False
''''''        frmPedido.lblcopago.Visible = False
''''''    End If
    
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName


End Sub


