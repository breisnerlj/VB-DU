VERSION 5.00
Begin VB.Form frm_VTA_Concepto_NotaCredito 
   BorderStyle     =   0  'None
   Caption         =   "F3"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlBuscaDocNew ctlBuscaDoc1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _extentx        =   5106
      _extenty        =   1931
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frm_VTA_Concepto_NotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lodynCab As oraDynaset
Dim lodynDet As oraDynaset
Dim lodyn1 As oraDynaset
Dim lodyn2 As oraDynaset

Private Sub BuscaDato(odynx1 As oraDynaset, odynx2 As oraDynaset)
On Error GoTo Handle1
    Set lodynCab = odynx1
    Set lodynDet = odynx2
    
    'Set lodyn1 = odynx1
    'Set lodyn2 = odynx2
     
    'If lodynCab.RecordCount = Nothing Then MsgBox "error", vbCritical, Caption: Exit Sub
     
    If lodynCab.RecordCount = 0 Then MsgBox "No existe el documento ", vbCritical, App.ProductName: Exit Sub
    
    
    'Carga las variables de referencia para cuando se grabe una nota de credito'
    objVenta.CodDocRef = lodynCab("COD_TIPO_DOCUMENTO").Value
    objVenta.NumDocRef = lodynCab("NUM_DOCUMENTO").Value
    objVenta.BtlRef = lodynCab("COD_LOCAL").Value
    objVenta.CodigoCliente = "" & lodynCab("COD_CLIENTE").Value
    objVenta.Ruc = "" & lodynCab("NUM_RUC_EMPRESA").Value
    objVenta.RazonSocial = "" & lodynCab("DES_RAZON_SOCIAL").Value
    objVenta.DesAuxCliNombre = "" & lodynCab("DES_AUX_CLI_NOMBRE").Value
    objVenta.DesAuxCliDirecc = "" & lodynCab("DES_AUX_CLI_DIRECC").Value
    objVenta.DesAuxCliTlf = "" & lodynCab("DES_AUX_CLI_TLF").Value
    'objVenta.CodigoConvenio = "" & lodynCab("COD_CONVENIO").Value
    '--------------------------------------------------------------------------'
    frm_VTA_NotaCredito.LblFchSysdate.Caption = Format(objUsuario.sysdate, "dd/mm/yyyy")
    
    frm_VTA_NotaCredito.TxtFactura.Text = lodynCab("NUM_DOCUMENTO").Value
    frm_VTA_NotaCredito.LblFactura.Caption = lodynCab("FCH_EMISION").Value
    frm_VTA_NotaCredito.TxtRazonSocial.Text = "" & lodynCab("DES_RAZON_SOCIAL").Value
    frm_VTA_NotaCredito.TxtRazonSocial.Enabled = IIf(frm_VTA_NotaCredito.TxtRazonSocial.Text = "", True, False)
    
    frm_VTA_NotaCredito.TxtDireccion.Text = "" & lodynCab("DES_AUX_CLI_DIRECC").Value
    frm_VTA_NotaCredito.TxtDireccion.Enabled = IIf(frm_VTA_NotaCredito.TxtDireccion.Text = "", True, False)
    
    frm_VTA_NotaCredito.TxtRuc.Text = "" & lodynCab("NUM_RUC_EMPRESA").Value
    frm_VTA_NotaCredito.TxtRuc.Enabled = IIf(frm_VTA_NotaCredito.TxtRuc.Text = "", True, False)
    
    frm_VTA_NotaCredito.TxtImpTot.Text = "" & lodynCab("MTO_TOTAL").Value
    'frm_VTA_NotaCredito.TxtImpReal.Text = "" & lodynCab("MTO_TOTAL").Value
    frm_VTA_NotaCredito.TxtImpReal.Text = "0.00"
        
    
        lodynDet.MoveFirst
    While Not lodynDet.EOF
            
        Call AgregaItem(lodynDet("COD_PRODUCTO").Value, _
                        lodynDet("DES_PRODUCTO").Value, _
                        lodynDet("CTD_PRODUCTO").Value, _
                        lodynDet("CTD_PRODUCTO_FRAC"), _
                        lodynDet("MTO_SUBTOTAL").Value, _
                        lodynDet("CTD_FRACCIONAMIENTO").Value, _
                        lodynDet("PRC_UNIT_NETO_VTA").Value, _
                        lodynDet("FLG_REGALO").Value)
        lodynDet.MoveNext
    Wend
    


    
    frm_VTA_NotaCredito.grdNC.Rebind
If frm_VTA_NotaCredito.TxtRazonSocial.Enabled = True Then frm_VTA_NotaCredito.TxtRazonSocial.Focus
    Unload Me
Exit Sub
Handle1:


    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ctlBuscaDoc1_RetornaDoc(odynCab As OracleInProcServer.oraDynaset, odynDet As OracleInProcServer.oraDynaset)
            
''    Set lodynCab = lodynCab
''    Set lodynDet = lodynDet
''
''    If lodynCab.RecordCount = 0 Then MsgBox "No existe el documento", vbCritical, Caption: Exit Sub
''
''    frm_VTA_NotaCredito.TxtFactura.Text = lodynCab("NUM_DOCUMENTO").Value
''    frm_VTA_NotaCredito.TxtRazonSocial.Text = "" & lodynCab("DES_RAZON_SOCIAL").Value
''    frm_VTA_NotaCredito.TxtDireccion.Text = "" & lodynCab("DES_DIRECCION_ENTREGA").Value
''    frm_VTA_NotaCredito.TxtRuc.Text = "" & lodynCab("NUM_RUC_EMPRESA").Value
''    frm_VTA_NotaCredito.TxtImpTot.Text = "" & lodynCab("IMP_TOTAL").Value
''    frm_VTA_NotaCredito.TxtImpReal.Text = "" & lodynCab("IMP_TOTAL").Value
''
''    lodynDet.MoveFirst
''    While Not lodynDet.EOF
''
''        Call AgregaItem(odynDet("COD_PRODUCTO").Value, _
''                        odynDet("DES_PRODUCTO").Value, _
''                        odynDet("CTD_UNIDADES").Value, _
''                        odynDet("CTD_FRACCIONES"), _
''                        odynDet("IMP_PRECIO_VENTA").Value)
''        lodynDet.MoveNext
''    Wend
''
''    frm_VTA_NotaCredito.grdNC.Rebind
''    Unload Me
    On Error GoTo CtrlErr
        BuscaDato odynCab, odynDet
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    End Sub

Sub psub_Items(ByRef rxdb As XArrayDB, ByVal intCol%)
    Dim i%
    For i = "0" & frm_VTA_NotaCredito.grdNC.Columns(0).Value To gxdbNC.UpperBound(1)
        gxdbNC(i, intCol) = i + 1
    Next i
End Sub

Sub AgregaItem(ByVal CodProd As String, ByVal DesProd As String, _
               ByVal ctdUnd As String, ByVal ctdFrac As String, _
               ByVal SubTotal As String, ByVal ctdFraccionamiento As String, _
               ByVal Precio As String, ByVal Regalo As String)
                       
    Dim encontro As Integer
    On Error GoTo CntError
            'comentado por PHERRERA  17/10/2008, cuando el doc. tenia productos repetidos (regalo promo) no los insertaba
            encontro = -1 'gxdbNC.Find(0, 1, CodProd)
    On Error GoTo 0
    GoTo OK
CntError:
    encontro = -1
    On Error GoTo 0
    GoTo OK
OK:
        If encontro = "-1" Then
            gxdbNC.AppendRows 1
            psub_Items gxdbNC, 0
            gxdbNC(gxdbNC.UpperBound(1), 1) = CodProd
            gxdbNC(gxdbNC.UpperBound(1), 2) = DesProd
            gxdbNC(gxdbNC.UpperBound(1), 3) = ctdUnd
            gxdbNC(gxdbNC.UpperBound(1), 4) = ctdFrac
            gxdbNC(gxdbNC.UpperBound(1), 5) = Precio
            gxdbNC(gxdbNC.UpperBound(1), 6) = SubTotal
            gxdbNC(gxdbNC.UpperBound(1), 7) = Empty                 'Cant Und Dev'
            gxdbNC(gxdbNC.UpperBound(1), 8) = Empty                 'Cant Fracc Dev'
            gxdbNC(gxdbNC.UpperBound(1), 9) = ctdFraccionamiento    'Cant. Fraccionamiento'
            gxdbNC(gxdbNC.UpperBound(1), 12) = Regalo               'flag regalo
        End If
End Sub




Private Sub Label4_Click(Index As Integer)

End Sub

Private Sub Form_Load()
        objVenta.LimpiaProductos
        ctlBuscaDoc1.CargarTipo
End Sub
