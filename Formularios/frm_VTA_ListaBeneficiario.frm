VERSION 5.00
Begin VB.Form frm_VTA_ListaBeneficiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Beneficiarios"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLineaCredito 
      Caption         =   "L. Credito"
      Height          =   435
      Left            =   5520
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox cboEstado 
      Height          =   315
      ItemData        =   "frm_VTA_ListaBeneficiario.frx":0000
      Left            =   840
      List            =   "frm_VTA_ListaBeneficiario.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4965
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   8190
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   6870
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla grdBeneficiario 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Tag             =   "."
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   8493
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   5025
      Width           =   495
   End
End
Attribute VB_Name = "frm_VTA_ListaBeneficiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCriterio As String
Public strCodConvenio As String
Public output_Codigo_Beneficiario As String
Public output_Nombre_Beneficiario As String
Public arrDatos As Variant
Dim objConvenio As New clsConvenio

Private Sub cboEstado_Click()
    consulta
End Sub

Private Sub cmdLineaCredito_Click()

    frmLineaCreditoNew.MuestraBeneficiario strCodConvenio, "" & grdBeneficiario.DataSource("COD_CLIENTE").Value
    consulta
End Sub

Private Sub Form_Load()
On Error GoTo handle
cboEstado.ListIndex = 1
    consulta
    'If strCriterio <> "" And strCriterio <> "%" Then strCriterio = Left(strCriterio, InStr(strCriterio, ",") - 1)
'    If Not strCodConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then ''CAMBIAR POR CONSTANTE EN RIMAC
Exit Sub
handle:
MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo handle
    grdBeneficiario_DblClick
    Unload Me
    Exit Sub
handle:
MsgBox Err.Description, vbCritical, App.ProductName

End Sub



Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub grdBeneficiario_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdBeneficiario_DblClick
    End Select
End Sub

Public Sub grdBeneficiario_DblClick()
Dim objBeneficiario As clsBeneficiario
Dim strCodCliente As String
Dim dblLineaCredito As Double
Dim LineaCred As Double
Dim Consumo As Double
Dim LineaCred2 As Double

On Error GoTo CtrlError

    If grdBeneficiario.Columns(5).Value = "0" Then
        MsgBox "El Beneficiario " & grdBeneficiario.Columns("DES_CLIENTE").Value & " está inactivo en la base de datos ", vbInformation + vbOKOnly, App.ProductName ''+ ", " + grdBeneficiario.Columns("DES_NOM_CLIENTE").Value & "
        Exit Sub
    End If
    
    dblLineaCredito = 0
    
    'If Not strCodConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then ''CAMBIAR POR CONSTANTE EN RIMAC
    If objConvenio.EsRimac(strCodConvenio) = False Then
    
    
            If objConvenio.EsDataRimac(strCodConvenio) = True Then
                Dim oraDatosNew As oraDynaset
                Set oraDatosNew = objConvenio.TransfierePaciente(strCodConvenio, grdBeneficiario.DataSource("COD_REFERENCIA").Value, grdBeneficiario.DataSource("DES_CLIENTE").Value)
        
                Set objBeneficiario = New clsBeneficiario
                
                strCodCliente = oraDatosNew("COD_CLIENTE").Value
                dblLineaCredito = Val(oraDatosNew("IMP_LINEA_CREDITO").Value)
'                objVenta.LineaCred = objBeneficiario.CreditoReal(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
'                objVenta.Consumo = objBeneficiario.Consumo(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
                LineaCred = objBeneficiario.CreditoReal(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
                Consumo = objBeneficiario.Consumo(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
                If objUsuario.CodLocalCallCenter = "1DLV" Then
                    LineaCred2 = LineaCred + Consumo 'esto por que el credito ya viene restado.
                    Consumo = Consumo + objBeneficiario.ConsumoUnif(strCodConvenio, strCodCliente, "10", "005")
                    LineaCred = LineaCred2 - Consumo
                End If
                objVenta.LineaCred = LineaCred
                objVenta.Consumo = Consumo
                Set objBeneficiario = Nothing
            
            
                output_Nombre_Beneficiario = oraDatosNew("DES_CLIENTE").Value '+ ", " + grdBeneficiario.Columns("DES_NOM_CLIENTE").Value
                output_Codigo_Beneficiario = oraDatosNew("COD_CLIENTE").Value
                objVenta.NumeroDocumentoID = "" & oraDatosNew("NUM_DOCUMENTO_ID").Value
                objVenta.NombreCliente = oraDatosNew("DES_CLIENTE").Value
                objVenta.DireccionCliente = "" & oraDatosNew("DES_DIRECCION_SOCIAL").Value
        
            Else
            
                Set objBeneficiario = New clsBeneficiario
                
                strCodCliente = grdBeneficiario.Columns("COD_CLIENTE").Value
                dblLineaCredito = Val(grdBeneficiario.Columns("IMP_LINEA_CREDITO").Value)
'                objVenta.LineaCred = objBeneficiario.CreditoReal(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
'                objVenta.Consumo = objBeneficiario.Consumo(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
                LineaCred = objBeneficiario.CreditoReal(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
                Consumo = objBeneficiario.Consumo(strCodConvenio, strCodCliente, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
                If objUsuario.CodLocalCallCenter = "1DLV" Then
                    LineaCred2 = LineaCred + Consumo 'esto por que el credito ya viene restado.
                    Consumo = Consumo + objBeneficiario.ConsumoUnif(strCodConvenio, strCodCliente, "10", "005")
                    LineaCred = LineaCred2 - Consumo
                End If
                objVenta.LineaCred = LineaCred
                objVenta.Consumo = Consumo
                
                Set objBeneficiario = Nothing
            
            
                output_Nombre_Beneficiario = grdBeneficiario.Columns("DES_CLIENTE").Value '+ ", " + grdBeneficiario.Columns("DES_NOM_CLIENTE").Value
                output_Codigo_Beneficiario = grdBeneficiario.Columns("COD_CLIENTE").Value
                objVenta.NumeroDocumentoID = "" & grdBeneficiario.Columns("NUM_DOCUMENTO_ID").Value
                objVenta.NombreCliente = grdBeneficiario.Columns("DES_CLIENTE").Value
                objVenta.DireccionCliente = "" & grdBeneficiario.Columns("DES_DIRECCION_SOCIAL").Value
            End If
    Else
        objVenta.LineaCred = "" & grdBeneficiario.Columns("CREDITO_REAL").Value
        
        arrDatos = Array(grdBeneficiario.DataSource(0).Value, _
                     grdBeneficiario.DataSource(1).Value, _
                     grdBeneficiario.DataSource(2).Value, _
                     grdBeneficiario.DataSource(3).Value, _
                     grdBeneficiario.DataSource(4).Value, _
                     grdBeneficiario.DataSource(5).Value, _
                     grdBeneficiario.DataSource(6).Value) 'grdBeneficiario.DataSource(7).Value
      
    End If
    
    
    
    Unload Me
    Exit Sub
    
CtrlError:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub
Sub consulta()
If objConvenio.EsRimac(strCodConvenio) = False And objConvenio.EsDataRimac(strCodConvenio) = False Then
        Set grdBeneficiario.DataSource = objConvenio.ListaBeneficiario(objUsuario.CodigoEmpresa, strCodConvenio, strCriterio, IIf(cboEstado.ListIndex = 2, "", cboEstado.ListIndex))
        Dim arrCampos As Variant
        Dim arrCaption As Variant
        Dim arrAncho As Variant
        Dim arrAlineacion As Variant
        arrCampos = Array("COD_REFERENCIA", "DES_CLIENTE", _
                          "IMP_LINEA_CREDITO", "IMP_CONSUMO", _
                          "CREDITO_REAL", "FLG_ACTIVO", _
                          "COD_CLIENTE", "NUM_DOCUMENTO_ID", _
                          "DES_DIRECCION_SOCIAL", "ESTADO")
        
        arrCaption = Array("Código", "Descripcion", _
                           "Crédito Inicial", "Consumo", _
                           "Saldo", "Activo", _
                           "Cliente", "DNI", _
                           "Dirección", "Activo")
                           
        arrAncho = Array(850, 4000, _
                         1500, 0, _
                         0, 0, _
                         0, 0, _
                         0, 800)
                         
        arrAlineacion = Array(vbAlignNone, vbAlignLeft, _
                              vbAlignNone, vbAlignNone, _
                              vbAlignNone, dbgCenter, _
                              dbgCenter, dbgLeft, _
                              dbgLeft, dbgCenter)
        
        grdBeneficiario.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdBeneficiario.Columns(3).Visible = False
        grdBeneficiario.Columns(4).Visible = False
        grdBeneficiario.Columns(5).Visible = False
        grdBeneficiario.Columns(6).Visible = False
        grdBeneficiario.Columns(7).Visible = False
        grdBeneficiario.Columns(8).Visible = False
    Else
        Set grdBeneficiario.DataSource = objConvenio.ListaPacienteRimac(strCriterio, objUsuario.CodigoLocal, strCodConvenio)
        Dim arrCamposX As Variant
        Dim arrCaptionX As Variant
        Dim arrAnchoX As Variant
        Dim arrAlineacionX As Variant
        
        arrCamposX = Array("POLIZA", "PLAN", _
                          "CODASEG", "N", _
                          "NOMBRE", "PRT", _
                          "NOMCONT", "ORG", _
                          "CREDITO_REAL")
        
        arrCaptionX = Array("Poliza", "Plan", _
                           "Cod. Aseg.", "Nª", _
                           "Nombre", "Prt.", _
                           "Nom. Cont.", "Org.", _
                           "Linea Cred")

        arrAnchoX = Array(800, 700, _
                         900, 500, _
                         4000, 500, _
                         2000, 500, _
                         900)
                         
        arrAlineacionX = Array(vbAlignNone, vbAlignLeft, _
                              vbAlignLeft, vbAlignLeft, _
                              vbAlignLeft, vbAlignLeft, _
                              dbgCenter, dbgLeft, _
                              dbgLeft)
        
        grdBeneficiario.FormatoGrilla arrCamposX, arrCaptionX, arrAnchoX, arrAlineacionX
    End If

End Sub

