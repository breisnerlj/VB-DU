VERSION 5.00
Begin VB.Form frm_ADM_PreviewDoc 
   Caption         =   "Visualización de Documentos"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   Icon            =   "frm_ADM_PreviewDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla grdCobros 
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2990
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   5880
      Picture         =   "frm_ADM_PreviewDoc.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin vbp_Ventas.ctlDocumento ctlDocPreview 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9551
   End
End
Attribute VB_Name = "frm_ADM_PreviewDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento


Public Sub datos(ByVal pCia As String, ByVal pCodLocal As String, ByVal pCodTipoDoc As String, ByVal pNumDocumento As String, ByVal pCodigo As String, ByVal CodigoModalidad As String)
    ctlDocPreview.Cia = pCia
    ctlDocPreview.CodLocal = pCodLocal
    ctlDocPreview.CodTipoDocumento = pCodTipoDoc
    ctlDocPreview.NumDocumento = pNumDocumento
    ctlDocPreview.Codigo = pCodigo
    ctlDocPreview.Mostrar
    
    If (ctlDocPreview.xCodigoMagistral = "86520") Or (ctlDocPreview.xCodigoMagistral = "90287") Then
        setteaGrilla ""
        Set grdCobros.DataSource = objDocumento.ListaInsumos_x_Doc(pCia, pCodTipoDoc, Replace(pNumDocumento, "-", ""))
        Set objDocumento = Nothing
        Me.Height = 7710
    Else
        If CodigoModalidad = "004" Then
            setteaGrilla CodigoModalidad
            Set grdCobros.DataSource = objDocumento.ListaCobros(pCia, pCodTipoDoc, Replace(pNumDocumento, "-", ""))
            Set objDocumento = Nothing
            Me.Height = 7710
                
        '    ElseIf CodigoModalidad = "017" Then
        '        'Dim objDocumento As New clsDocumento
        '        setteaGrilla CodigoModalidad
        '        Set grdCobros.DataSource = objDocumento.ListaInsumos_x_Doc(pCia, pCodTipoDoc, Replace(pNumDocumento, "-", ""))
        '        Set objDocumento = Nothing
        '        Me.Height = 7710
        Else
            Me.Height = 5895
        End If
        'Me.Height = 5895
    End If
Me.Show vbModal
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub setteaGrilla(ByVal GrdGrilla As String)
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    If GrdGrilla = "004" Then
        grdCobros.Caption = ""
        arrCampos = Array("DES_MOTIVO_COBRO", "NOMBRE", "IMP_COBRO")
        arrCaption = Array("Motivo", "Nombre", "Importe")
        arrAncho = Array(3000, 3000, 1000)
        arrAlineacion = Array(vbAlignNone, vbAlignNone, vbAlignNone)
        grdCobros.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    ElseIf GrdGrilla = "" Then
        grdCobros.Caption = "Cotización Recetario Magistral"
        arrCampos = Array("COD_PRODUCTO_INS", "DES_PRODUCTO", "CTD_PRODUCTO", "CTD_BASE", "PCT_BASE", "IMP_PRECIO", "COD_UND_CAPACIDAD")
        arrCaption = Array("Insumo", "Descripción", "Cant", "Ctd Base", "Pct Base", "Precio", "Und Medida")
        arrAncho = Array(800, 2800, 800, 800, 800, 800, 1000)
        arrAlineacion = Array(vbAlignNone, vbAlignNone, vbAlignNone, vbAlignNone, vbAlignNone, vbAlignNone, vbAlignNone)
        grdCobros.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        grdCobros.Columns("COD_UND_CAPACIDAD").Visible = False
    End If
End Sub

