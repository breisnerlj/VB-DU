VERSION 5.00
Begin VB.Form frm_OFF_ConsultaDocumento 
   BorderStyle     =   0  'None
   Caption         =   "Consulta de Documentos"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbp_Ventas.ctlGrillaArray grdDocumento 
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtDocFin 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
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
   Begin vbp_Ventas.ctlTextBox txtDocIni 
      Height          =   315
      Left            =   2400
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
   Begin VB.ComboBox cboTipoDocumento 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   5880
      Picture         =   "frm_OFF_ConsultaDocumento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   900
   End
   Begin VB.CommandButton cmdReactivar 
      Caption         =   "&Reactivar"
      Height          =   615
      Left            =   120
      Picture         =   "frm_OFF_ConsultaDocumento.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   900
   End
   Begin VB.CommandButton cmdAnulacion 
      Caption         =   "An&ular"
      Height          =   615
      Left            =   1275
      Picture         =   "frm_OFF_ConsultaDocumento.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   900
   End
   Begin VB.CommandButton cmdReimprimir 
      Caption         =   "Re &Imprimir"
      Height          =   615
      Left            =   2430
      Picture         =   "frm_OFF_ConsultaDocumento.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   900
   End
   Begin VB.CommandButton cmdFormaPago 
      Caption         =   "&F. Pago"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3600
      Picture         =   "frm_OFF_ConsultaDocumento.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   900
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Detalle"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4725
      Picture         =   "frm_OFF_ConsultaDocumento.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   900
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   555
      Left            =   5400
      Picture         =   "frm_OFF_ConsultaDocumento.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Documento :"
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
      Left            =   120
      TabIndex        =   19
      Top             =   900
      Width           =   2070
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
      Left            =   6120
      TabIndex        =   18
      Top             =   6480
      Width           =   285
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
      TabIndex        =   17
      Top             =   6480
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
      Left            =   1635
      TabIndex        =   16
      Top             =   6480
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
      Left            =   2790
      TabIndex        =   15
      Top             =   6480
      Width           =   180
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
      Left            =   3930
      TabIndex        =   14
      Top             =   6480
      Width           =   180
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
      Left            =   4965
      TabIndex        =   13
      Top             =   6480
      Width           =   435
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
      Left            =   2655
      TabIndex        =   12
      Top             =   600
      Width           =   1545
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
      Left            =   4860
      TabIndex        =   11
      Top             =   600
      Width           =   1440
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
      TabIndex        =   10
      Top             =   180
      Width           =   1800
   End
End
Attribute VB_Name = "frm_OFF_ConsultaDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim xarr As XArrayDB



Private Sub cmdAnulacion_Click()
    Dim Bookmark As Variant
    Dim objDocumento As cls_OFF_Documento

    On Error GoTo CtrlErr
    If grdDocumento.ApproxCount = 0 Then Exit Sub
    
    If grdDocumento.Columns(5).Value = COD_ESTADO_ANU Then
        MsgBox "El Documento se encuentra anulado", vbInformation, App.ProductName
        grdDocumento.SetFocus
        Exit Sub
    End If
    
    If Trim(Month(CDate(grdDocumento.Columns(2).Value))) <> Trim(Str(Month(objOFFUsuario.sysdate))) Then
        MsgBox "No se puede anular un documento de otro mes", vbInformation, App.ProductName
        grdDocumento.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Desea anular la " + grdDocumento.Columns(0).Value + " Nº " + grdDocumento.Columns(1).Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        grdDocumento.SetFocus
        Exit Sub
    End If
    
    Bookmark = grdDocumento.Bookmark
    
    Set objDocumento = New cls_OFF_Documento
    If objDocumento.Anula(grdDocumento.Columns(0).Value, _
                                grdDocumento.Columns(1).Value) Then
        cmdBuscar_Click
        MsgBox "Se anuló el documento correctamente", vbInformation, App.ProductName
    Else
        MsgBox "No se pudo anular el documento", vbCritical, App.ProductName
    End If
    grdDocumento.Refresh
    grdDocumento.Bookmark = Bookmark
    grdDocumento.SetFocus

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo CtrlErr

    Dim strTipoDocumento As String
    Dim strNumDocIni As String
    Dim strNumDocFin As String

    If cboTipoDocumento.Text = "" Then
        MsgBox "El Tipo de Documento no ha sido seleccionado", vbCritical, App.ProductName
        cboTipoDocumento.SetFocus
        Exit Sub
    End If

    If txtDocIni.Text = "" Then
        MsgBox "El número del documento inicial no puede estar en blanco", vbCritical, App.ProductName
        txtDocIni.SetFocus
        Exit Sub
    End If
    
    If txtDocFin.Text = "" Then
        MsgBox "El número del documento final no puede estar en blanco", vbCritical, App.ProductName
        txtDocFin.SetFocus
        Exit Sub
    End If
    
    If Val(txtDocIni.Text) > Val(txtDocFin.Text) Then
        MsgBox "El número Inicial es mayor que el número final", vbCritical, App.ProductName
        txtDocIni.SetFocus
        Exit Sub
    End If

    strTipoDocumento = left(cboTipoDocumento.Text, 3)
    strNumDocIni = Replace(txtDocIni.Text, "-", "")
    strNumDocFin = Replace(txtDocFin.Text, "-", "")

    Call Buscar(strTipoDocumento, strNumDocIni, strNumDocFin)

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdReactivar_Click()
    Dim Bookmark As Variant
    Dim objDocumento As cls_OFF_Documento

    On Error GoTo CtrlErr
    If grdDocumento.ApproxCount = 0 Then Exit Sub
        
    If grdDocumento.Columns(5).Value = COD_ESTADO_EMI Then
        MsgBox "El Documento se encuentra activo", vbInformation, App.ProductName
        grdDocumento.SetFocus
        Exit Sub
    End If
    
    If Trim(Month(CDate(grdDocumento.Columns(2).Value))) <> Trim(Str(Month(objOFFUsuario.sysdate))) Then
        MsgBox "No se puede reactivar un documento de otro mes", vbInformation, App.ProductName
        grdDocumento.SetFocus
        Exit Sub
    End If
    
    
    If MsgBox("¿Desea reactivar la " + grdDocumento.Columns(0).Value + " Nº " + grdDocumento.Columns(1).Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        grdDocumento.SetFocus
        Exit Sub
    End If
    
    Bookmark = grdDocumento.Bookmark
    
    Set objDocumento = New cls_OFF_Documento
    If objDocumento.Reactiva(grdDocumento.Columns(0).Value, grdDocumento.Columns(1).Value) Then
        cmdBuscar_Click
        MsgBox "Se reactivó el documento correctamente", vbInformation, App.ProductName
    Else
        MsgBox "No se pudo reactivar el documento", vbCritical, App.ProductName
    End If
    grdDocumento.Refresh
    grdDocumento.Bookmark = Bookmark
    grdDocumento.SetFocus

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdReimprimir_Click()
Dim objDocumento As New cls_OFF_Documento

On Error GoTo CtrlErr

    If grdDocumento.ApproxCount = 0 Then Exit Sub


    If MsgBox("¿Desea re-imprimir la " + grdDocumento.Columns(0).Value + " Nº " + grdDocumento.Columns(1).Value + " ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
        grdDocumento.SetFocus
        Exit Sub
    End If

    Set objDocumento = New cls_OFF_Documento
    Call objDocumento.ImprimePorDocumento(grdDocumento.Columns(0).Value, grdDocumento.Columns(1).Value)
    Set objDocumento = Nothing



Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr

    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0

    Select Case KeyCode
        Case vbKeyF1
            cmdReactivar_Click
        Case vbKeyF3
            cmdAnulacion_Click
        Case vbKeyF4
            cmdReimprimir_Click
        Case vbKeyEscape
            cmdCancelar_Click
        ''''    DESDE ACA COPIA
        Case tmpCtrl And vbKeyQ And False
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyM And False
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyE And False
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyD And False
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyC
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF5
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF6
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF7
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF8 And False
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyX And False
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyF
            cmdCancelar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
    End Select

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()
setteaFormulario Me
Call CargarTipoDocumento
Call SetGrid

End Sub


Private Sub CargarTipoDocumento()
Dim i As Integer
Dim objDocumento As cls_OFF_Documento
Dim xTipoDocumento As XArrayDB


    Set objDocumento = New cls_OFF_Documento
    Set xTipoDocumento = objDocumento.ListaTipoDocumento
    Set objDocumento = Nothing

    For i = 0 To xTipoDocumento.UpperBound(1)
        cboTipoDocumento.AddItem xTipoDocumento(i, 0) & "-" & UCase(xTipoDocumento(i, 1))
    Next i
End Sub




Private Sub Buscar(ByVal pstrTipoDocumento As String, ByVal pstrNumDocumentoIni As String, pstrNumDocumentoFin As String)
'    Dim cnn As ADODB.Connection
    Dim strSQL As String
    Dim rsDetalleVenta As New ADODB.Recordset
    Dim objDocumento As cls_OFF_Documento
    Dim strTipDoc As String
    Dim strNumDoc As String
    Dim strFecEmi As String
    Dim dblImporte As Double
    Dim strNomCliente As String
    Dim strEstadoDoc As String
    Dim strUsuario As String
    Dim i As Integer
    Dim xTemp As XArrayDB

    On Error GoTo CtrlErr

    Screen.MousePointer = vbHourglass
        
    Set objDocumento = New cls_OFF_Documento

'        Set cnn = New ADODB.Connection
'        cnn.Open gstrConexion
        
'        strSQL = "select * from detalleventa.txt"
'        rsDetalleVenta.Open strSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText
    rsDetalleVenta.Open strDetalleVentaXML, gstrConexion, adOpenStatic, adLockOptimistic
    rsDetalleVenta.Filter = "COD_TIPO_DOCUMENTO = '" & pstrTipoDocumento & "' AND NUM_DOCUMENTO >= '" & pstrNumDocumentoIni & "' AND NUM_DOCUMENTO <= '" & pstrNumDocumentoFin & "'"
    If rsDetalleVenta.RecordCount > 0 Then
        rsDetalleVenta.MoveFirst
        dblImporte = 0
        objDocumento.LimpiaConsultaDoc
        While Not rsDetalleVenta.EOF
            strTipDoc = rsDetalleVenta!COD_TIPO_DOCUMENTO
            strNumDoc = rsDetalleVenta!NUM_DOCUMENTO
            strFecEmi = rsDetalleVenta!FCH_EMISION
            dblImporte = dblImporte + rsDetalleVenta!MTO_SUBTOTAL
            strNomCliente = IIf(IsNull(rsDetalleVenta!NOM_CLIENTE), "", rsDetalleVenta!NOM_CLIENTE)
            strEstadoDoc = rsDetalleVenta!COD_ESTADO
            strUsuario = objOFFUsuario.DesUsuarioVenta(rsDetalleVenta!USU_EMISION)
            grdDocumento.Array1 = objDocumento.AgregaConsultaDoc(strTipDoc, strNumDoc, strFecEmi, dblImporte, strNomCliente, strEstadoDoc, strUsuario)
            rsDetalleVenta.MoveNext
        Wend
            
        rsDetalleVenta.Filter = adFilterNone
        Set xTemp = objDocumento.ConsultaDoc
        For i = 0 To xTemp.UpperBound(1)
            rsDetalleVenta.Filter = "COD_TIPO_DOCUMENTO = '" & xTemp(i, 0) & "' AND NUM_DOCUMENTO = '" & xTemp(i, 1) & "' AND TIP_MOVIMIENTO = '" & COD_TIP_MOV_VENTA & "'"
            If rsDetalleVenta.RecordCount > 0 Then
                rsDetalleVenta.MoveFirst
                dblImporte = 0
                While Not rsDetalleVenta.EOF
                    dblImporte = dblImporte + rsDetalleVenta!MTO_SUBTOTAL
                    grdDocumento.Array1 = objDocumento.AgregaConsultaDoc(xTemp(i, 0), xTemp(i, 1), xTemp(i, 2), dblImporte, xTemp(i, 4), xTemp(i, 5), xTemp(i, 6))
                    rsDetalleVenta.MoveNext
                Wend
            Else
                MsgBox "No se pudo calcular el importe total del documento", vbCritical, App.ProductName
            End If
        Next i
        grdDocumento.Rebind
        grdDocumento.SetFocus
    Else
        MsgBox "No se encontró datos para el rango de documentos", vbCritical, App.ProductName
        grdDocumento.Limpiar
        txtDocIni.SetFocus
    End If
        
    rsDetalleVenta.Filter = adFilterNone
    rsDetalleVenta.Close
        'cnn.Close
        
    Set objDocumento = Nothing

    Screen.MousePointer = vbDefault
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    Screen.MousePointer = vbDefault

End Sub


Private Sub SetGrid()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
    arrCampos = Array("", "", "", "", "", "", "")
    arrCaption = Array("Tipo", "Documento", "Emisión", "Importe", "Cliente", "Estado", "Usuario")
    arrAncho = Array(0, 1200, 1500, 900, 2400, 450, 3000)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgRight, dbgLeft, dbgLeft, dbgLeft)
    grdDocumento.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDocumento.Columns(0).Visible = False
    grdDocumento.Columns(0).FetchStyle = True
    grdDocumento.Columns(1).FetchStyle = True
    grdDocumento.Columns(2).FetchStyle = True
    grdDocumento.Columns(3).FetchStyle = True
    grdDocumento.Columns(4).FetchStyle = True
    grdDocumento.Columns(5).FetchStyle = True
    grdDocumento.Columns(6).FetchStyle = True
    
    grdDocumento.Col = 0
    
End Sub


Private Sub grdDocumento_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)

    Select Case Condition
        Case 0
            Select Case Col
                Case 1, 2, 3, 5
                    If grdDocumento.Columns(5).CellValue(Bookmark) = COD_ESTADO_ANU Then
                        CellStyle.ForeColor = vbRed
                    End If
                Case 3
                    If grdDocumento.Columns(5).CellValue(Bookmark) = COD_ESTADO_ANU Then
                        CellStyle.ForeColor = vbRed
                        CellStyle.Font.Bold = True
                    End If
            End Select
        Case 1, 2
            Select Case Col
                Case 1, 2, 4, 5
                    If grdDocumento.Columns(5).CellValue(Bookmark) = COD_ESTADO_ANU Then
                        CellStyle.ForeColor = vbBlue
                    End If
                Case 3
                    If grdDocumento.Columns(5).CellValue(Bookmark) = COD_ESTADO_ANU Then
                        CellStyle.ForeColor = vbBlue
                        CellStyle.Font.Bold = True
                    End If
            End Select
    End Select

End Sub

'Private Sub txtDocFin_KeyPress(KeyAscii As Integer)
'
'    ' Don't disable the Esc or Backspace keys
'    If (KeyAscii = 27) Or (KeyAscii = 8) Then Exit Sub
'
'    ' Cancel user key input if it is not a letter or a digit
'    If (KeyAscii < 48 Or KeyAscii > 57) Then
'        KeyAscii = 0
'    End If
'
'End Sub
'
'Private Sub txtDocIni_KeyPress(KeyAscii As Integer)
'
'    ' Don't disable the Esc or Backspace keys
'    If (KeyAscii = 27) Or (KeyAscii = 8) Then Exit Sub
'
'    ' Cancel user key input if it is not a letter or a digit
'    If (KeyAscii < 48 Or KeyAscii > 57) Then
'        KeyAscii = 0
'    End If
'
'End Sub

Private Sub txtDocIni_LostFocus()
    txtDocFin.Text = Replace(txtDocIni.Text, "-", "")
End Sub
