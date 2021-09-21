VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_FormaPagoDD 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_FormaPagoDD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskFecEmi 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin vbp_Ventas.ctlTextBox txtNroDoc 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   2820
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      MaxLength       =   10
      TABAuto         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlGrilla grdDocDescuento 
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      _ExtentX        =   13256
      _ExtentY        =   3413
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtNombre 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox txtDNI 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   4320
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Tipo            =   3
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox txtValor 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Alignment       =   1
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FormaPagoDD.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
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
      Index           =   11
      Left            =   6097
      TabIndex        =   15
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
      Index           =   12
      Left            =   4380
      TabIndex        =   14
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "DNI : "
      Height          =   255
      Left            =   420
      TabIndex        =   13
      Top             =   4350
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre : "
      Height          =   255
      Left            =   420
      TabIndex        =   12
      Top             =   3870
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   -1440
      X2              =   6360
      Y1              =   2715
      Y2              =   2715
   End
   Begin VB.Line Line2 
      X1              =   -1380
      X2              =   6360
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Documento de descuento"
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
      TabIndex        =   11
      Top             =   60
      Width           =   4515
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoDD.frx":0B14
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Número : "
      Height          =   255
      Left            =   420
      TabIndex        =   10
      Top             =   2850
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha  Emisión :"
      Height          =   255
      Left            =   420
      TabIndex        =   9
      Top             =   3390
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Valor S/. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4830
      Width           =   1575
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFormaPago As New clsFormaPago
Dim objDocPago As New clsDocumentoPago
Dim odynR1 As oraDynaset
Dim odynr2 As oraDynaset
Dim strDato As String
Dim strDatoDes As String
Dim strMoneda As String
Dim dblImpTotal As Double
''nuevas variables
Public od As oraDynaset
Public pstrDato As String
Public pstrDatoDes As String

Private Sub Form_Load()
On Error GoTo handle
    Me.top = 0
    Me.left = 0
    setteaFormulario Me
    SeteaGrilla
    Set od = objDocPago.Lista(objUsuario.CodigoLocal)
    Set grdDocDescuento.DataSource = od
    Set odynR1 = objFormaPago.ListaHijo(pstrDato)
    strDato = "" & odynR1("COD_HIJO").Value
    strDatoDes = "" & odynR1("DES_HIJO").Value
    strMoneda = "" & odynR1("COD_MONEDA").Value
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Public Sub aceptar()
If validaDocumento = False Then Exit Sub

On Error GoTo Control
         'If txtNroDoc.Text = "" Then MsgBox "Ingresa el numero de descuento", vbCritical, App.ProductName: txtNroDoc.SetFocus: Exit Sub
        'If mskFecEmi.Mask = "##/##/####" Then MsgBox "Ingrese Fecha de emisión de documento", vbCritical, Caption: mskFecEmi.SetFocus: Exit Sub
        If IsDate(mskFecEmi.Text) = False Then MsgBox "Ingrese Fecha de emisión de documento", vbCritical, App.ProductName: mskFecEmi.SetFocus: Exit Sub
        If txtNombre.Text = "" Then MsgBox "Ingrese el nombre del cliente", vbCritical, App.ProductName: txtNombre.SetFocus: Exit Sub
        If txtDNI.Text = "" Then MsgBox "Ingrese el dni del cliente", vbCritical, App.ProductName: txtDNI.SetFocus: Exit Sub
        If txtValor.Text = "" Then MsgBox "Ingrese el Importe del documento", vbCritical, App.ProductName: txtValor.SetFocus: Exit Sub
        
        objVenta.AgregaFormaPago pstrDato, _
                                 pstrDatoDes, _
                                 grdDocDescuento.Columns(0).Value, _
                                 grdDocDescuento.Columns(1).Value, _
                                 dblImpTotal, _
                                 "", IIf(strMoneda = "", "1", strMoneda), _
                                 "", "", _
                                 "", "", _
                                 0, "", _
                                 "", "", _
                                 "", txtNroDoc.Text, _
                                 "", "", _
                                 "", "", _
                                 mskFecEmi.Text, "", _
                                 "", txtNombre.Text, _
                                 txtDNI.Text, grdDocDescuento.Columns(0).Value

                                 frmPedido.Cal_Promo
    Unload Me
    '***************************************'
    'Arma el arreglo cada ez que se modifica'
      frm_VTA_FormaPago.SetFocus
      'frm_VTA_FormaPago.GrdListaFP.Array = objVenta.FormaPago
      frm_VTA_FormaPago.GrdListaFP.Rebind
    '***************************************'
    frmPedido.Cal_Montos
Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdAceptar_Click()
aceptar

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyEscape
            cmdCancelar_Click
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub SeteaGrilla()
On Error GoTo handle
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_HIJO", "DES_HIJO")
    arrCaption = Array("Codigo", "Vale")
    arrAncho = Array(900, 4500)
    arrAlineacion = Array(vbCenter, vbAlignLeft)
    grdDocDescuento.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim i%
    For i = 0 To grdDocDescuento.Columns.Count - 1
        grdDocDescuento.Columns(i).Visible = False
    Next i
    grdDocDescuento.Columns("COD_HIJO").Visible = True
    grdDocDescuento.Columns("DES_HIJO").Visible = True
    grdDocDescuento.Columns(1).WrapText = True
    'grdDocDescuento.RowHeight = 1.5 * grdDocDescuento.RowHeight
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub grdDocDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub grdDocDescuento_RegistroSeleccionado(ByVal DatoColumna0 As String)
     CargarValorFP
End Sub

Public Sub CargarValorFP()
   If Val("" & grdDocDescuento.DataSource("FLG_TIPO_VALOR").Value) = 1 Then
            txtValor.Text = "" & grdDocDescuento.DataSource("IMP_VALOR").Value
            txtValor.Enabled = False
        Else
            txtValor.Text = "0.00"
            txtValor.Enabled = True
        End If
            ''este linea es nueva para bloquear y desbloquear
    If Val("" & grdDocDescuento.DataSource("FLG_VARIABLE").Value) = 1 Then
        txtValor.Enabled = True
        Else
        txtValor.Enabled = False
    End If
End Sub


Private Sub mskFecEmi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub mskFecEmi_Validate(Cancel As Boolean)
''''On Error GoTo HANDLE
''''    Cancel = Not fbln_Valida_Fecha("MM/yyyy", "Error en el Ingreso de fechas", mskFecEmi.Text)
''''    If Cancel Then
''''        MsgBox "Error en el Ingreso de fechas", vbExclamation, Caption
''''        mskFecEmi.SetFocus
''''    End If
''''    Exit Sub
''''HANDLE:
''''    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    'txtDNI.Tipo = Entero
    'If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
'    txtNombre.Tipo = Mayusculas
'    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    'txtNroDoc.Tipo = Entero
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If Val("" & grdDocDescuento.DataSource("FLG_CONTROL_NUMERACION").Value) = 1 Then 'Indica que ira al detalle para validar el numero
            Dim objDocDest As New clsDocumentoPago
            Dim rsDocDest As oraDynaset
                Set rsDocDest = objDocDest.DetalleDocumento("" & grdDocDescuento.DataSource("NUM_DOCUMENTO_PAGO").Value, Trim(txtNroDoc.Text))
                If rsDocDest.RecordCount = 0 Then
                    MsgBox "El número de documento ingresado no existe o ya fue usado.", vbOKOnly + vbInformation, "Advertencia"
                    Set objDocDest = Nothing
                    KeyAscii = 0
                    txtNroDoc.SetFocus
                    Exit Sub
                End If
        End If
        SendKeys "{TAB}"
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtValor_Change()
On Error GoTo handle
    dblImpTotal = Val(txtValor.Text)
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
'    txtValor.Tipo = Real
'    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub ValidaNumeroCupon()
On Error GoTo handle
      gclsOracle.Num_Intentos = 1
      txtNroDoc.Text = objDocPago.validavale(txtNroDoc.Text)
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
Function validaDocumento() As Boolean
On Error GoTo handle
With grdDocDescuento

    'Valida que se haya selecionado alguna documento de pago
    If .ApproxCount < 0 Then validaDocumento = False: Exit Function
    
    'valida que el documento de pago se encuentre activo
    If Val(.DataSource("FLG_ACTIVO").Value) = 0 Then
        validaDocumento = False
        MsgBox "Esta documento de pago se encuentra inactivo, lo sentimos", vbCritical, App.ProductName
        Exit Function
    End If
    'Aca valido que las fecha no se hayan pasado
    
    Dim FechaInicio As String
    Dim FechaFin As String
    FechaInicio = "" & .DataSource("FCH_INICIO").Value
    FechaFin = "" & .DataSource("FCH_FIN").Value
    If IsDate(FechaInicio) = False Then
        validaDocumento = False
        MsgBox "La fecha inicio no es valida: " & FechaInicio, vbCritical, App.ProductName
        Exit Function
    End If
    If IsDate(FechaFin) = False Then
        validaDocumento = False
        MsgBox "La fecha fin no es valida: " & FechaFin, vbCritical, App.ProductName
        Exit Function
    End If

    If FechaInicio > objUsuario.sysdate Then 'valida que la fecha inicio sea sobrepasada
        validaDocumento = False
        MsgBox "Aun no es la fecha de actividad de est documento de descuento -->" & FechaInicio, vbCritical, App.ProductName
        Exit Function
    End If
    If FechaFin < objUsuario.sysdate Then 'Valida que no haya acabado el documento de descuento
        validaDocumento = False
        MsgBox "El documento de descuento ya ha sobrepasado la fecha fin -->" & FechaFin, vbCritical, App.ProductName
        Exit Function
    End If
    'Decide si va a validar o no el número de documento
    If Val(.DataSource("FLG_CONTROL_NUMERACION").Value) = 1 Then 'Indica que ira al detalle para validar el numero
        Dim objDocDest As New clsDocumentoPago
        Dim rsDocDest As oraDynaset
            Set rsDocDest = objDocDest.DetalleDocumento("" & grdDocDescuento.DataSource("NUM_DOCUMENTO_PAGO").Value, Trim(txtNroDoc.Text))
            If rsDocDest.RecordCount = 0 Then MsgBox "Este numero de documento no existe", vbCritical, App.ProductName: validaDocumento = False: txtNroDoc.Focus: Exit Function
        Set objDocDest = Nothing

    Else 'validara si exite o no
        If Trim("" & .DataSource("NUM_INICIO").Value) = "" And Trim("" & .DataSource("NUM_FIN").Value) = "" Then 'Si los dos son nulos no debo de validar nada
            'esto significa que no validara el numero de documento de pago
        Else
            Dim Inicio As Integer
            Dim FIN As Integer
            Inicio = Val("" & .DataSource("NUM_INICIO").Value)
            FIN = Val("" & .DataSource("NUM_FIN").Value)
            If Val(Inicio) > Val(txtNroDoc.Text) Then 'Aca valido que el numero sea mayor al rango minimo
                validaDocumento = False
                MsgBox "El número de documento es menor al rango establecido, lo sentimos", vbCritical, App.ProductName
                txtNroDoc.SetFocus
                Exit Function
            End If
            If Val(FIN) < Val(txtNroDoc.Text) Then 'Aca valido que el numero se menor al rango establecido
                validaDocumento = False
                MsgBox "El número de documento es mayor al rango establecido, lo sentimos", vbCritical, App.ProductName
                txtNroDoc.SetFocus
                Exit Function
            End If
        End If
        If Val("" & .DataSource("FLG_TIPO_VALOR").Value) = 1 Then
            If Val("" & .DataSource("IMP_VALOR").Value) = 0 Then
                validaDocumento = False
                MsgBox "La el importe no debe de ser 0, consulte con promociones", vbCritical, App.ProductName
                Exit Function
            End If
            txtValor.Text = "" & .DataSource("IMP_VALOR").Value
            txtValor.Enabled = False
        End If
    End If

    
End With
validaDocumento = True 'Solo llega a este punto cuando todo los datos estan realmente bien ingresados
    Exit Function
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Function
