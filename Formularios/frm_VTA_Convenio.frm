VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_Convenio 
   BorderStyle     =   0  'None
   Caption         =   $"frm_VTA_Convenio.frx":0000
   ClientHeight    =   7245
   ClientLeft      =   390
   ClientTop       =   780
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdlgPlanVital 
      Left            =   360
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCrgArch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuscarDiagnostico 
      Caption         =   "Bu&scar diagnóstico [Alt+S]"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   5040
      Width           =   2295
   End
   Begin VB.ListBox lstDestino 
      Height          =   840
      Left            =   3600
      TabIndex        =   13
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox txtFlgBeneficiario 
      Height          =   315
      Left            =   4080
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCodigoBeneficiario 
      Height          =   285
      Left            =   3600
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCodigoConvenio 
      Height          =   285
      Left            =   3120
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin vbp_Ventas.ctlGrillaArray ctlGrillaArray1 
      Height          =   1155
      Left            =   60
      TabIndex        =   7
      Top             =   3780
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2037
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.ComboBox cboTipoCopago 
      Height          =   315
      ItemData        =   "frm_VTA_Convenio.frx":016F
      Left            =   1080
      List            =   "frm_VTA_Convenio.frx":0179
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1740
      Width           =   1815
   End
   Begin VB.CommandButton cmdBuscarBeneficiario 
      Caption         =   "B&uscar"
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_Convenio.frx":0192
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_Convenio.frx":071C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrilla grdDocumentos 
      Height          =   1095
      Left            =   60
      TabIndex        =   6
      Top             =   2400
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1931
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdBuscarConvenio 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin vbp_Ventas.ctlTextBox txtConvenio 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   660
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      ColorDefault    =   -2147483634
      ColorDefault    =   -2147483634
      Tipo            =   2
      Bloqueado       =   -1  'True
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
   Begin vbp_Ventas.ctlTextBox txtCopago 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   1740
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Tipo            =   6
      MaxLength       =   7
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
   Begin vbp_Ventas.ctlTextBox txtBeneficiario 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      ColorDefault    =   -2147483643
      ColorDefault    =   -2147483643
      Tipo            =   2
      Bloqueado       =   -1  'True
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
   Begin vbp_Ventas.ctlDataCombo cboOrigen 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   5760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin MSMask.MaskEdBox txtFechaReceta 
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
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
   Begin vbp_Ventas.ctlDataCombo dbcMedico 
      Height          =   315
      Left            =   840
      TabIndex        =   9
      Top             =   5400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlDataCombo dbcRepartidor 
      Height          =   315
      Left            =   840
      TabIndex        =   8
      Top             =   5040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   46
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label lblConsumo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   45
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Saldo:"
      Height          =   195
      Left            =   5760
      TabIndex        =   44
      Top             =   1050
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Consumo:"
      Height          =   195
      Left            =   4200
      TabIndex        =   43
      Top             =   1050
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "/"
      Height          =   195
      Left            =   6240
      TabIndex        =   42
      Top             =   1800
      Width           =   75
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "L. Origen"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   5790
      Width           =   735
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Repartidor:"
      Height          =   195
      Index           =   11
      Left            =   0
      TabIndex        =   40
      Top             =   5100
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Medicos"
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   5460
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. Receta"
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   6180
      Width           =   705
   End
   Begin VB.Label lblTipoConvenio 
      AutoSize        =   -1  'True
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
      Left            =   3960
      TabIndex        =   36
      Top             =   0
      Width           =   2100
   End
   Begin VB.Label lblDocEmpresa 
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
      Left            =   5760
      TabIndex        =   33
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label lblDocBeneficiario 
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
      Left            =   6360
      TabIndex        =   32
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc. Emp / Benef. :"
      Height          =   195
      Left            =   4200
      TabIndex        =   31
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblBeneficiario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   30
      Top             =   1020
      Width           =   3195
   End
   Begin VB.Label lblConvenio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   29
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   9
      Left            =   120
      TabIndex        =   28
      Top             =   3525
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   8
      Left            =   120
      TabIndex        =   27
      Top             =   2145
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   26
      Top             =   1785
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   1365
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Beneficiario:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   24
      Top             =   1050
      Width           =   870
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   705
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Documentos a verificar para atender el convenio:"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   22
      Top             =   2160
      Width           =   3495
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Datos adicionales:"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   19
      Top             =   3540
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Copago:"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Convenio:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   390
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Convenio Institucional"
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
      Left            =   240
      TabIndex        =   16
      Top             =   60
      Width           =   2385
   End
End
Attribute VB_Name = "frm_VTA_Convenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************

                '"Programando un Mundo Mejor"

'********************************************************************************
Public pxdbDatos As New XArrayDB
Dim lstrDireccionSocial As String
Dim DireccRuta As String
Dim inFile1 As Integer              'para trabajar el arhivo
Dim inFile2 As Integer              'para trabajar el arhivo
'Dim lxdbArray As New XArrayDB
Dim objProducto As New clsProducto

'** Valores del detalle Plan Vital **'
Dim inp_CadNumPedido As String      'para capturar el número de archivo
Dim inp_CadFecAtencion As String    'para capturar el nombre del archivo
Dim inp_CadCodProducto As String    'para capturar la secuencia
Dim inp_CadCantidad As String       'para capturar la fecha de atención
        
Private Sub cmdAceptar_Click()
On Error GoTo Handle
ValidaConvenioBTL
ValidaPorcentaje
    If Graba_Convenio_Objeto = True Then
        'MsgBox "Me cerre"
        Unload Me
    End If
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBuscarBeneficiario_Click()
On Error GoTo Handle
If gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "NOBUSBEN0") = Trim(txtCodigoConvenio.Text) Then
    MsgBox "No se permite buscar beneficiarios en este convenio", vbCritical, App.ProductName
    Exit Sub
End If

If Len(Trim(txtBeneficiario.Text)) < 3 Then MsgBox "Debe ingresar como mínimo 3 caracteres.", vbOKOnly + vbExclamation, "Error": txtBeneficiario.SetFocus: Exit Sub
    With frm_VTA_ListaBeneficiario
        .strCriterio = ""
        .strCodConvenio = ""
        .strCriterio = Trim(txtBeneficiario.Text)
        .strCodConvenio = Trim(txtCodigoConvenio.Text)
        .output_Codigo_Beneficiario = ""
        .output_Nombre_Beneficiario = ""
        .Show vbModal
        txtBeneficiario.Text = ""
        txtCodigoBeneficiario = .output_Codigo_Beneficiario
        lblBeneficiario.Caption = .output_Nombre_Beneficiario
        lblConsumo.Caption = Format(objVenta.Consumo, "###,###.00")
        lblSaldo.Caption = Format(objVenta.LineaCred, "###,###.00")
        Dim objConv As New clsConvenio
        'If txtCodigoConvenio.Text = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then
        If objConv.EsRimac(txtCodigoConvenio.Text) = True Then
               On Error GoTo f
            Dim t As Integer
            While t <= UBound(.arrDatos)
                pxdbDatos(t, 5) = .arrDatos(t)
                If t = 4 Then
                    txtBeneficiario.Text = pxdbDatos(t, 5)
                    lblBeneficiario.Caption = pxdbDatos(t, 5)
                End If
                
             t = t + 1
            Wend
        ctlGrillaArray1.Rebind
f:

       End If

        cboTipoCopago.SetFocus
    End With
Exit Sub
Handle:
   MsgBox Err.Description, vbCritical, App.ProductName
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Inicializa_pantalla() 'Limpia todo el formulario
On Error GoTo Handle
    txtConvenio.Text = ""
    txtCodigoConvenio.Text = ""
    txtBeneficiario.Text = ""
    txtCodigoBeneficiario.Text = ""
    txtCopago.Text = ""
    grdDocumentos.Limpiar
    ctlGrillaArray1.Limpiar

    ctlGrillaArray1.Rebind
    
    grdDocumentos.Rebind
    
    dbcMedico.Visible = False
    dbcRepartidor.Visible = False
    cboTipoCopago.Enabled = True

    lblBeneficiario.Caption = ""

    lblConvenio.Caption = ""
    lblDocBeneficiario.Caption = ""
    lblDocEmpresa.Caption = ""
    
    cmdBuscarBeneficiario.Enabled = False
    
    lblTitle(11).Visible = False
    Label2.Visible = False
    Label6.Visible = False
    Label5.Visible = False
    cmdBuscarDiagnostico.Visible = False
    dbcRepartidor.Visible = False
    dbcMedico.Visible = False
    cboOrigen.Visible = False
    txtFechaReceta.Visible = False
    lstDestino.Visible = False
    
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub Carga_Convenio(ByVal CodigoConvenio As String, Optional Borra As Boolean = True)

On Error GoTo Handle
        If CodigoConvenio = "" Then MsgBox "Error al seleccionar el Convenio", vbCritical, App.ProductName: Exit Sub
        'Declaracion de las variable locales
        Dim rsConvenio As oraDynaset
        Dim rsDatosAdicionales As oraDynaset
        Dim objConvenio As New clsConvenio
        '
            Receta False
            Diagnostico False
            Repartidor False
            Medico False

        Set rsConvenio = objConvenio.Lista(CodigoConvenio)
        
        If Not txtCodigoConvenio.Text = CodigoConvenio Then
            txtBeneficiario.Text = ""
            ''''objVenta.NombreBeneficiario = ""
            ''''objVenta.CodigoBeneficiario = ""
            
            
            txtCodigoBeneficiario = ""
            lblBeneficiario.Caption = ""
        End If
        
        txtCodigoConvenio.Text = CodigoConvenio
        
        '** Valores que obtienen configuraciones generales 30/12/2008 **'
        
        objVenta.ImprimeImp = "" & rsConvenio("FLG_IMPRIME_IMPORTES").Value
        objVenta.PrecioMenor = "" & rsConvenio("FLG_PRECIO_MENOR").Value
        objVenta.PrecioMenorDeducible = "" & rsConvenio("FLG_PRECIO_DEDUCIBLE").Value
        
        '*********************************************************************************'
        
        'Carga los valores del convenio [LABEL]
        lblConvenio.Caption = "" & rsConvenio("DES_CONVENIO")
        lblDocBeneficiario.Caption = "" & rsConvenio("COD_TIPDOC_BENEFICIARIO")
        lblDocEmpresa.Caption = "" & rsConvenio("COD_TIPDOC_CLIENTE")
        lblTipoConvenio.Caption = "" & rsConvenio("TIPO_CONVENIO")
        
        '** Se asigna a estos valores para la grabación de los convenios que solo emitan 1 documento
        '** 04/02/2008 Por Cristhian Rueda
        If Trim(lblDocEmpresa.Caption) <> "" And Trim(lblDocBeneficiario.Caption) = "" Then
            objVenta.TipoCliente = "" & rsConvenio("FLG_TIPO_JURIDICA").Value
            objVenta.Ruc = "" & rsConvenio("NUM_DOCUMENTO_ID").Value
            objVenta.RazonSocial = "" & rsConvenio("CLIENTE").Value
        End If
        If "" & rsConvenio("FLG_TIPO_PAGO").Value = "0" Then
            cboTipoCopago.Locked = True
            cboTipoCopago.ListIndex = 0
        Else
            cboTipoCopago.Locked = True
            cboTipoCopago.ListIndex = 1
        End If
'''''        txtCodigoBeneficiario = "" & frm_VTA_ListaBeneficiario.output_Codigo_Beneficiario
'''''        lblBeneficiario.Caption = "" & frm_VTA_ListaBeneficiario.output_Nombre_Beneficiario
        
        'Identifica si es Politica Variable(0) o Fija (1)
        txtCopago.Text = "" & rsConvenio("PCT_BENEFICIARIO")
        If Val(rsConvenio("FLG_POLITICA")) = 1 Then
            txtCopago.Enabled = False
        Else
            cboTipoCopago.Enabled = True
            txtCopago.Enabled = True
        End If
        If Val(rsConvenio("FLG_DIAGNOSTICO")) = 1 Then
            
            cargaLista
            Receta True
            Diagnostico True
            If txtBeneficiario.Visible = True And txtBeneficiario.Enabled = True Then txtBeneficiario.SetFocus
            cmdBuscarBeneficiario.Enabled = True
        End If
        Dim objLocal As New clsLocal
        Set cboOrigen.RowSource = objLocal.ListaXConvenio(objUsuario.CodigoEmpresa, CodigoConvenio)
        cboOrigen.ListField = "DES_LOCAL_PROV"
        cboOrigen.BoundColumn = "cod_local"
        Set objLocal = Nothing
        
        'Evalua si lleva o no lista de Beneficiario
        txtFlgBeneficiario.Text = "" & rsConvenio("FLG_BENEFICIARIOS")
        If Val(txtFlgBeneficiario.Text) = 1 Then
            cmdBuscarBeneficiario.Enabled = True
            txtBeneficiario.Enabled = True
            
            On Error GoTo t
            If txtBeneficiario.Visible = True And txtBeneficiario.Enabled = True Then txtBeneficiario.SetFocus
t:
        Else
            cmdBuscarBeneficiario.Enabled = False
            txtBeneficiario.Enabled = False

            'If CodigoConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then ''CAMBIAR POR EL FLG TELEFONICA
            If objConvenio.EsRimac(CodigoConvenio) = True Then
                cmdBuscarBeneficiario.Enabled = True
                txtBeneficiario.Enabled = True
            End If
            
            On Error GoTo Y
            If txtCopago.Enabled = True Then
                '''txtCopago.SetFocus
            Else
                If txtFechaReceta.Visible = True Then
                    cmdBuscarBeneficiario.Enabled = True
                    '''txtBeneficiario.SetFocus
                Else
                    '''ctlGrillaArray1.SetFocus
                End If
            End If
Y:
        End If
        
        'Evalua si lleva o no una lista de Repartidores
        If Val(rsConvenio("FLG_REPARTIDOR")) = 1 Then
            Set dbcRepartidor.RowSource = objConvenio.ListaRepartidor(CodigoConvenio)
            dbcRepartidor.ListField = "DES_REPARTIDOR"
            dbcRepartidor.BoundColumn = "COD_REPARTIDOR"
            dbcRepartidor.Visible = True
            Repartidor True
            If CodigoConvenio <> "" Then dbcRepartidor.BoundText = objVenta.CodMotorizado
        Else
            Repartidor False
        End If
            
        'Evalua si lleva o no lleva una lista de Medicos
        If Val(rsConvenio("FLG_MEDICO")) = 1 Then
            Set dbcMedico.RowSource = objConvenio.ListaMedico(CodigoConvenio)
            dbcMedico.ListField = "NOMBRE"
            dbcMedico.BoundColumn = "COD_MEDICO"
            dbcMedico.Visible = True
            Medico True
            If CodigoConvenio <> "" Then dbcMedico.BoundText = objVenta.CodMedico
        Else
            Medico False
        End If

        'Calcula los importe de los productos para este convenio
        'ECASTILLO 22.06.2020
        If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = 0 Then
           frmPedido.ReCalculaPrecio Format("3", "000"), CodigoConvenio
        Else
           frmPedido.ReCalculaPrecio Format(ptmTipoPrecio, "000"), CodigoConvenio
        End If
        'frmPedido.ReCalculaPrecio Format(ptmTipoPrecio, "000"), CodigoConvenio
        
        'Consulta los documentos a verificara para el convenio
        Set grdDocumentos.DataSource = objConvenio.ListaDocVerif(CodigoConvenio)
        
        'Llena los datos adicionales del convenios
        If Borra = True Then
        Set rsDatosAdicionales = objConvenio.ListaTipoCampo(CodigoConvenio)
            If rsDatosAdicionales.RecordCount > 0 Then
                rsDatosAdicionales.MoveFirst
                pxdbDatos.ReDim 0, -1, 0, 6
                Dim ultimo As Integer
                While Not rsDatosAdicionales.EOF
                    ultimo = pxdbDatos.Count(1)
                    pxdbDatos.AppendRows
                    pxdbDatos(ultimo, 0) = "" & rsDatosAdicionales(0).Value
                    pxdbDatos(ultimo, 1) = "" & rsDatosAdicionales(1).Value
                    pxdbDatos(ultimo, 2) = "" & rsDatosAdicionales(2).Value
                    pxdbDatos(ultimo, 3) = "" & rsDatosAdicionales(3).Value
                    pxdbDatos(ultimo, 4) = "" & rsDatosAdicionales(4).Value
                    pxdbDatos(ultimo, 5) = "" '& rsDatosAdicionales(5).Value
                    pxdbDatos(ultimo, 6) = "" & rsDatosAdicionales("FLG_EDITABLE").Value
                    rsDatosAdicionales.MoveNext
                Wend
            End If
        End If
        ctlGrillaArray1.Array1 = pxdbDatos
        ctlGrillaArray1.Rebind
        
        
        'Liberando el objeto llamado
        Set objConvenio = Nothing
    Exit Sub
Handle:
        Set objConvenio = Nothing
        MsgBox Err.Description, vbCritical, App.ProductName
End Sub
Private Function Graba_Convenio_Objeto() As Boolean
 Dim i As Integer
 Dim DatosOk As Boolean
    Graba_Convenio_Objeto = False
With objVenta
        
        'Validaciones
         If dbcMedico.Visible = True And dbcMedico.BoundText = "" Then MsgBox "Debe seleccionar un Medico de la lista", vbCritical, App.ProductName: dbcMedico.SetFocus: Exit Function
         If dbcRepartidor.Visible = True And dbcRepartidor.BoundText = "" Then MsgBox "Debe seleccionar un Motorizado de la lista", vbCritical, App.ProductName: dbcRepartidor.SetFocus: Exit Function
         If cboOrigen.Visible = True And cboOrigen.BoundText = "" Then MsgBox "Debe seleccionar un local origen de la lista", vbCritical, App.ProductName: cboOrigen.SetFocus: Exit Function
         
         If txtCodigoConvenio.Text = "" Then MsgBox "Seleccion un convenio", vbCritical, App.ProductName: Exit Function
         If lstDestino.Visible = True And lstDestino.ListCount = 0 Then MsgBox "Seleccione un diagnostico ", vbCritical, App.ProductName: lstDestino.SetFocus: Exit Function
         Dim objConv As New clsConvenio
         
         'If Val(txtFlgBeneficiario.Text) = 1 And Not txtCodigoConvenio.Text = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then
         If Val(txtFlgBeneficiario.Text) = 1 Then
            If txtCodigoBeneficiario.Text = "" Then MsgBox "Debe seleccionar un Beneficiario", vbCritical, App.ProductName: Exit Function
         End If
         If objConv.EsRimac(txtCodigoConvenio.Text) And Trim(txtBeneficiario.Text) = "" Then
            MsgBox "Debe seleccionar un Beneficiario", vbCritical, App.ProductName: Exit Function
         End If
         
        'Carga en objeto
        .CodigoConvenio = txtCodigoConvenio.Text
        ''If Not objUsuario.EsDelivery Then
            .CodigoCliente = txtCodigoBeneficiario.Text
            .bk_codBeneficiario = .CodigoCliente 'ECASTILLO 17.12.2020
        ''End If
        .CodigoBeneficiario = txtCodigoBeneficiario.Text
        .NombreConvenio = lblConvenio.Caption
        .DesAuxCliDirecc = lstrDireccionSocial
        .Out_Direccion = lstrDireccionSocial
        .NombreBeneficiario = lblBeneficiario.Caption
        .DesAuxCliNombre = lblBeneficiario.Caption
        .PctBeneficiario = Val(txtCopago.Text)
        .ImpPctBeneficiario = Val(txtCopago.Text)
        .flgPctBeneficiario = cboTipoCopago.ListIndex
        .CodMotorizado = dbcRepartidor.BoundText
        .CodMotorizadoConv = dbcRepartidor.BoundText
        .CodMedico = dbcMedico.BoundText
        If Not IsDate(txtFechaReceta.Text) And txtFechaReceta.Visible = True Then
            MsgBox "La fecha no es valida ", vbCritical, App.ProductName
            Graba_Convenio_Objeto = False
            Exit Function
        End If
        
        .FechaReceta = txtFechaReceta.Text
        Dim strDiagnostico As String
        .LocalReceta = cboOrigen.BoundText
        Dim h As Integer
        While h < lstDestino.ListCount
        Dim strValor As String
        strValor = lstDestino.List(h)
        .AgregaDiagnostico fncPalote(strValor, , "-"), fncPalote(strValor, 1, "-")
        strDiagnostico = strDiagnostico & fncPalote(strValor, , "-") & "|"
        h = h + 1
        Wend
        
        .DetalleReceta = strDiagnostico

    DatosOk = True
    Dim strMensajeFaltante As String
    If pxdbDatos.Count(1) > 0 Then
    For i = 0 To pxdbDatos.UpperBound(1)
        If pxdbDatos(i, 5) = "" Then
            If pxdbDatos(i, 6) = "1" Then
                strMensajeFaltante = strMensajeFaltante & ", " & pxdbDatos(i, 1)
            End If
           DatosOk = False
        Else
            If pxdbDatos(i, 4) > 0 Then
                If Len(pxdbDatos(i, 5)) < pxdbDatos(i, 4) Then
                    DatosOk = True
                End If
            End If
        End If
    Next i
    Set .DatosAdicional = pxdbDatos
    End If
    
    If Not DatosOk Then
        MsgBox "Error en el registro de Datos Adicionales" + Chr(13) + "Falta ingresar o el dato es incorrecto, completar " & strMensajeFaltante, vbCritical, App.ProductName
        ctlGrillaArray1.SetFocus
        Exit Function
    End If
    
    'Calcula los montos de la Pre-Venta
    
    ''' ***** 24/03/2007
    frmPedido.Cal_Promo
    ''' *****
    
    'Agregado por JLOPEZ para mostrar mensaje de productos que no estan en el petitorio
    
        If frmPedido.pstrdxPrcCero <> "" Then
            Dim w As Integer
            Dim X As String
            Dim varMsgPrc As String
            
            varMsgPrc = ""
            X = ""
            For w = 1 To Len(frmPedido.pstrdxPrcCero)
              If Mid(frmPedido.pstrdxPrcCero, w, 1) = "|" Then
                    varMsgPrc = varMsgPrc & X & Chr(13)
                    X = ""
              Else
                    X = Mid(frmPedido.pstrdxPrcCero, w, 1)
                    varMsgPrc = varMsgPrc & X
              End If
            Next w
            
            
            frmPedido.grdPedido.Refresh
            
            'MsgBox "Productos excluidos por no pertencer al petitorio del convenio" & Chr(13) & _
                    varMsgPrc, vbInformation, "Productos Excluidos"
        End If
    
    
    '-------------
    frmPedido.Cal_Montos
    End With
    Graba_Convenio_Objeto = True
        Exit Function
Handle:
    Graba_Convenio_Objeto = False
    MsgBox Err.Description, vbCritical, App.ProductName
End Function
Private Sub Carga_Convenio_Objeto(CodigoConvenio As String)

With objVenta
        
        cmdBuscarBeneficiario.Enabled = False
        lblBeneficiario.Caption = .NombreBeneficiario
        txtCodigoConvenio.Text = objVenta.CodigoConvenio
        txtConvenio.Text = objVenta.NombreConvenio
        txtCodigoBeneficiario.Text = objVenta.CodigoBeneficiario
        txtBeneficiario.Text = objVenta.NombreBeneficiario
        If objVenta.flgPctBeneficiario = 0 Then
            txtCopago.Text = objVenta.PctBeneficiario
        Else
            txtCopago.Text = objVenta.ImpPctBeneficiario
        End If
        objVenta.CodigoCliente = objVenta.CodigoBeneficiario
        lblConvenio.Caption = objVenta.NombreConvenio
        Debug.Print objVenta.NombreBeneficiario & "<---"
        lblBeneficiario.Caption = objVenta.NombreBeneficiario
        
        txtFechaReceta.Mask = ""
        txtFechaReceta.Text = objVenta.FechaReceta
        txtFechaReceta.Mask = "##/##/####"
        cboOrigen.BoundText = objVenta.LocalReceta
        
        ctlGrillaArray1.Array1 = pxdbDatos
        ctlGrillaArray1.Rebind
        Dim d As Integer
        While d < .Diagnostico.Count(1)
            lstDestino.AddItem .Diagnostico(d, 0) & "-" & .Diagnostico(d, 1)
            d = d + 1
        Wend
        
    End With
    
        Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

'Valida que el porcentaje se Valido
Private Sub ValidaPorcentaje()
        If cboTipoCopago.Text = "PORCENTAJE" Then
            If Val(txtCopago.Text) > 100 Then
                txtCopago.selection
                Err.Raise -2100, App.ProductName, "El porcentaje no puede ser mayor al 100%"
            End If
        End If

End Sub
'Evento del Combo de Tipo de Co-Pago
Private Sub cboTipoCopago_Click()
    On Error GoTo Handle
        If cboTipoCopago.ListIndex = 0 Then
            txtCopago.Tipo = Porcentaje
            txtCopago.MaxLength = 3
            txtCopago.Text = 0
        Else
            txtCopago.Tipo = Real
            txtCopago.MaxLength = 0
            txtCopago.Text = 0
        End If
    
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
Private Sub cboTipoCopago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
'Evento del boton buscar Convenio
Private Sub cmdBuscarConvenio_Click()
On Error GoTo Control
If Len(Trim(txtConvenio.Text)) < 3 Then MsgBox "Debe ingresar como mínimo 3 caracteres.", vbOKOnly + vbExclamation, "Error": txtConvenio.SetFocus: Exit Sub

    With frm_VTA_ListaConvenio
        .strCriterio = txtConvenio.Text
        .Show vbModal
        If Not txtCodigoConvenio.Text = .out_CodigoConvenio Then pxdbDatos.ReDim 0, -1, 0, 6: Inicializa_pantalla
        lstrDireccionSocial = .out_DireccionSocial
        Carga_Convenio .out_CodigoConvenio
        
            If objVenta.FlgPlanVital = "1" And objUsuario.EsDelivery = True Then
                cmdCrgArch.Visible = True
              Else
                cmdCrgArch.Visible = False
            End If
        
    End With

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.number
End Sub

Private Sub cmdBuscarDiagnostico_Click()
With frm_VTA_BuscaDiagnostico
    
    .OUTPUT_CODIGO = ""
    .OUTPUT_NOMBRE = ""
    .Show vbModal
    If Not .OUTPUT_CODIGO = "" Then
    lstDestino.AddItem .OUTPUT_CODIGO & "-" & .OUTPUT_NOMBRE
    End If
End With
End Sub

Private Sub cmdCrgArch_Click()
   
    Dim inp_NumArchivo As String        'para capturar el número de archivo
    Dim inp_NomArchivo As String        'para capturar el nombre del archivo
    
    Dim inpArchivo As String
    Dim inpRuta As String
    Dim RutaCab As String
    Dim RutaDet As String
    Dim Cargo As Boolean
   On Error GoTo Control

    DireccRuta = ExaminarArch
    txtFlgBeneficiario.Text = 1
    
    inFile1 = FreeFile
    inFile2 = FreeFile
    
    inp_NomArchivo = Trim(Mid(right(DireccRuta, 23), 1, 19))
    inp_NumArchivo = Trim(Mid(right(DireccRuta, 23), 2, 18))
    
    inpRuta = left(DireccRuta, 13)
    inpArchivo = right(DireccRuta, 22)
    
    inpRuta = Replace(DireccRuta, StrReverse(fncGuion(StrReverse(DireccRuta), 0, "\")), "")
    inpArchivo = Mid(StrReverse(fncGuion(StrReverse(DireccRuta), 0, "\")), 2, Len(StrReverse(fncGuion(StrReverse(DireccRuta), 0, "\"))) - 1)
    
    
    objVenta.NomArchVital = inp_NomArchivo
    
    '** Carga Cabecera **'
    RutaCab = inpRuta & "C" & inpArchivo
    Open RutaCab For Input Access Read As #inFile1
        CadenaCab Cargo
    Close #inFile1
    If Cargo = False Then
        GoTo Control
    End If
    '** Carga Detalle **'
    RutaDet = inpRuta & "D" & inpArchivo
    Open RutaDet For Input Access Read As #inFile2
        CadenaDet
    Close #inFile2
    
    
    Dim GrabaVital As String
    Dim strMensaje As String
    
    GrabaVital = objVenta.Graba_Carga_CNV(gclsOracle.ODataBase, objVenta.NomArchVital, objUsuario.Codigo, strMensaje, _
                                          objVenta.FecAteVital, objVenta.NumAteVital, objVenta.NombBenef, _
                                          objVenta.ApePatBenef, objVenta.ApeMatBenef, objVenta.CodCentralVital, _
                                          objVenta.NumPedidoVital, objVenta.CodPostalVital, objVenta.DirBenef, _
                                          objVenta.DesRefVital, Replace(objVenta.FonoBenefVital, "_", ""), objVenta.SexoBenefVital, _
                                          objVenta.CtdPedidoVital, objVenta.NumPolizaVital, objVenta.NumCertifVital, _
                                          objVenta.CodParentesco, objVenta.CodDeducibleVital, objVenta.TipoFacturaVital, _
                                          objVenta.FecNacVital, objVenta.CodInternoVital, objVenta.DesObsVital, _
                                          objVenta.PctCoPagoVital, Trim(txtCodigoConvenio.Text), objVenta.CMPVital, objVenta.CIE10Vital, objVenta.LocalOrigenVital)
                                                                                       
    If strMensaje = "" Then
        FileCopy RutaCab, RutaCab & ".dat"
        FileCopy RutaDet, RutaDet & ".dat"
        Kill RutaCab
        Kill RutaDet
        MsgBox "Se cargo satisfactoriamente el archivo", vbInformation
    Else
        MsgBox strMensaje, vbCritical, App.ProductName
        Exit Sub
    End If
    
    Dim odynCabVital As oraDynaset
    Dim odynDetVital As oraDynaset
    Dim strCodLocal As String
    
    Set odynCabVital = objVenta.DevCabVital
    
    strCodLocal = mdiPrincipal.ctlCliente1.LocalAsignado
    If strCodLocal = "" Then strCodLocal = objUsuario.CodigoLocal
    Set odynDetVital = objVenta.DevDetVital(strCodLocal)
    
    '** Llena las clases del sistema de ventas **'
    
    objVenta.NombreBeneficiario = "" & odynCabVital("DES_CLIENTE").Value
    objVenta.NombreCliente = "" & odynCabVital("DES_CLIENTE").Value
    objVenta.DireccionCliente = "" & odynCabVital("DIR_BENEFICIARIO").Value
    objVenta.PctBeneficiario = Val("" & odynCabVital("PCT_COPAGO").Value)
    
    lblBeneficiario.Caption = "" & odynCabVital("DES_CLIENTE").Value
    txtCodigoBeneficiario.Text = "" & odynCabVital("COD_CLIENTE").Value
    frmPedido.lblPctCopago.Caption = "" & odynCabVital("PCT_COPAGO").Value
    mdiPrincipal.ctlCliente1.Codigo = "" & odynCabVital("COD_CLIENTE").Value
    mdiPrincipal.ctlCliente1.ConsultaCliente odynCabVital("COD_CLIENTE").Value
    ''ESTOS NO SE GRABABAN BIEN
    objVenta.CodigoBeneficiario = "" & odynCabVital("COD_CLIENTE").Value
    objVenta.CodigoCliente = "" & odynCabVital("COD_CLIENTE").Value
    objVenta.CodigoClienteDLV = "" & odynCabVital("COD_CLIENTE").Value
    objVenta.DesAuxCliDirecc = "" & odynCabVital("DIR_BENEFICIARIO").Value
    objVenta.DesAuxCliNombre = "" & odynCabVital("DES_CLIENTE").Value
    objVenta.DesAuxCliTlf = "" & odynCabVital("FONO_BENEFICIARIO").Value
    objVenta.DireccionCliente = "" & odynCabVital("DIR_BENEFICIARIO").Value
    mdiPrincipal.txtDireccion.Text = "" & odynCabVital("DIR_BENEFICIARIO").Value
    objVenta.DireccionClienteDLV = "" & odynCabVital("DIR_BENEFICIARIO").Value
    objVenta.NombreCliente = "" & odynCabVital("DES_CLIENTE").Value
    objVenta.NombBenef = "" & odynCabVital("DES_CLIENTE").Value
    objVenta.NombreClienteDLV = "" & odynCabVital("DES_CLIENTE").Value
    objVenta.CodigoTipoVenta = Venta_Convenio
    
    objVenta.DesUrbanizacionDLV = ""
    objVenta.DesReferenciaCli = "" & odynCabVital("DES_REFERENCIA").Value
    Dim objUbigeo As New clsUbigeo
    objVenta.DesDistritoDLV = "" & objUbigeo.DevDistrito("" & odynCabVital("COD_POSTAL").Value)
    objVenta.UbigeoEntrega = "" & objUbigeo.DevUbigeo("" & odynCabVital("COD_POSTAL").Value)
    '** Estos es en para llenar los datos adicionales **
    
        txtCopago.Text = objVenta.PctCoPagoVital
        
        Dim rsDatosAdicionales As oraDynaset
        Dim objConvenio As New clsConvenio

        Set rsDatosAdicionales = objConvenio.ListaTipoCampo(txtCodigoConvenio.Text)
        If rsDatosAdicionales.RecordCount > 0 Then
            rsDatosAdicionales.MoveFirst
            pxdbDatos.ReDim 0, -1, 0, 6
            Dim ultimo As Integer
            While Not rsDatosAdicionales.EOF
                ultimo = pxdbDatos.Count(1)
                pxdbDatos.AppendRows
                pxdbDatos(ultimo, 0) = "" & rsDatosAdicionales(0).Value
                pxdbDatos(ultimo, 1) = "" & rsDatosAdicionales(1).Value
                pxdbDatos(ultimo, 2) = "" & rsDatosAdicionales(2).Value
                pxdbDatos(ultimo, 3) = "" & rsDatosAdicionales(3).Value
                pxdbDatos(ultimo, 4) = "" & rsDatosAdicionales(4).Value
                If "" & rsDatosAdicionales(0).Value = "0000000031" Then pxdbDatos(ultimo, 5) = "" & objVenta.NumAteVital
                If "" & rsDatosAdicionales(0).Value = "0000000051" Then pxdbDatos(ultimo, 5) = "" & objVenta.TipoFacturaVital
                pxdbDatos(ultimo, 6) = "" & rsDatosAdicionales("FLG_EDITABLE").Value
                rsDatosAdicionales.MoveNext
            Wend
        End If
        
        If txtFechaReceta.Visible = True Then txtFechaReceta.Text = "" & odynCabVital("FCH_RECETA").Value
        If dbcMedico.Visible = True Then dbcMedico.BoundText = "" & odynCabVital("COD_MEDICO").Value
        If cboOrigen.Visible = True Then cboOrigen.BoundText = "" & odynCabVital("COD_LOCAL_ORIGEN").Value
        If Not "" & odynCabVital("CIE10").Value = "" Then
            Dim arrcie10() As String
            arrcie10 = Split("" & odynCabVital("CIE10").Value, ",")
            Dim ix  As Integer
            ix = 0
            Dim strDescri As String
            For ix = LBound(arrcie10) To UBound(arrcie10)
                strDescri = gclsOracle.FN_Valor("BTLPROD.PKG_CARGA_ARCHIVO_CNV.FN_DEV_CIE10", arrcie10(ix))
                lstDestino.AddItem strDescri
            Next
        End If
    
        ctlGrillaArray1.Array1 = pxdbDatos
        ctlGrillaArray1.Refresh
''        If cmdCrgArch.Visible = True And objVenta.NumAteVital <> "" Then pxdbDatos(ultimo, 5) = objVenta.NumAteVital
''        ctlGrillaArray1.Columns(5).Value = pxdbDatos(ultimo, 5)
        ctlGrillaArray1.Rebind

    
    
    '** Llena las clases del sistema de ventas **'
    
    odynDetVital.MoveFirst
    While Not odynDetVital.EOF
    
        Dim PctComi As Double
        PctComi = objProducto.pctComision(odynDetVital("COD_PROD_PROV").Value, strCodLocal, Format(ptmTipoPrecio, "000"))
    
        Dim oraDato As oraDynaset
        If objUsuario.CodLocalCallCenter = "1DLV" Then 'ECASTILLO 22.06.2020
            Set oraDato = objProducto.ListaDato("94", strCodLocal, "001", odynDetVital("COD_PROD_PROV").Value, odynDetVital("CTD_PRODUCTO").Value, odynDetVital("FLG_FRACCIONAMIENTO").Value, "", objUsuario.CodLocalCallCenter)
        Else
            Set oraDato = objProducto.ListaDato(objUsuario.CodigoEmpresa, strCodLocal, "001", odynDetVital("COD_PROD_PROV").Value, odynDetVital("CTD_PRODUCTO").Value, odynDetVital("FLG_FRACCIONAMIENTO").Value, "", objUsuario.CodLocalCallCenter)
        End If
        frmPedido.grdPedido.Limpiar
    
        frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(odynDetVital("COD_PROD_PROV").Value, odynDetVital("DES_PRODUCTO").Value, Val(odynDetVital("CTD_PRODUCTO").Value), odynDetVital("FLG_FRACCIONAMIENTO").Value, Val(odynDetVital("PRC_CONVENIO").Value), "000", "0", , , , , , "", PctComi, "", "", , Val(odynDetVital("PRC_PUBLICO").Value))
        ctlGrillaArray1.Limpiar
        odynDetVital.MoveNext
        
    Wend
    ctlGrillaArray1.Rebind
    frmPedido.Cal_Montos
    frmPedido.grdPedido.Rebind
    
    
    
   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.number
    
End Sub

Private Function ExaminarArch() As String
 On Error GoTo CtrlErr

    Dim RutaArchivo As String
    
    With frm_VTA_Convenio.cdlgPlanVital
        .DefaultExt = ".txt"
        .Filter = "Archivos de Texto (*.txt)|*.txt"
        .FilterIndex = 1
        .FileName = "C*.txt"
        .ShowOpen
        RutaArchivo = .FileName
    End With
    
    If Trim(RutaArchivo) = "" Or Len(Trim(RutaArchivo)) < 2 Then Exit Function
    If Dir(Trim(RutaArchivo)) = "" Then Exit Function
    
    ExaminarArch = RutaArchivo
    
    Exit Function
CtrlErr:
    Select Case Err.number
    Case 32755  'cancel del common dialog
        
    Case Else
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.number
    End Select
    
End Function

Private Sub CadenaCab(ByRef ConError As Boolean)
   On Error GoTo Control
    
'        Dim inp_FecAtencion As String       'para capturar la fecha de atención
'        Dim inp_NumAtencion As String       'para capturar el número de atencion
'        Dim inp_NomBenef As String          'para capturar el nombre del beneficiario
'        Dim inp_ApePatBenef As String       'para capturar el apellido Paterno del beneficiario
'        Dim inp_ApeMatBenef As String       'para capturar el apellido materno del beneficiario
'        Dim inp_CodCentral As String        'para capturar el código central
'        Dim inp_NumPedido As String         'para capturar el número de pedido
'        Dim inp_CodPostal As String         'para capturar el código de postal
'        Dim inp_DirBenef As String          'para capturar la dirección del beneficiario
'        Dim inp_DesReferencia As String     'para capturar la descripción de la referencia
'        Dim inp_FonoBenef As String         'para capturar el fono del beneficiario
'        Dim inp_SexoBenef As String         'para capturar el sexo del beneficiario
'        Dim inp_CtdPedido As String         'para capturar la cantidad del pedido
'        Dim inp_NumPoliza As String         'para capturar el número de poliza
'        Dim inp_NumCertifi As String        'para capturar el número de certificado
'        Dim inp_CodParentesco As String     'para capturar el código de parentesco
'        Dim inp_CodPlanDeducible As String  'para capturar el código del plan deducible
'        Dim inp_TipoFacturacion As String   'para capturar el tipo de facturación
'        Dim inp_FecNacimiento As String     'para capturar la fecha de nacimiento
'        Dim inp_CodInterno As String        'para capturar el código interno
'        Dim inp_DesObservacion As String    'para capturar la descripción de la observación
'        Dim inp_PctCoPago As String         'para capturar el porcentaje del copago
        
        Dim Cadena As String
        Dim TestArray() As String
        Dim TestArrayVal(26) As Variant
        Dim byInd As Byte
        
        Line Input #inFile1, Cadena
'''nuevo para los campos que envian con comas
        Dim i, n As Integer
        n = 26
        i = 0
        While (i < n)
                If Mid(Cadena, 1, 1) = Chr(34) Then
                    TestArrayVal(i) = fncGuion(fncGuion(Cadena, 1, Chr(34)), 0, Chr(34))
                    Cadena = fncGuion(fncGuion(fncGuion(Cadena, 1, Chr(34)), 1, Chr(34)), 1, Chr(44))
                Else
                    TestArrayVal(i) = fncGuion(Cadena, 0, Chr(44))
                    Cadena = fncGuion(Cadena, 1, Chr(44))
                End If
                i = i + 1
        Wend
'''nuevo para los campos que envian con comas

        
'        byInd = 0
'        TestArray() = Split(CADENA, ",")
'        While TestArray(byInd) <> ""
'            TestArrayVal(byInd) = Trim(TestArray(byInd))
'            byInd = byInd + 1
'            If byInd = 22 Then GoTo Sigue
'        Wend
        
Sigue:
        '** ARMA LA CADENA DE LA CABECERA DEL ARCHIVO PLAN VITAL **'
        
        objVenta.FecAteVital = Replace(TestArrayVal(0), """", "")
        objVenta.NumAteVital = Replace(TestArrayVal(1), """", "")
        objVenta.NombBenef = Replace(TestArrayVal(2), """", "")
        objVenta.ApePatBenef = Replace(TestArrayVal(3), """", "")
        objVenta.ApeMatBenef = Replace(TestArrayVal(4), """", "")
        objVenta.CodCentralVital = Replace(TestArrayVal(5), """", "")
        objVenta.NumPedidoVital = Replace(TestArrayVal(6), """", "")
        objVenta.CodPostalVital = Replace(TestArrayVal(7), """", "")
        objVenta.DirBenef = Replace(TestArrayVal(8), """", "")
        objVenta.DesRefVital = Replace(TestArrayVal(9), """", "")
        objVenta.FonoBenefVital = Replace(TestArrayVal(10), """", "")
        objVenta.SexoBenefVital = Replace(TestArrayVal(11), """", "")
        objVenta.CtdPedidoVital = Replace(TestArrayVal(12), """", "")
        objVenta.NumPolizaVital = Replace(TestArrayVal(13), """", "")
        objVenta.NumCertifVital = Replace(TestArrayVal(14), """", "")
        objVenta.CodParentesco = Replace(TestArrayVal(15), """", "")
        objVenta.CodDeducibleVital = Replace(TestArrayVal(16), """", "")
        objVenta.TipoFacturaVital = Replace(TestArrayVal(17), """", "")
        objVenta.FecNacVital = Replace(TestArrayVal(18), """", "")
        objVenta.CodInternoVital = Replace(TestArrayVal(19), """", "")
        objVenta.DesObsVital = Replace(TestArrayVal(20), """", "")
        objVenta.PctCoPagoVital = CDbl(Replace(TestArrayVal(21), """", ""))
            objVenta.CMPVital = TestArrayVal(22)
            objVenta.CIE10Vital = TestArrayVal(23)
            objVenta.LocalOrigenVital = TestArrayVal(24)
            objVenta.CodigoConvenioVital = TestArrayVal(25)
            If objVenta.CodigoConvenioVital <> txtCodigoConvenio.Text Then
                MsgBox "Este archivo no pertenece al convenio seleccionado", vbCritical, App.ProductName
                ConError = False
                Exit Sub
            End If
            ConError = True
''''''''''''''''        txtCopago.Text = objVenta.PctCoPagoVital
''''''''''''''''
''''''''''''''''        Dim rsDatosAdicionales As oraDynaset
''''''''''''''''        Dim objConvenio As New clsConvenio
''''''''''''''''
''''''''''''''''        Set rsDatosAdicionales = objConvenio.ListaTipoCampo(txtCodigoConvenio.Text)
''''''''''''''''        If rsDatosAdicionales.RecordCount > 0 Then
''''''''''''''''            rsDatosAdicionales.MoveFirst
''''''''''''''''            pxdbDatos.ReDim 0, -1, 0, 6
''''''''''''''''            Dim ultimo As Integer
''''''''''''''''            While Not rsDatosAdicionales.EOF
''''''''''''''''                ultimo = pxdbDatos.Count(1)
''''''''''''''''                pxdbDatos.AppendRows
''''''''''''''''                pxdbDatos(ultimo, 0) = "" & rsDatosAdicionales(0).Value
''''''''''''''''                pxdbDatos(ultimo, 1) = "" & rsDatosAdicionales(1).Value
''''''''''''''''                pxdbDatos(ultimo, 2) = "" & rsDatosAdicionales(2).Value
''''''''''''''''                pxdbDatos(ultimo, 3) = "" & rsDatosAdicionales(3).Value
''''''''''''''''                pxdbDatos(ultimo, 4) = "" & rsDatosAdicionales(4).Value
''''''''''''''''                pxdbDatos(ultimo, 5) = "" '& rsDatosAdicionales(5).Value
''''''''''''''''                pxdbDatos(ultimo, 6) = "" & rsDatosAdicionales("FLG_EDITABLE").Value
''''''''''''''''                rsDatosAdicionales.MoveNext
''''''''''''''''            Wend
''''''''''''''''        End If
''''''''''''''''        ctlGrillaArray1.Array1 = pxdbDatos
''''''''''''''''        ctlGrillaArray1.Rebind
'''''''''''''''''
'''''''''''''''''
'''''''''''''''''        frm_VTA_DatoAdicional.subDatos objVenta.NombreConvenio, _
'''''''''''''''''                        ctlGrillaArray1.Columns(0).Value, _
'''''''''''''''''                        ctlGrillaArray1.Columns(1).Value, _
'''''''''''''''''                        ctlGrillaArray1.Columns(2).Value, _
'''''''''''''''''                        ctlGrillaArray1.Columns(3).Value, _
'''''''''''''''''                        ctlGrillaArray1.Columns(4).Value, _
'''''''''''''''''                        ctlGrillaArray1.Columns(5).Value
''''''''''''''''
''''''''''''''''        If cmdCrgArch.Visible = True And objVenta.NumAteVital <> "" Then pxdbDatos(ultimo, 5) = objVenta.NumAteVital
''''''''''''''''        ctlGrillaArray1.Columns(5).Value = pxdbDatos(ultimo, 5)
''''''''''''''''        'ctlGrillaArray1.Array1 = pxdbDatos
''''''''''''''''        ctlGrillaArray1.Rebind
   
   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.number
End Sub

Private Sub CadenaDet()
    On Error GoTo Control
               
        
        Dim Cadena As String
        Dim TestArray() As String
        Dim TestArrayVal(4) As String
        Dim byInd As Byte
        Dim conteo As Integer
        
        'lxdbArray.ReDim 0, -1, 0, 3
        objVenta.PlanVitalArray.ReDim 0, -1, 0, 3
        
        While Not EOF(inFile2)
            
            Line Input #inFile2, Cadena
            byInd = 0
            TestArray() = Split(Cadena, ",", -1)
            While TestArray(byInd) <> ""
                TestArrayVal(byInd) = TestArray(byInd)
                byInd = byInd + 1
                If byInd = 4 Then GoTo Sigue
            Wend
            
Sigue:
            '** ARMA LA CADENA EL DETALLE DEL ARCHIVO PLAN VITAL **'
            
            inp_CadNumPedido = Replace(TestArrayVal(0), """", "")
            inp_CadFecAtencion = Replace(TestArrayVal(1), """", "")
            inp_CadCodProducto = Replace(TestArrayVal(2), """", "")
            inp_CadCantidad = Replace(TestArrayVal(3), """", "")
            
            objVenta.PlanVitalArray.AppendRows 1
            objVenta.PlanVitalArray(conteo, 0) = inp_CadNumPedido
            objVenta.PlanVitalArray(conteo, 1) = inp_CadFecAtencion
            objVenta.PlanVitalArray(conteo, 2) = inp_CadCodProducto
            objVenta.PlanVitalArray(conteo, 3) = inp_CadCantidad
            conteo = conteo + 1
            
        Wend
        
   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.number
End Sub

'Eventos del formulario que ejecutan lo shortcut
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Control
        psub_KeyDownAplicacion KeyCode, Shift
        
        Select Case KeyCode
            Case vbKeyF1
                txtConvenio.SetFocus
            Case vbKeyF2
                txtBeneficiario.SetFocus
            Case vbKeyF3
                If objVenta.FlgPolitica = "0" Then txtCopago.SetFocus
            Case vbKeyF4
                grdDocumentos.SetFocus
            Case vbKeyF9
                ctlGrillaArray1.SetFocus
            Case vbKeyEscape
                cmdCancelar_Click
            Case vbKeyReturn
                If Shift = 1 Then cmdAceptar_Click
        End Select
    Exit Sub
Control:
        MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub lstDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
        Dim k As Integer
    k = 0
    Dim final As Integer
    final = lstDestino.ListCount
    While k < final

        If lstDestino.Selected(k) = True Then
            lstDestino.RemoveItem (k)
            k = k - 1
        End If
        final = lstDestino.ListCount
        k = k + 1
    Wend

End If
End Sub

'''Private Sub lstOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
'''
'''    If KeyCode = vbKeyAdd Then
'''            Dim k As Integer
'''    k = 0
'''    Dim final As Integer
'''    final = lstOrigen.ListCount
'''    While k < final
'''
'''        If lstOrigen.Selected(k) = True Then
'''            lstDestino.AddItem lstOrigen.List(k)
'''            lstOrigen.RemoveItem (k)
'''            k = k - 1
'''        End If
'''        final = lstOrigen.ListCount
'''        k = k + 1
'''    Wend
'''
'''    End If
'''End Sub

Private Sub txtCopago_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    'txtCopago.Tipo = Entero
    If KeyAscii = 13 Then
        ValidaPorcentaje
    End If
Exit Sub
Handle:
   MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub ctlGrillaArray1_DblClick()
On Error GoTo Handle
    If ctlGrillaArray1.ApproxCount > 0 Then
        If Not ctlGrillaArray1.Columns(6).Value = "0" Then
        frm_VTA_DatoAdicional.subDatos objVenta.NombreConvenio, _
                        ctlGrillaArray1.Columns(0).Value, _
                        ctlGrillaArray1.Columns(1).Value, _
                        ctlGrillaArray1.Columns(2).Value, _
                        ctlGrillaArray1.Columns(3).Value, _
                        ctlGrillaArray1.Columns(4).Value, _
                        ctlGrillaArray1.Columns(5).Value
        frm_VTA_DatoAdicional.Show vbModal
        pxdbDatos(ctlGrillaArray1.Bookmark, 5) = frm_VTA_DatoAdicional.output_valor
        Else
            MsgBox "Esta campo no es editable", vbCritical, App.ProductName
            ctlGrillaArray1.SetFocus
        End If
    ctlGrillaArray1.Array1 = pxdbDatos
    ctlGrillaArray1.Rebind
    End If
'    Debug.Print ctlGrillaArray1.Bookmark
    


Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Formato_Grilla()
    On Error GoTo Handle
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant

        arrCampos = Array("", "", "", "", "", "", "")
        arrCaption = Array("Código", "Descripción", "T", "Max", "Min", "Valor", "Ed.")
        arrAncho = Array(1000, 2500, 200, 600, 800, 1200, "800")
        arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgRight, dbgRight)
        ctlGrillaArray1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        ctlGrillaArray1.Columns(6).Visible = False
        arrCampos = Array("Codigo", "Descripcion", "FLG_RETENCION")
        arrCaption = Array("Código", "Descripción", "Retener")
        arrAncho = Array(700, 3000, 900)
        arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft)
    
        grdDocumentos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        Dim item As New TrueDBGrid70.ValueItem
        With grdDocumentos.Columns("FLG_RETENCION").ValueItems
            item.Value = "1"
            item.DisplayValue = "Retener"
            .Add item
            item.Value = "0"
            item.DisplayValue = "Verificar"
            .Add item
            .Translate = True
        End With
        'grdDocumentos.Columns("FLG_RETENCION").NumberFormat = "Yes/No"
    
    Exit Sub
Handle:
        MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub ctlGrillaArray1_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 0 Then ctlGrillaArray1_DblClick
    End Select
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo Handle
    Unload Me
    'objVenta.CancelarVenta
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
        setteaFormulario Me
        Formato_Grilla
        Inicializa_pantalla
        cboTipoCopago.ListIndex = 0
        txtCopago.Text = 0
        If Not objVenta.CodigoConvenio = "" Then
            Set pxdbDatos = objVenta.DatosAdicional
            Carga_Convenio objVenta.CodigoConvenio, False
            Carga_Convenio_Objeto objVenta.CodigoConvenio
        End If
        'txtConvenio.Text = objVenta.CodigoConvenio
        lblBeneficiario.BackStyle = 1
        lblConvenio.BackStyle = 1
        
        
    Exit Sub
Handle:
        MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Function cargaLista()
    lstDestino.Visible = True
End Function
Public Function fncPalote(strTexto$, Optional intFlag% = 0, Optional strCar$ = "|") As String
Dim intA
    strTexto = CStr(strTexto)
    intA = InStr(strTexto, strCar)
    If intA = 0 Then fncPalote = Trim(strTexto): Exit Function
    If intFlag = 0 Then fncPalote = Trim(left(strTexto, intA - 1))
    If intFlag = 1 Then fncPalote = Trim(right(strTexto, Len(strTexto) - intA))
End Function


Function Repartidor(Activo As Boolean)
    dbcRepartidor.Visible = Activo
    lblTitle(11).Visible = Activo
    
End Function

Function Medico(Activo As Boolean)
    dbcMedico.Visible = Activo
    Label2.Visible = Activo
End Function

Function Diagnostico(Activo As Boolean)
    lstDestino.Visible = Activo
    cmdBuscarDiagnostico.Visible = Activo
End Function

Function Receta(Activo As Boolean)
    Label6.Visible = Activo
    cboOrigen.Visible = Activo
    txtFechaReceta.Visible = Activo
    Label5.Visible = Activo
End Function


Sub ValidaConvenioBTL()

    Dim objConvenio As New clsConvenio
    objConvenio.ValidaConvenioBTL txtCodigoConvenio.Text, objUsuario.Ruc, txtCodigoBeneficiario, objUsuario.Codigo
    Set objConvenio = Nothing
    Screen.MousePointer = vbDefault
End Sub

