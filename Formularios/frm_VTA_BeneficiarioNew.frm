VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_VTA_BeneficiarioNew 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlDataCombo cboConvenio 
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   5280
      Picture         =   "frm_VTA_BeneficiarioNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "&Solicitar"
      Height          =   615
      Left            =   3960
      Picture         =   "frm_VTA_BeneficiarioNew.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Beneficiarío"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   6495
      Begin vbp_Ventas.ctlDataCombo cboZonal 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   3840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboEstadoCivil 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   2595
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox TxtCodRef 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Tipo            =   2
         MaxLength       =   20
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
      Begin vbp_Ventas.ctlTextBox TxtNombre 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   900
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Tipo            =   2
         MaxLength       =   100
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
      Begin vbp_Ventas.ctlTextBox TxtApellidoPat 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Tipo            =   2
         MaxLength       =   100
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
      Begin vbp_Ventas.ctlTextBox TxtApellidoMat 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1730
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Tipo            =   2
         MaxLength       =   100
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
      Begin vbp_Ventas.ctlTextBox TxtCargo 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2145
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Tipo            =   2
         MaxLength       =   50
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
      Begin vbp_Ventas.ctlTextBox TxtEmail 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   3405
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Tipo            =   2
         MaxLength       =   50
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
      Begin MSComCtl2.DTPicker dtpFchNac 
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   61472769
         CurrentDate     =   39084
      End
      Begin vbp_Ventas.ctlTextBox TxtDocumento 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Tipo            =   3
         MaxLength       =   8
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
      Begin VB.Label Label4 
         Caption         =   "DNI"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3060
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Nombres"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Cargo"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2210
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Estado Civil"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2620
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha Nac."
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   3060
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Ape. Paterno"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Ape. Materno"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1790
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3460
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Zonal"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3870
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio"
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
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Solicitud de Nuevo Beneficiario"
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
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frm_VTA_BeneficiarioNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objConvenio As New clsConvenio

Private Sub cmdAsignar_Click()

    Dim CtrlErr As String
        
    '*** Validaciones ***'
    If CboConvenio.BoundText = "" Then MsgBox "Seleccione un convenio", vbCritical, App.ProductName: CboConvenio.SetFocus: Exit Sub
    If TxtCodRef.Text = "" Then MsgBox "Ingrese Codigo de Referencia", vbCritical, App.ProductName: TxtCodRef.selection: Exit Sub
    If TxtDocumento.Text = "" Then MsgBox "Ingrese Documento de Indentidad", vbCritical, App.ProductName: TxtDocumento.selection: Exit Sub
    If TxtNombre.Text = "" Then MsgBox "Ingrese Nombre del Beneficiario", vbCritical, App.ProductName: TxtNombre.selection: Exit Sub
    If TxtApellidoPat.Text = "" Then MsgBox "Ingrese Apellido Paterno del Beneficiario", vbCritical, App.ProductName: TxtApellidoPat.selection: Exit Sub
    If TxtApellidoMat.Text = "" Then MsgBox "Ingrese Apellido Materno del Beneficiario", vbCritical, App.ProductName: TxtApellidoMat.selection: Exit Sub
        
    TxtCodRef.Text = Trim(TxtCodRef.Text)
    
    
    
    
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Autor : Arturo Escate
    'Fecha : 10/11/2009
    'Proposito: Esto es para validar si necesita autorizacion previa
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim ObjValidacion As New clsAprobacion
    Dim strNumeroSolicitud As String
    Dim strAccion As String
    Dim strMensaje As String
    Dim strCodigoAutorizacion As String
    Dim srtCodigoAUTH As String
    Dim strStore As String
    srtCodigoAUTH = ""
valida:
'TxtImpTot.text

    strStore = ObjValidacion.Solicita("2", strAccion, strMensaje, srtCodigoAUTH, objUsuario.CodigoLocal, objUsuario.CodigoLiquidacion, "", "", "", "", "", objConvenio.LineaCreditoBase(CboConvenio.BoundText), "", "", objUsuario.Codigo, "", strCodigoAutorizacion, "", "", TxtCodRef.Text, TxtApellidoPat.Text, TxtApellidoMat.Text, TxtNombre.Text, TxtCargo.Text, TxtDocumento.Text, ctlCboEstadoCivil.BoundText, TxtEmail.Text, "", "", "", "", "", "", "", "", "", "", CboConvenio.BoundText, "", "", "", "", "", "", "")
    If Not strStore = "" Then
        MsgBox strStore, vbCritical, App.ProductName
        Exit Sub
    Else
        Select Case strAccion
            Case 0
                    MsgBox strMensaje, vbInformation, App.ProductName
            Case 1
                   MsgBox strMensaje, vbCritical, App.ProductName
                   Exit Sub
            Case 2
                   MsgBox strMensaje, vbInformation, App.ProductName
                   Exit Sub
            Case 3
                If MsgBox(strMensaje & Chr(13) & "¿Desea ingresar el codigo de autorización?", vbYesNo + vbInformation, App.ProductName) = vbYes Then
                    srtCodigoAUTH = frmAprobacion.Carga
                    If Not srtCodigoAUTH = "" Then
                        GoTo valida
                        Exit Sub
                    End If
                   Exit Sub
                Else
                    Exit Sub
                End If
            Case Else
                   MsgBox "no esta implementado", vbInformation, App.ProductName
                   Exit Sub
        End Select
    End If

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    
    
    
    
    
    CtrlErr = objConvenio.GrabaBenefiarioCnv(objUsuario.CodigoEmpresa, _
                                             CboConvenio.BoundText, _
                                             "", _
                                             Trim(TxtDocumento.Text), _
                                             Trim(TxtApellidoPat.Text), _
                                             Trim(TxtApellidoMat.Text), _
                                             Trim(TxtNombre.Text), _
                                             Trim(TxtCargo.Text), _
                                             objConvenio.LineaCreditoBase(CboConvenio.BoundText), _
                                              "0", _
                                             "0", _
                                             "", _
                                             "", _
                                             "", _
                                             Trim(TxtCodRef.Text), _
                                             "1", _
                                             objUsuario.Codigo, _
                                             Trim(ctlCboEstadoCivil.BoundText), _
                                             CStr(Format(dtpFchNac.Value, "dd/mm/yyyy")), _
                                             Trim(TxtEmail.Text), _
                                             "S", _
                                             "1", cboZonal.BoundText)
    
    If CtrlErr = "" Then
        MsgBox "Se Grabó el Beneficiario Satisfactoriamente", vbInformation, App.ProductName
       ' frmMantBeneficiario.GrdBenificario.Limpiar
        Unload Me
    Else
        MsgBox CtrlErr, vbCritical, App.ProductName
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    setteaFormulario Me
    Me.left = 0
    Me.top = 0
    Set ctlCboEstadoCivil.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_ESTADO_CIVIL.FN_LISTA", 0, "")
    ctlCboEstadoCivil.ListField = "DES_ESTADO_CIVIL"
    ctlCboEstadoCivil.BoundColumn = "COD_ESTADO_CIVIL"
    Set cboZonal.RowSource = objConvenio.ListaZonal("")
    cboZonal.ListField = "DES_ZONAL"
    cboZonal.BoundColumn = "COD_ZONAL"
    With CboConvenio
        Set .RowSource = objConvenio.ListaAddBeneficiario
        .BoundColumn = "COD_CONVENIO"
        .ListField = "DES_CONVENIO"
    End With

End Sub

