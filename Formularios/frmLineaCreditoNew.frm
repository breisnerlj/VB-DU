VERSION 5.00
Begin VB.Form frmLineaCreditoNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de Ampliación de Linea de Credito"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtCodBeneficiario 
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtCodConvenio 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      Begin vbp_Ventas.ctlTextBox txtLineaActual 
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   990
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         ColorDefault    =   -2147483639
         ColorDefault    =   -2147483639
         Enabled         =   0   'False
         TABAuto         =   0   'False
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtLineaNew 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   1470
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         AutoSize        =   -1  'True
         Caption         =   "Linea Crédito Nueva"
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
         TabIndex        =   11
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Linea Crédito Actual"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Label lblBeneficiario 
         Caption         =   "lblBeneficiario"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   5280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Beneficiario"
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
         TabIndex        =   6
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         TabIndex        =   5
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblConvenio 
         Caption         =   "lblConvenio"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   5280
      End
   End
End
Attribute VB_Name = "frmLineaCreditoNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objBeneficiario As New clsBeneficiario
Dim rs As oraDynaset
Dim objConvenio As New clsConvenio

Public Function MuestraBeneficiario(ByVal CodigoConvenio As String, ByVal CodigoCliente As String)
    
    Set rs = objBeneficiario.Lista(CodigoConvenio, CodigoCliente)
    lblBeneficiario.Caption = "" & rs("DES_CLIENTE").Value
    lblConvenio.Caption = "" & rs("DES_CONVENIO").Value
    txtCodBeneficiario.Text = "" & rs("COD_CLIENTE").Value
    txtCodConvenio.Text = "" & rs("COD_CONVENIO").Value
    txtLineaActual.Text = "" & rs("IMP_LINEA_CREDITO_ORI").Value
        Me.Show vbModal
    
End Function


Private Sub cmdGrabar_Click()
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

    strStore = ObjValidacion.Solicita("9", strAccion, strMensaje, srtCodigoAUTH, objUsuario.CodigoLocal, objUsuario.CodigoLiquidacion, "", "", "", "", "", Val(txtLineaNew.Text), "", "", objUsuario.Codigo, "", strCodigoAutorizacion, "", "", _
        "" & rs("COD_REFERENCIA"), _
        "" & rs("DES_APE_CLIENTE"), _
        "" & rs("DES_APE2_CLIENTE"), _
        "" & rs("DES_NOM_CLIENTE"), _
            "" & rs("DES_CARGO"), _
            "" & rs("NUM_DOCUMENTO_ID"), _
            "" & rs("COD_ESTADO_CIVIL"), _
            "" & rs("DES_EMAIL"), _
            "", "", "", "", "", "", "", "", "", "", txtCodConvenio.Text, "", "", "", "", "", "", "")
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
    
    
    Dim CtrlErr As String
    
    
    
    CtrlErr = objConvenio.GrabaBenefiarioCnv(objUsuario.CodigoEmpresa, _
                                             txtCodConvenio.Text, _
                                             txtCodBeneficiario.Text, _
                                             Trim("" & rs("NUM_DOCUMENTO_ID")), _
                                             Trim("" & rs("DES_APE_CLIENTE")), _
                                             Trim("" & rs("DES_APE2_CLIENTE")), _
                                             Trim("" & rs("DES_NOM_CLIENTE")), _
                                             Trim("" & rs("DES_CARGO")), _
                                             Val(txtLineaActual.Text), _
                                              "1", _
                                             Val(txtLineaNew.Text), _
                                             objUsuario.sysdate, _
                                             objUsuario.sysdate, _
                                             "", _
                                             Trim("" & rs("COD_REFERENCIA")), _
                                             "1", _
                                             objUsuario.Codigo, _
                                             Trim("" & rs("COD_ESTADO_CIVIL")), _
                                             CStr(Format(rs("FCH_NACIMIENTO"), "dd/mm/yyyy")), _
                                             Trim("" & rs("DES_EMAIL")), _
                                             "S", _
                                             "0", "" & rs("COD_ZONAL"))
    
    If CtrlErr = "" Then
        MsgBox "Se Grabó el Beneficiario Satisfactoriamente", vbInformation, Caption
       ' frmMantBeneficiario.GrdBenificario.Limpiar
        Unload Me
    Else
        MsgBox CtrlErr, vbCritical, Caption
    End If
    Unload Me
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
