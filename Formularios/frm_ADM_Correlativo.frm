VERSION 5.00
Begin VB.Form frm_ADM_Correlativo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correlativo"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frm_ADM_Correlativo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   3960
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Enviar la impresión  a la Máquina Alternativa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   4335
      End
      Begin vbp_Ventas.ctlTextBox TxtNumero 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Tipo            =   3
         MaxLength       =   10
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
      Begin vbp_Ventas.ctlDataCombo ctlCboMaquina 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboTicketera 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   1160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlcboFormato 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "T. &Formato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1605
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "&Ticketera:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1210
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "&M. Alternativa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° &Documento:"
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
         TabIndex        =   0
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   1080
      Width           =   135
   End
End
Attribute VB_Name = "frm_ADM_Correlativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public input_CodigoMaquina As String
Public input_TipoDocumento As String
Public input_CodigoMaquinaRel As String
Public input_CodigoFlagEncola As String
Public input_NumeroDocumento As String
Public input_Ticketera As String
Public input_CodFormato As String

Dim objMaquina As New clsMaquina
Dim intEncontro As Integer

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    TxtNumero.Tipo = Entero
    'If TxtNumero.Text = "" Then MsgBox "Ingrese el Numero Correlativo", vbCritical, App.ProductName: TxtNumero.SetFocus: Exit Sub
    
    'On Error GoTo CtrlErr
    intEncontro = objMaquina.fn_Existe_Rel_Tkt_Maq(objUsuario.CodigoEmpresa, _
                                                   objUsuario.CodigoLocal, _
                                                   input_TipoDocumento, _
                                                   Replace(ctlCboTicketera.BoundText, "*", ""))
    
    If intEncontro > 0 Then
        MsgBox "La Ticketera ya está asociada a una máquina", vbCritical, Caption: ctlCboTicketera.SetFocus: Exit Sub
    End If
    
    frm_ADM_Maquinas.grdMaquinas.Array1 = objVenta.AgregaMaquina(input_TipoDocumento, _
                                                                 input_CodigoMaquina, _
                                                                 TxtNumero.Text, _
                                                                 Check1.Value, _
                                                                 Replace(ctlCboMaquina.BoundText, "*", ""), _
                                                                 Replace(ctlCboTicketera.BoundText, "*", ""), _
                                                                 Replace(ctlcboFormato.BoundText, "*", "") _
                                                                 )
    frm_ADM_Maquinas.grdMaquinas.Refresh
    Unload Me
End Sub

Private Sub ctlCboMaquina_Click(Area As Integer)
Dim strCodFormato As String
On Error GoTo CtrlErr
    If ctlCboMaquina.BoundText <> "*" Then
        strCodFormato = objMaquina.CodFormato(objUsuario.CodigoEmpresa, input_TipoDocumento, ctlCboMaquina.BoundText)
        strCodFormato = IIf(strCodFormato = "", "*", strCodFormato)
    Else
        Select Case input_TipoDocumento
            Case "FAC"
                strCodFormato = "002"
            Case "BOL"
                strCodFormato = "004"
            Case "GRL"
                strCodFormato = "005"
            Case Else
                strCodFormato = "*"
        End Select
    End If
    
    ctlcboFormato.BoundText = strCodFormato
    
    ctlcboFormato.Enabled = (ctlCboMaquina.BoundText = "*")
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Load()
    Dim objMaquina As New clsMaquina
    TxtNumero.Text = input_NumeroDocumento
    
    Set ctlCboMaquina.RowSource = objMaquina.MaquinaLocal(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, "1")
    ctlCboMaquina.ListField = "DES"
    ctlCboMaquina.BoundColumn = "COD"
    ctlCboMaquina.BoundText = IIf(input_CodigoMaquinaRel = "", "*", input_CodigoMaquinaRel)
    
    Set ctlCboTicketera.RowSource = objMaquina.ListaTicketeraLocal(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
    ctlCboTicketera.ListField = "DES"
    ctlCboTicketera.BoundColumn = "COD"
    ctlCboTicketera.BoundText = IIf(input_Ticketera = "", "*", input_Ticketera)
    
    Set objMaquina = Nothing
    Check1.Value = IIf(input_CodigoFlagEncola = "", 0, input_CodigoFlagEncola)
    
    '03/06/2008 jcrazuri
    Dim objFormato As New clsFormato
    Set ctlcboFormato.RowSource = objFormato.ListaxDoc(input_TipoDocumento, "NINGUNO")
    ctlcboFormato.ListField = "DES"
    ctlcboFormato.BoundColumn = "COD"
    ctlcboFormato.BoundText = IIf(input_CodFormato = "", "*", input_CodFormato)
    Set objFormato = Nothing
    
    If input_TipoDocumento = "TKB" Or input_TipoDocumento = "TKF" Then
        TxtNumero.Text = ""
        TxtNumero.Bloqueado = True
        ctlCboMaquina.Enabled = False
        ctlcboFormato.Enabled = False
        Check1.Enabled = False
        ctlCboTicketera.Enabled = True
    Else
        ctlCboTicketera.Enabled = False
        ctlcboFormato.Enabled = True
        TxtNumero.Bloqueado = False
        ctlCboMaquina.Enabled = True
        Check1.Enabled = True
    End If
    
End Sub

