VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGrabaMedico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Medicos"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmGrabaMedico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1058
      ModoBotones     =   6
      EnabledEfecto   =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3255
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   6135
      Begin vbp_Ventas.ctlTextBox txtDireccion 
         Height          =   345
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
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
      Begin vbp_Ventas.ctlTextBox txtNumCMP 
         Height          =   345
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
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
      Begin vbp_Ventas.ctlTextBox txtApellido 
         Height          =   345
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
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
      Begin MSComCtl2.DTPicker DTPFechaNac 
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   17104897
         CurrentDate     =   39195
      End
      Begin VB.ComboBox cboSexo 
         Height          =   315
         ItemData        =   "frmGrabaMedico.frx":000C
         Left            =   1080
         List            =   "frmGrabaMedico.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1215
      End
      Begin vbp_Ventas.ctlTextBox txtNombre 
         Height          =   345
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
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
      Begin vbp_Ventas.ctlTextBox txtCodigo 
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ColorDefault    =   -2147483633
         ColorDefault    =   -2147483633
         Enabled         =   0   'False
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nac. :"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nº C.M.P. :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Sexo :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   1860
         Visible         =   0   'False
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmGrabaMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim objMedico As New clsMedico
Dim strCodigo As String
Dim strNombre As String
Dim strApellido As String
Dim strDireccion As String
Dim strCMP As String
Dim strSexo As String
Dim strFecha As String
Dim strActivo As String
Dim strModo As String

Private Sub Form_Load()
    txtCodigo.Text = strCodigo
    txtNombre.Text = strNombre
    txtApellido.Text = strApellido
    txtDireccion.Text = strDireccion
    txtNumCMP.Text = strCMP
    cboSexo.ListIndex = IIf(strSexo = "F", 0, IIf(strSexo = "M", 1, -1))
    chkActivo.Value = IIf(strActivo = "ACTIVO", "1", "0")
    DTPFechaNac.Value = strFecha
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Dim a As String
    Dim e As String

    On Error GoTo CtrlErr
    If strModo = "-(Editar)" Then
        e = "Se ha editado al Médico " & Chr(13) & "Codigo : " & txtCodigo.Text & " - " & txtNombre.Text & " " & txtApellido.Text
    Else
        e = "Se ha creado al Médico " & Chr(13) & txtNombre.Text & " " & txtApellido.Text
    End If

    Select Case Index
        Case 1
                                                
            If txtNombre.Text = "" Then MsgBox "Ingrese Nombre de Medico", vbCritical, Caption: txtNombre.selection: Exit Sub
            If txtApellido.Text = "" Then MsgBox "Ingrese Apellidos de Medico", vbCritical, Caption: txtApellido.selection: Exit Sub
            If txtNumCMP.Text = "" Then MsgBox "Ingrese Nº CMP de Medico", vbCritical, Caption: txtNumCMP.selection: Exit Sub
            If cboSexo.Text = "" Then MsgBox "Seleccione el Sexo de Medico", vbCritical, Caption: Exit Sub
            
            
            a = objMedico.Graba(txtCodigo.Text, _
                                txtNombre.Text, _
                                txtApellido.Text, _
                                txtDireccion.Text, _
                                txtNumCMP.Text, _
                                IIf(left(cboSexo.Text, 1) = "1", "M", "F"), _
                                CStr(Format(DTPFechaNac.Value, "DD/MM/YYYY")), _
                                chkActivo.Value, _
                                objUsuario.Codigo)
                    
            If a = "" Then
                MsgBox e, vbInformation + vbOKOnly, "Grabar"
                ''''Set frmMedico.grdMedico.DataSource = objMedico.Lista("")
                Unload Me
            Else
                MsgBox a, vbCritical + vbOKOnly, "Atención"
            End If
            
        Case 2
            If txtCodigo.Text <> "" Then
                txtCodigo.Text = strCodigo
                txtNombre.Text = strNombre
                txtApellido.Text = strApellido
                txtNumCMP.Text = strCMP
                txtDireccion.Text = strDireccion
                cboSexo.ListIndex = IIf(strSexo = "F", 0, IIf(strSexo = "M", 1, -1))
                DTPFechaNac.Value = strFecha
                chkActivo.Value = strActivo
            Else
                Unload Me
            End If
        Case 3
            Unload Me
    
    End Select
    
    Exit Sub
CtrlErr:
        MsgBox Err.Description, vbCritical + vbOKOnly, Err.Number
    
    
End Sub

Public Sub Datos(ByVal Codigo As String, _
                    ByVal Nombre As String, _
                    ByVal Apellido As String, _
                    ByVal Direccion As String, _
                    ByVal CMP As String, _
                    ByVal Sexo As String, _
                    ByVal Activo As String, _
                    ByVal Fecha As String, _
                    ByVal Caption As String)

        strCodigo = Codigo
        strNombre = Nombre
        strApellido = Apellido
        strDireccion = Direccion
        strCMP = CMP
        strSexo = Sexo
        strActivo = Activo
        strFecha = Fecha
        strModo = Caption
    
    Me.Caption = Me.Caption & Caption
    Me.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            Call ctlToolBar1_Click(Grabar, 1)
        Case vbKeyEscape
            Unload Me
    End Select
End Sub
