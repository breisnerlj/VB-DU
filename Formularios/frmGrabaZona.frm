VERSION 5.00
Begin VB.Form frmGrabaZona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Zonas"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmGrabaZona.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese Datos"
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   4500
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Activo"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   1515
         Width           =   1245
      End
      Begin vbp_Ventas.ctlTextBox txtAbreviatura 
         Height          =   315
         Left            =   1155
         TabIndex        =   1
         Top             =   1110
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         MaxLength       =   5
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
      Begin vbp_Ventas.ctlTextBox txtDescripcion 
         Height          =   315
         Left            =   1155
         TabIndex        =   0
         Top             =   735
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
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
      Begin vbp_Ventas.ctlTextBox txtCodigo 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   785
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Abreviatura"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1150
         Width           =   810
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   1058
      ModoBotones     =   6
   End
End
Attribute VB_Name = "frmGrabaZona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X_acc As String
Dim XCodigo As String
Dim XDescripcion As String
Dim XAbreviatura As String
Dim intActivo As Integer
Dim objZona As New clsZona

Private Sub CmdLocales_Click()
   ' Call frmLocalesPorZona.SetCodZona(Trim(txtCodigo.Text), Trim(txtDescripcion.Text))
   ' frmLocalesPorZona.Caption = frmLocalesPorZona.Caption & "- Zona " & Trim(txtDescripcion.Text)
   ' frmLocalesPorZona.Show 1
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case Index
        Case 1
            Call Grabar
        Case 2
            Call Deshacer
        Case 3
            Unload Me
    End Select
End Sub

Private Sub Deshacer()
    If X_acc = "Editar" Then
        If Trim(txtDescripcion.Text) = XDescripcion Then
            If Trim(txtAbreviatura.Text) = XAbreviatura Then
                If Trim(chkActivo.Value) = intActivo Then
                    MsgBox "No han habido cambios", vbInformation, "Aviso"
                End If
            End If
        End If
        txtDescripcion.Text = XDescripcion
        txtAbreviatura.Text = XAbreviatura
        chkActivo.Value = intActivo
    End If
End Sub

Private Sub Grabar()
On Error GoTo registra
Dim MS, a As String
    If Trim(txtDescripcion.Text) = "" Then MsgBox "Ingrese el Nombre de la Zona", vbInformation, "Mensaje": txtDescripcion.SetFocus: Exit Sub
    If X_acc = "Nuevo" Then
        MS = "Registro Satisfactorio de la Zona : " & Trim(txtDescripcion.Text)
    ElseIf X_acc = "Editar" Then
        MS = "Edición Satisfactoria de la Zona : " & Trim(txtDescripcion.Text)
    End If
    a = objZona.Graba(txtCodigo.Text, _
                      txtDescripcion.Text, _
                      txtAbreviatura.Text, _
                      chkActivo.Value, objUsuario.Codigo)
                    
    If a = "" Then
        MsgBox MS, vbInformation + vbOKOnly, "Grabar"
        Unload Me
    Else
        MsgBox a, vbCritical + vbOKOnly, "Atención"
    End If
Exit Sub
registra:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Public Sub SetearDatos(ByVal Codigo As String, _
                       ByVal Descripcion As String, _
                       ByVal Abreviatura As String, _
                       ByVal Activo As Integer, _
                       ByVal Caption As String)
    XCodigo = Codigo
    XDescripcion = Descripcion
    XAbreviatura = Abreviatura
    intActivo = Activo
    X_acc = Caption
    Me.Caption = Me.Caption & "-" & Caption
End Sub

Private Sub Form_Load()
    txtCodigo.Text = XCodigo
    txtDescripcion.Text = XDescripcion
    txtAbreviatura.Text = XAbreviatura
    chkActivo.Value = intActivo
End Sub
