VERSION 5.00
Begin VB.Form frmGrabaMotorizado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Motorizados"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frmGrabaMotorizado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese Datos"
      Height          =   3855
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   5655
      Begin vbp_Ventas.ctlDataCombo cboCia 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox txtAlias 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   2880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
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
      Begin vbp_Ventas.ctlTextBox txtNroRef 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   2520
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
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
      Begin vbp_Ventas.ctlTextBox txtApellido 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   2160
         Width           =   4095
         _ExtentX        =   7223
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
      Begin vbp_Ventas.ctlTextBox txtNombre 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
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
      Begin vbp_Ventas.ctlDataCombo dbcLocalAsignado 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1440
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
      End
      Begin vbp_Ventas.ctlDataCombo dbcEstado 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
      End
      Begin vbp_Ventas.ctlTextBox txtCodigo 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   360
         Width           =   675
         _ExtentX        =   1191
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
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label Label8 
         Caption         =   "Cia. Asignada"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Alias"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2940
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Ref."
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2580
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Apellido"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2220
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Local Asignado"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1058
      ModoBotones     =   6
   End
End
Attribute VB_Name = "frmGrabaMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMotorizado As New clsMotorizado
Dim objLocal As New clsLocal
Dim strCodigo As String
Dim strEstado As String
Dim sCiaAsig As String
Dim strLocalAsignado As String
Dim strNombre As String
Dim strApellido As String
Dim strNroRef As String
Dim strAlias As String
Dim intActivo As Integer
Dim strModo As String
Dim sCia As String


Private Sub cboCia_Change()
  ' bjct, listar locales segun la cia elegida
  sCia = cboCia.BoundText
  Set dbcLocalAsignado.RowSource = objLocal.Lista(sCia, "", "", "1")
  dbcLocalAsignado.BoundColumn = "COD_LOCAL"
  dbcLocalAsignado.ListField = "LOCAL_DEX2"
  dbcLocalAsignado.BoundText = ""
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
Dim a As String
Dim e As String
On Error GoTo CtrlErr
    If strModo = "-(Editar)" Then
        e = "Se ha editado al motorizado " & Chr(13) & "Codigo : " & txtCodigo.Text & " - " & txtNombre.Text & " " & txtApellido.Text
    Else
        e = "Se ha creado al motorizado " & Chr(13) & txtNombre.Text & " " & txtApellido.Text
    End If
    Select Case Index
        Case 1
            'If Trim(dbcEstado.Text) = "" Then MsgBox "Seleccione el Estado", vbCritical, Caption: dbcEstado.SetFocus: Exit Sub
            If Trim(dbcLocalAsignado.Text) = "" Then MsgBox "Seleccione el Local Asignado", vbCritical, Caption: dbcLocalAsignado.SetFocus: Exit Sub
            If Trim(txtNombre.Text) = "" Then MsgBox "Ingrese el Nombre del Motorizado", vbCritical, Caption: txtNombre.selection: Exit Sub
            If Trim(txtApellido.Text) = "" Then MsgBox "Ingrese el Apellido del Motorizado", vbCritical, Caption: txtApellido.selection: Exit Sub
            If Trim(txtNroRef.Text) = "" Then MsgBox "Ingrese el Nº Referencia", vbCritical, Caption: txtNroRef.selection: Exit Sub
            
            a = objMotorizado.Graba(txtCodigo.Text, _
                            dbcEstado.BoundText, _
                            dbcLocalAsignado.BoundText, _
                            txtNombre.Text, _
                            txtApellido.Text, _
                            txtNroRef.Text, _
                            txtAlias.Text, _
                            chkActivo.Value)
                    
            If a = "" Then
                MsgBox e, vbInformation + vbOKOnly, "Grabar"
                Unload Me
            Else
                MsgBox a, vbCritical + vbOKOnly, "Atención"
            End If
        Case 2
            If txtCodigo.Text <> "" Then
                If Trim(dbcEstado.BoundText) = strEstado Then
                    If Trim(dbcLocalAsignado.BoundText) = strLocalAsignado Then
                        If Trim(txtNombre.Text) = strNombre Then
                            If Trim(txtApellido.Text) = strApellido Then
                                If Trim(txtNroRef.Text) = strNroRef Then
                                    If Trim(txtAlias.Text) = strAlias Then
                                        If Trim(chkActivo.Value) = intActivo Then
                                            MsgBox "No han Habido Cambios", vbInformation, "Aviso"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                txtCodigo.Text = strCodigo
                dbcEstado.BoundText = strEstado
                dbcLocalAsignado.BoundText = strLocalAsignado
                txtNombre.Text = strNombre
                txtApellido.Text = strApellido
                txtNroRef.Text = strNroRef
                txtAlias.Text = strAlias
                chkActivo.Value = intActivo
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

Public Sub datos(ByVal Codigo As String, _
                    ByVal Estado As String, _
                    ByVal LocalAsignado As String, _
                    ByVal Nombre As String, _
                    ByVal Apellido As String, _
                    ByVal NroRef As String, _
                    ByVal Alias As String, _
                    ByVal Activo As Integer, _
                    ByVal Caption As String, _
                    Optional ByVal ciaAsig As String = "")

    strCodigo = Codigo
    strEstado = Estado
    sCiaAsig = ciaAsig
    strLocalAsignado = LocalAsignado
    strNombre = Nombre
    strApellido = Apellido
    strNroRef = NroRef
    strAlias = Alias
    intActivo = Activo
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

Private Sub Form_Load()
    Dim sCodLoc As String
    Dim rsCia As oraDynaset
    
    Set dbcEstado.RowSource = objMotorizado.ListaEstado
    dbcEstado.ListField = "DES_ESTADO_MOTORIZADO"
    dbcEstado.BoundColumn = "COD_ESTADO_MOTORIZADO"
    
    ' bjct,load cia y selecionar segun local
    Set cboCia.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCia.ListField = "Des"
    cboCia.BoundColumn = "Cod"
    
    ' set cia segun local
    sCodLoc = "" & strLocalAsignado
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, sCodLoc)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
      cboCia.BoundText = sCia
      
     Else
      MsgBox "No se puede Asignar la CIA....", vbCritical, App.ProductName
    End If
    Set rsCia = Nothing
    
    ' ejct
    
    
    'Set dbcLocalAsignado.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "")
    Set dbcLocalAsignado.RowSource = objLocal.Lista(sCia, "")
    dbcLocalAsignado.BoundColumn = "COD_LOCAL"
    dbcLocalAsignado.ListField = "LOCAL_DEX2"

    txtCodigo.Text = strCodigo
    dbcEstado.BoundText = strEstado
    dbcLocalAsignado.BoundText = strLocalAsignado
    txtNombre.Text = strNombre
    txtApellido.Text = strApellido
    txtNroRef.Text = strNroRef
    txtAlias.Text = strAlias
    chkActivo.Value = intActivo
End Sub
