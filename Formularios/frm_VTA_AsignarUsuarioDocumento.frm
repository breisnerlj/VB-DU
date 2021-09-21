VERSION 5.00
Begin VB.Form frm_VTA_AsignarUsuarioDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   615
      Left            =   3720
      Picture         =   "frm_VTA_AsignarUsuarioDocumento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "&Asignar"
      Height          =   615
      Left            =   2280
      Picture         =   "frm_VTA_AsignarUsuarioDocumento.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraAsignar 
      Caption         =   "Ingresar Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.OptionButton optDoc 
         Caption         =   "Todos"
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
         Index           =   0
         Left            =   2640
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optDoc 
         Caption         =   "Documento Actual"
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
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   2055
      End
      Begin vbp_Ventas.ctlDataCombo dbcUsuario 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label1 
         Caption         =   $"frm_VTA_AsignarUsuarioDocumento.frx":0B14
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código Vendedor"
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
         Top             =   360
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frm_VTA_AsignarUsuarioDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDoc As oraDynaset
Dim objDocumento As New clsDocumento
Dim strNumDocumento As String
Dim strCodUsuario As String
Dim strTipoDoc As String

Private Sub CmdAsignar_Click()



On Error GoTo Control

        If dbcUsuario.BoundText = "" Then
            MsgBox "Seleccionar el codigo de Usuario", vbCritical, App.ProductName
            Exit Sub
        End If

        
        If optDoc(0).Value Then
                        
                Do While Not rsDoc.EOF
                Dim gvarError As String
                        gvarError = objDocumento.AsignaUsuario(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strTipoDoc, rsDoc(0).Value, dbcUsuario.BoundText, objUsuario.Codigo)
                            If gvarError <> "" Then
                                MsgBox gvarError, vbCritical, App.ProductName
                                Exit Do
                            End If
                        rsDoc.MoveNext
                Loop
                
                
        
        Else
               gvarError = objDocumento.AsignaUsuario(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strTipoDoc, strNumDocumento, dbcUsuario.BoundText, objUsuario.Codigo)
                
                If gvarError = "" Then
                    MsgBox "Se asigno el nuevo usuario al documento", vbInformation, App.ProductName
                    cmdSalir_Click
                Else
                    MsgBox gvarError, vbCritical, App.ProductName
                End If
        End If
        


   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set dbcUsuario.RowSource = objUsuario.Lista("", objUsuario.CodigoLocal)
    dbcUsuario.ListField = "NOM_USUARIO"
    dbcUsuario.BoundColumn = "COD_USUARIO"
    
    dbcUsuario.BoundText = strCodUsuario

End Sub

Public Sub datos(ByVal Caption As String, _
                    ByVal recDoc As oraDynaset, _
                    ByVal NumDocAct As String, _
                    ByVal CodUsuario As String, _
                    ByVal TipoDoc As String)


    Set rsDoc = recDoc
    strNumDocumento = NumDocAct
    strCodUsuario = CodUsuario
    strTipoDoc = TipoDoc

    Me.Caption = Caption
    
    Me.Show vbModal

End Sub

Private Sub optDoc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdAsignar.SetFocus
End Sub
