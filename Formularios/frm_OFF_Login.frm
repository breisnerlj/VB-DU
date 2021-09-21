VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_OFF_Logeo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   Icon            =   "frm_OFF_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3810
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "&Ingresar"
      Height          =   350
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2190
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox txtUsuario 
      Height          =   375
      Left            =   1620
      TabIndex        =   0
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CONTINGENCIA DE VENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   340
      Width           =   3615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   885
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su usuario (Código de Planilla)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   1980
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BIENVENIDO AL SISTEMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   50
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   60
      X2              =   3660
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frm_OFF_Logeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdIngresar_Click()
Dim objDocumento As cls_OFF_Documento
Dim objIni As cls_ArchivoIni
Dim tmpDocumento As New XArrayDB
Dim i As Integer
Dim strCorrelativo As String
Dim strDocumento As String


    On Error GoTo handle
    If txtUsuario.Text = "" Then Exit Sub
    If objOFFUsuario.Login(txtUsuario.Text) = True Then
        'Bienvenida
        MsgBox "BIENVENIDO(A) " & Chr(13) & _
                    "    Usuario: " & objOFFUsuario.NombreUsuario & Chr(13) & _
                    "    Local  : " & objOFFUsuario.DireccionLocal & " BTL " & objOFFUsuario.CodLocal & Chr(13) & _
                    "    Máquina: " & objOFFUsuario.CodMaquina & Chr(13), vbInformation, App.ProductName
                    
                    
        Set objIni = New cls_ArchivoIni
        objIni.GuardarIni gstrIni, "general", "FLG_CONTINGENCIA", "1"
        Set objIni = Nothing
        
        Unload Me
        
        If objOFFUsuario.Contingencia = False Then
            frm_OFF_Documentos.Show
        Else

            Set objIni = New cls_ArchivoIni
            Set objDocumento = New cls_OFF_Documento
            Set tmpDocumento = objDocumento.ListaTipoDocumento
            For i = tmpDocumento.LowerBound(1) To tmpDocumento.UpperBound(1)
                strDocumento = Trim(tmpDocumento(i, 0))
                strCorrelativo = Trim(tmpDocumento(i, 2))
                If Len(strDocumento) = 3 Then
                    objDocumento.ActualizaCorrelativo strDocumento, strCorrelativo
                End If
            Next i
            objOFFVenta.CodDocDefault = objIni.LeerIni(gstrIni, "general", "COD_DOC_DEFAULT", "")
            objOFFVenta.NumDocDefault = objDocumento.UltimoCorrelativo(objOFFVenta.CodDocDefault)
            objIni.GuardarIni gstrIni, "general", "COD_DOC_DEFAULT", objOFFVenta.CodDocDefault
            Set objIni = Nothing
            Set objDocumento = Nothing
        
            frm_OFF_Principal.Show
        End If
    Else
        MsgBox "El usuario no existe o no está asignado a este Local", vbCritical, App.ProductName
        txtUsuario.SetFocus
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub txtUsuario_GotFocus()
    stbPrincipal.Panels(1).Text = Label4.Caption
End Sub
