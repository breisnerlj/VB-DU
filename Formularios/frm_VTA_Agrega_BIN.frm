VERSION 5.00
Begin VB.Form frm_VTA_Agrega_BIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Bin de Tarjeta"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Nueva Tarjeta"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin vbp_Ventas.ctlDataCombo cboTarjeta 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox txtNuevoBin 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
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
         Locked          =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de Bin"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Tarjeta"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "frm_VTA_Agrega_BIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pbNumeroTarjeta As String
Dim objFormaPago As New clsFormaPago

Private Sub cmdAceptar_Click()
Dim strMensaje As String
If cboTarjeta.BoundText = "*" Or cboTarjeta.BoundText = "" Then MsgBox "Seleccione el tipo de tarjeta", vbCritical, App.ProductName: Exit Sub
strMensaje = objFormaPago.GrabarBin("", cboTarjeta.BoundText, objUsuario.Codigo, Mid(txtNuevoBin.Text, 1, 6))
    If strMensaje = "" Then
        MsgBox "Se grabo satisfactoriamente", vbExclamation, App.ProductName
        pbNumeroTarjeta = txtNuevoBin.Text
        Unload Me
    Else
        MsgBox strMensaje, vbCritical, App.ProductName
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim objTarjetas As New clsFormaPago
Set cboTarjeta.RowSource = objTarjetas.ListaTarjetasVenta
    cboTarjeta.ListField = "DES"
    cboTarjeta.BoundColumn = "COD"
    cboTarjeta.BoundText = "*"
    
txtNuevoBin.Text = pbNumeroTarjeta
pbNumeroTarjeta = ""
End Sub

