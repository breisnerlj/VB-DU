VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ADM_MantMaquinasxLocal 
   Caption         =   "Mantenimiento de Maquinas Por Local"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMaquinasxLocal 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      Begin vbp_Ventas.ctlDataCombo CboLocales 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblMaquina 
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblLocal 
         Caption         =   "Local:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblMaq 
         Caption         =   "Maquina:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tooMantenimiento 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   1111
      ButtonWidth     =   953
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Graba"
            Key             =   "Save"
            Object.ToolTipText     =   "Agrega Documentos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Exit"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4320
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_ADM_MantMaquinasxLocal.frx":0000
               Key             =   "Machine"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_ADM_MantMaquinasxLocal.frx":059A
               Key             =   "File"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_ADM_MantMaquinasxLocal.frx":0B34
               Key             =   "mail"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_ADM_MantMaquinasxLocal.frx":10CE
               Key             =   "Exit"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_ADM_MantMaquinasxLocal.frx":1668
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Command1"
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   2640
         Width           =   255
      End
   End
End
Attribute VB_Name = "frm_ADM_MantMaquinasxLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim objLocal As New clsLocal

Private Sub Form_Load()
lblMaquina.Caption = objUsuario.NombrePC
CargarLocales
End Sub

Private Sub CargarLocales()
  Set CboLocales.RowSource = objLocal.Lista(objUsuario.CodigoEmpresa, "", "")
                                                         
    CboLocales.ListField = "LOCAL_DEX"
    CboLocales.BoundColumn = "COD_LOCAL"
    
End Sub

Private Sub tooMantenimiento_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Save"
            Graba
        Case "Exit"
            Unload Me
    End Select
End Sub


Private Sub Graba()
    Dim strMensaje As String
    If CboLocales.BoundText <> "" Then
        strMensaje = objLocal.ActualizaLocalxMaquina(CboLocales.BoundText, objUsuario.NombrePC, objUsuario.Codigo)
        If strMensaje = "" Then
           MsgBox "Se grabo satisfactoriamente, Se cerrará el Sistema...", vbInformation, App.ProductName
           End
        Else
            MsgBox strMensaje, vbCritical, App.ProductName
        End If
    Else
         MsgBox "Falta Elegir el Local", vbOKOnly + vbExclamation, "Validación"
         CboLocales.SetFocus
    End If
    
End Sub

