VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_VTA_Ctrl_Depositos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ilsPrincipal 
      Left            =   6120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   57
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":0000
            Key             =   "Revisar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":059A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":0B34
            Key             =   "AddDoc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":10CE
            Key             =   "AddUser"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":1668
            Key             =   "Admin"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":1C02
            Key             =   "Avance"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":219C
            Key             =   "Attach"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":2736
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":2CD0
            Key             =   "Book"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":326A
            Key             =   "Calc"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":3804
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":3D9E
            Key             =   "Chat"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":4338
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":48D2
            Key             =   "Clientes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":4E6C
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":5406
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":59A0
            Key             =   "Config"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":5F3A
            Key             =   "Contac"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":64D4
            Key             =   "Control"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":6A6E
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":7008
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":75A2
            Key             =   "Date"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":7B3C
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":80D6
            Key             =   "Document"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":8670
            Key             =   "DownLevel"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":8C0A
            Key             =   "DownLoad"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":91A4
            Key             =   "Draw"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":973E
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":9CD8
            Key             =   ""
            Object.Tag             =   "Estatic"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":A272
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":A80C
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":ADA6
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":B340
            Key             =   "Games"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":B8DA
            Key             =   "Group"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":BE74
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":C40E
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":C9A8
            Key             =   "Idea"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":CF42
            Key             =   "Info"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":D4DC
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":DA76
            Key             =   "Level"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":E010
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":E5AA
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":EB44
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":F0DE
            Key             =   "Notes"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":F678
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":FC12
            Key             =   "Paint"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":101AC
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":10746
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":10CE0
            Key             =   "Phone"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":1127A
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":11814
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":11DAE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":12348
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":128E2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":12E7C
            Key             =   "UnLoock"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":13416
            Key             =   "Tools"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Ctrl_Depositos.frx":139B0
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1111
      ButtonWidth     =   1429
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ilsPrincipal"
      DisabledImageList=   "ilsPrincipal"
      HotImageList    =   "ilsPrincipal"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Registro"
            Key             =   "Copy"
            ImageIndex      =   20
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Copy"
                  Text            =   "Remesa"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Document"
                  Text            =   "Remito"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Deposito"
                  Text            =   "Deposito"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Re - Imprimir"
            Key             =   "Printer"
            ImageIndex      =   52
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultas"
            Key             =   "Notes"
            ImageIndex      =   47
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Next"
                  Text            =   "Remesa"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Open"
                  Text            =   "Remito"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Consulta"
                  Text            =   "Depositos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Exit"
            ImageIndex      =   30
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm_VTA_Ctrl_Depositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objRemesa As New clsRemesa
Dim objRemito As New clsRemito
Public pblnMostrar As Boolean
Dim flgControl As String
Dim objDepositos As New clsDepositos

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    
    'Toolbar1.Buttons(1).ButtonMenus(3).Visible = objDepositos.LocalPortaValor(objUsuario.CodigoLocal)
    'Toolbar1.Buttons(6).ButtonMenus(3).Visible = objDepositos.LocalPortaValor(objUsuario.CodigoLocal)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Exit"
            Unload Me
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    Select Case ButtonMenu.Key
        Case "Copy"     'Genera Remesa'
            frm_VTA_Remesas.Show vbModal

        Case "Document" 'Genera Remito'
            frm_VTA_Remitos.Show vbModal
        
        Case "Deposito" 'Graba Deposito del Banco
            frm_VTA_Depositos.Show vbModal

        Case "Next"     'Consulta Remesa'
            'flgControl = "1"
            objDepositos.Control = "1"
            pblnMostrar = True
            frm_VTA_ConsultaDepositos.Show

        Case "Open"     'Cosnulta Remito'
            'flgControl = "0"
            objDepositos.Control = "0"
            pblnMostrar = False
            frm_VTA_ConsultaDepositos.Show
            
        Case "Consulta" 'Consulta Deposito del Banco
           objDepositos.Control = "2"
           Set frm_VTA_ConsultaDepositos.Depositos = objDepositos
           frm_VTA_ConsultaDepositos.Show vbModal
           pblnMostrar = True
    End Select
End Sub
