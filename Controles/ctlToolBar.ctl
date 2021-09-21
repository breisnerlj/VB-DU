VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Begin VB.UserControl ctlToolBar 
   Alignable       =   -1  'True
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   ScaleHeight     =   600
   ScaleWidth      =   9660
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1058
      BandCount       =   1
      _CBWidth        =   9615
      _CBHeight       =   600
      _Version        =   "6.7.9816"
      Child1          =   "Toolbar1"
      MinWidth1       =   9375
      MinHeight1      =   540
      Width1          =   9555
      NewRow1         =   0   'False
      BandStyle1      =   1
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   540
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   953
         ButtonWidth     =   1429
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Object.ToolTipText     =   "Nuevo"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Editar"
               Object.ToolTipText     =   "Editar"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Actualizar"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               Object.Width           =   400
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Excel"
               Object.ToolTipText     =   "Excel"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "E-mail"
               Object.ToolTipText     =   "Enviar e-mail"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               Object.Width           =   400
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Object.ToolTipText     =   "Guardar"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Object.ToolTipText     =   "Cancelar"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               Object.Width           =   400
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Object.ToolTipText     =   "Eliminar"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               Object.Width           =   400
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":02DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":045E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":0762
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":08BC
            Key             =   "new"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":0A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":0B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":0E24
            Key             =   "pend"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":0F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":10D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":1BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":1CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":1E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":1FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":210A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":2424
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":273E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlToolBar.ctx":2898
            Key             =   "check"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum tlbTipoBoton
    Nuevo = 1
    Modificar = 2
    Buscar = 3
    tb_Actualizar = 4
    Imprimir = 6
    tb_Excel = 7
    tb_email = 8
    Grabar = 10
    Cancelar = 11
    Eliminar = 13
    salir = 15
End Enum

Public Enum tlbTipoEstado
    Activar = 1
    Desactivar = 2
End Enum

Public Enum tlbModoToolBar
    Modo1 = 1
    modo2 = 2
    Modo3 = 3
    modo4 = 4
    Modo5 = 5
    Modo6 = 6
    Modo7 = 7
    Mantenimiento = 8
    Guias = 9
    GuardarCancelar = 10
End Enum

Private intModo As tlbModoToolBar
Private blnEfecto As Boolean
Public Event Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Integer, j As Integer, temp As tlbTipoBoton
    temp = Button.Index
    For j = 1 To Button.Index
        If Toolbar1.Buttons(j).Visible = True And _
           Not Toolbar1.Buttons(j).Style = tbrSeparator Then i = i + 1
    Next
    Select Case temp
        Case 1, 2
            If blnEfecto Then sub_ModoGrabar
        Case 10, 11
            If blnEfecto Then sub_ModoNormal
    End Select
    
    RaiseEvent Click(temp, i)
End Sub

Public Sub PrimerRegistro(ByVal vEstado As tlbTipoEstado)
    Dim temp As Boolean
    If vEstado = Activar Then temp = True Else temp = False
    Toolbar1.Buttons(1).Enabled = temp: Toolbar1.Buttons(2).Enabled = temp
End Sub

Public Sub UltimoRegistro(ByVal vEstado As tlbTipoEstado)
    Dim temp As Boolean
    If vEstado = Activar Then temp = True Else temp = False
    Toolbar1.Buttons(3).Enabled = temp: Toolbar1.Buttons(4).Enabled = temp
End Sub


'Especificar el boton y el estado de este boton
Public Function VisibleBoton(ByVal boton As tlbTipoBoton, ByVal Modo As Boolean)
    Toolbar1.Buttons(boton).Visible = Modo
End Function
Public Function EnabledBoton(ByVal boton As tlbTipoBoton, ByVal Modo As Boolean)
    Toolbar1.Buttons(boton).Enabled = Modo
End Function
Public Function StyleBoton(ByVal boton As tlbTipoBoton, ByVal Modo As MSComctlLib.ButtonStyleConstants)
    Toolbar1.Buttons(boton).Style = Modo
End Function





Private Sub UserControl_Initialize()
    intModo = 1
    sub_Modo1
    blnEfecto = True
End Sub

Private Sub UserControl_Resize()
    CoolBar1.Width = UserControl.Width
    UserControl.Height = CoolBar1.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ModoBotones", intModo, 1
    PropBag.WriteProperty "EnabledEfecto", blnEfecto, True
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    intModo = PropBag.ReadProperty("ModoBotones", 1)
    sub_SeleccionarModo intModo
    blnEfecto = PropBag.ReadProperty("EnabledEfecto", True)
End Sub





Public Property Get Buttons() As MSComctlLib.Buttons
   Set Buttons = Toolbar1.Buttons
End Property

Public Property Get ModoBotones() As tlbModoToolBar
    ModoBotones = intModo
End Property
Public Property Let ModoBotones(ByVal vNewValue As tlbModoToolBar)
    intModo = vNewValue
    sub_SeleccionarModo vNewValue
End Property

Private Sub sub_SeleccionarModo(ByVal Index As Integer)
    Toolbar1.Buttons(13).Caption = "Eliminar"
    Select Case Index
        Case 1: sub_Modo1
        Case 2: sub_Modo2
        Case 3: sub_Modo3
        Case 4: sub_Modo4
        Case 5: sub_Modo5
        Case 6: sub_Modo6
        Case 7: sub_Modo7
        Case 8: sub_Mantenimiento
        Case 9: sub_Guias
        Case 10: sub_GuardarCancelar
    End Select
End Sub
Private Sub sub_Modo1()
    Dim i As Byte
    For i = 1 To 15: Toolbar1.Buttons(i).Visible = True: Next
End Sub
Private Sub sub_Modo2()
    Dim i As Byte
    For i = 1 To 5: Toolbar1.Buttons(i).Visible = False: Next
    For i = 6 To 12: Toolbar1.Buttons(i).Visible = True: Next
    For i = 13 To 14: Toolbar1.Buttons(i).Visible = False: Next
    Toolbar1.Buttons(15).Visible = True
End Sub
Private Sub sub_Modo3()
    Dim i As Byte
    For i = 1 To 9: Toolbar1.Buttons(i).Visible = True: Next
    For i = 10 To 14: Toolbar1.Buttons(i).Visible = False: Next
    Toolbar1.Buttons(15).Visible = True
End Sub
Private Sub sub_Modo4()
    Dim i As Byte
    For i = 1 To 9: Toolbar1.Buttons(i).Visible = False: Next
    For i = 10 To 15: Toolbar1.Buttons(i).Visible = True: Next
End Sub
Private Sub sub_Modo5()
    Dim i As Byte
    For i = 1 To 5: Toolbar1.Buttons(i).Visible = True: Next
    For i = 6 To 9: Toolbar1.Buttons(i).Visible = False: Next
    For i = 10 To 15: Toolbar1.Buttons(i).Visible = True: Next
End Sub
Private Sub sub_Modo6()
    Dim i As Byte
    For i = 1 To 9: Toolbar1.Buttons(i).Visible = False: Next
    For i = 10 To 12: Toolbar1.Buttons(i).Visible = True: Next
    For i = 13 To 14: Toolbar1.Buttons(i).Visible = False: Next
    Toolbar1.Buttons(15).Visible = True
End Sub
Private Sub sub_Modo7()
    Dim i As Byte
    For i = 1 To 2: Toolbar1.Buttons(i).Visible = False: Next
    For i = 3 To 9: Toolbar1.Buttons(i).Visible = True: Next
    For i = 10 To 14: Toolbar1.Buttons(i).Visible = False: Next
    Toolbar1.Buttons(15).Visible = True
End Sub

Private Sub sub_Mantenimiento()
    Dim i As Byte
    For i = 1 To 9: Toolbar1.Buttons(i).Visible = True: Next
    For i = 10 To 14: Toolbar1.Buttons(i).Visible = False: Next
    Toolbar1.Buttons(15).Visible = True
    Toolbar1.Buttons(13).Visible = True
    Toolbar1.Buttons(13).Caption = "Anular"
    Toolbar1.Buttons(13).ToolTipText = "Anular"
    Toolbar1.Buttons(8).Visible = False
    
End Sub

Private Sub sub_Guias()
    Dim i As Byte
    For i = 1 To 9: Toolbar1.Buttons(i).Visible = True: Next
    For i = 10 To 14: Toolbar1.Buttons(i).Visible = False: Next
    Toolbar1.Buttons(15).Visible = True
    Toolbar1.Buttons(13).Visible = True
    Toolbar1.Buttons(13).Caption = "Anular"
    Toolbar1.Buttons(13).ToolTipText = "Anular"
    Toolbar1.Buttons(2).Caption = "Recep."
    Toolbar1.Buttons(2).ToolTipText = "Recepcionar"
    Toolbar1.Buttons(2).Image = "check"
    Toolbar1.Buttons(8).Visible = False
    
End Sub

Private Sub sub_GuardarCancelar()
    Dim i As Byte
    For i = 1 To 9: Toolbar1.Buttons(i).Visible = False: Next
    For i = 10 To 12: Toolbar1.Buttons(i).Visible = True: Next
    For i = 13 To 15: Toolbar1.Buttons(i).Visible = False: Next
    'Toolbar1.Buttons(15).Visible = True
End Sub

'Prepara el control para Editar o Agregar Registro.
'Solo los botones grabar, cancelar y salir deben estar activos.
Public Sub sub_ModoGrabar()
    Dim i As Byte
    For i = 1 To 14
        Toolbar1.Buttons(i).Enabled = False
    Next
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(11).Enabled = True
End Sub
Private Sub sub_ModoNormal()
    Dim i As Byte
    For i = 1 To 15: Toolbar1.Buttons(i).Enabled = True: Next
End Sub
Public Property Get EnabledEfecto() As Boolean
    EnabledEfecto = blnEfecto
End Property

Public Property Let EnabledEfecto(ByVal vNewValue As Boolean)
    blnEfecto = vNewValue
End Property




