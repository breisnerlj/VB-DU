VERSION 5.00
Begin VB.Form frm_DLV_Verificacion_Trajeta_Chk 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_DLV_Verificacion_Trajeta_Chk.frx":0000
   ScaleHeight     =   1800
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox chkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manual"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin vbp_Ventas.ctlTextBox TxtNumAutor 
         Height          =   495
         Left            =   150
         TabIndex        =   1
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         Alignment       =   2
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbp_Ventas.ctlDataCombo CboTarjetas 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Verificación de Tarjeta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   555
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Numero Verificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   450
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_DLV_Verificacion_Trajeta_Chk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmVTarjeta As frm_DLV_Verificacion_Tarjeta
Dim objFpago As New clsFormaPago

Private Sub CboTarjetas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Set CboTarjetas.RowSource = objFpago.ListaTarjetasVenta
    CboTarjetas.BoundColumn = "COD"
    CboTarjetas.ListField = "DES"
    CboTarjetas.BoundText = Trim(frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 0))
    
    If frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 3) <> "" Then
        frm_DLV_Verificacion_Trajeta_Chk.TxtNumAutor.Text = frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 3)
    End If
    'Trim(CboTarjetas.BoundText) = frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 0)
    'frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 1) = Trim(CboTarjetas.Text)
    If frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 10) <> "" Then
        frm_DLV_Verificacion_Trajeta_Chk.chkPos.Value = Val(Mid(frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 10), 1, 1))
    End If
    
End Sub

Private Sub TxtNumAutor_KeyPress(KeyAscii As Integer)
    On Error GoTo CtrlErr
    TxtNumAutor.Tipo = AlfaNumerico
    If KeyAscii = 13 Then
        If frmVTarjeta.objFormaPago.TarjetaVerif.UpperBound(1) = -1 Then Exit Sub
            If CboTarjetas.BoundText = "*" Then MsgBox "Seleccione Tarjeta con la que pagara", vbCritical, "Tarjeta": Exit Sub
            
            frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 3) = Trim(TxtNumAutor.Text)
            frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 0) = Trim(CboTarjetas.BoundText)
            frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 1) = Trim(CboTarjetas.Text)
            frmVTarjeta.objFormaPago.TarjetaVerif(frmVTarjeta.grdTarjetas.Row, 10) = IIf(chkPos.Value = "1", "1", "0") & " - " & IIf(chkPos.Value = "1", "Manual", "POS")
            
            
            frmVTarjeta.grdTarjetas.Columns(10).Refetch
            frmVTarjeta.grdTarjetas.Columns(3).Refetch
            frmVTarjeta.grdTarjetas.Columns(0).Refetch
            frmVTarjeta.grdTarjetas.Columns(1).Refetch
        Unload Me
    End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

