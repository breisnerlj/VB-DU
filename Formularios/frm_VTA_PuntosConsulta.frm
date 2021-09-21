VERSION 5.00
Begin VB.Form frm_VTA_PuntosConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Puntos"
   ClientHeight    =   1680
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "[Esc] Cancelar"
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprime 
      Caption         =   "[F11] Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4410
      Begin VB.TextBox txtNumeroTarjeta 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   425
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_VTA_PuntosConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarEscaneoTarjeta As Boolean

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprime_Click()
    If mvarEscaneoTarjeta Then
        ConsultarPuntosMonedero (Trim$(Me.txtNumeroTarjeta.Text))
    End If
End Sub

Private Sub Form_Activate()
    Me.txtNumeroTarjeta.SetFocus
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub txtNumeroTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
    If EsEscaneado(KeyCode, Shift) = True Then
        mvarEscaneoTarjeta = True
    Else
        mvarEscaneoTarjeta = False
    End If
End Sub

Private Sub txtNumeroTarjeta_KeyPress(KeyAscii As Integer)
    'On Error GoTo CtrlErr
    If KeyAscii = 13 Then
        cmdImprime_Click
    End If
    If KeyAscii = 27 Then
        cmdCancelar_Click
    End If
End Sub

Private Sub ConsultarPuntosMonedero(ByVal vNumeroTarjeta As String)
    Dim oFPC As New clsFPConstante
    Dim oFP As New clsFarmaPuntos
    Dim oTB As New clsTarjetaBean
    Dim strMensaje As String
    Dim intConstante As Integer
    
    On Error GoTo Control
    
    Screen.MousePointer = vbHourglass
    
    intConstante = 184
    
    If SetearImpresoraCupon = False Then
        MsgBox "No tiene la impresora de cupones Instalada o tiene diferente nombre", vbCritical, App.ProductName
        Exit Sub
    End If
    
    Set oTB = oFP.ConsultarSaldo(vNumeroTarjeta, objUsuario.Codigo)
    
    If oTB.EstadoTarjeta = oFPC.EstadoTarjeta.SIN_ESTADO Then
        MsgBox oTB.Mensaje, vbCritical, App.ProductName
        Exit Sub
    End If
    
    'Logo
    If objVenta.EsMFA(objUsuario.CodigoLocal) = True Then
        mdiPrincipal.imgLogoBtl.Picture = mdiPrincipal.ImageList1.ListImages(3).Picture
        Printer.PaintPicture mdiPrincipal.imgLogoBtl, 570, 20, 2950, 850
    Else
        mdiPrincipal.imgLogoBtl.Picture = mdiPrincipal.ImageList1.ListImages(1).Picture
        Printer.PaintPicture mdiPrincipal.imgLogoBtl, 1000, 20, 2150, 950
    End If
    Printer.CurrentY = 1154
    
    'Titulo
    Printer.FontName = "Arial"
    Printer.FontSize = 9
    Printer.Font.Bold = False
    centra_printer "Saldo"
    
    'Nombre cliente
    Printer.FontName = "Printer FontB 10cpi Tall"
    centra_printer oTB.NombreCompleto
    
    'Numero de tarjeta
    Printer.FontName = "Printer FontB 11cpi Tall"
    centra_printer EncriptarNumeroTarjeta(oTB.NumeroTarjeta)

    Printer.CurrentX = 20
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.Print
    centra_printer "Fecha: " & Format$(Now, "dd/mm/YYYY") & Space(10) & "Hora: " & Format$(Now, "hh:nn:ss")
    
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = True
    centra_printer "Puntos Acumulados: " & CStr(oTB.PuntosTotalAcumulados)
    
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = False
    strMensaje = "Para saber su estado actual de puntos, consultar la pagina www.mifarma.com.pe"
    justifica_printer 20, 4000, Printer.CurrentY + intConstante, strMensaje
    
'    Printer.FontSize = 6
'    Printer.CurrentY = Printer.CurrentY + intConstante
'    Printer.Print
'    strMensaje = objUsuario.Codigo & Space(10) & _
'                 App.Major & "." & App.Minor & "." & App.Revision & Space(9) & _
'                 "BTL" & objUsuario.CodigoLocal
'    centra_printer strMensaje
    
    Printer.Print
    Printer.EndDoc
   
    Me.txtNumeroTarjeta.Text = ""
    Me.txtNumeroTarjeta.SetFocus
    Screen.MousePointer = vbDefault
    MsgBox "Recoger el voucher de la cuponera", vbInformation, App.ProductName
    Exit Sub
Control:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub


