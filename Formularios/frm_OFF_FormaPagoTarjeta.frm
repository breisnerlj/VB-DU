VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_OFF_FormaPagoTarjeta 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   180
      TabIndex        =   11
      Top             =   480
      Width           =   6675
      Begin vbp_Ventas.ctlTextBox txtImporte 
         Height          =   315
         Left            =   1740
         TabIndex        =   5
         Top             =   2040
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Tipo            =   4
         Alignment       =   1
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
      Begin vbp_Ventas.ctlTextBox txtNumAutorizacion 
         Height          =   315
         Left            =   1740
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Tipo            =   3
         MaxLength       =   6
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
      Begin vbp_Ventas.ctlTextBox txtNroCuota 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   1140
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         Tipo            =   3
         Alignment       =   1
         MaxLength       =   2
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
      Begin vbp_Ventas.ctlTextBox txtNroTar 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         Tipo            =   3
         MaxLength       =   21
         TABAuto         =   0   'False
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
      Begin VB.ComboBox cboTipoCuota 
         Height          =   315
         ItemData        =   "frm_OFF_FormaPagoTarjeta.frx":0000
         Left            =   2280
         List            =   "frm_OFF_FormaPagoTarjeta.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   2295
      End
      Begin MSMask.MaskEdBox mskVencimiento 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   690
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Autorización :"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1627
         Width           =   1410
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. Cuotas : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. Tarjeta : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Vencimiento : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Importe S/. : "
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
         Left            =   180
         TabIndex        =   12
         Top             =   2070
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4320
      Picture         =   "frm_OFF_FormaPagoTarjeta.frx":0027
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5625
      Picture         =   "frm_OFF_FormaPagoTarjeta.frx":05B1
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Tarjeta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   4
      Left            =   420
      TabIndex        =   10
      Top             =   60
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_OFF_FormaPagoTarjeta.frx":0B3B
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   12
      Left            =   4260
      TabIndex        =   9
      Top             =   6780
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   11
      Left            =   5970
      TabIndex        =   6
      Top             =   6780
      Width           =   390
   End
End
Attribute VB_Name = "frm_OFF_FormaPagoTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public strCodF As String
Dim strPagaCon As String
Dim strCodMoneda As String
Dim dblImpTot As Double
Dim intCount As Integer
Dim dblTipoCambio As Double

'Devuelve el último días del Mes
Private Function Mes_Anterior(Fecha As Variant) As Date
  
    If IsDate(Fecha) Then
        Mes_Anterior = DateSerial(Year(Fecha), Month(Fecha), 1)
        Mes_Anterior = DateAdd("d", -1, Mes_Anterior)
    End If
  
End Function

Private Sub cboTipoCuota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click()

On Error GoTo CtrlErr

    If fmod(txtNroTar.Text) = 0 Then
        MsgBox "El número de Tarjeta no es válido", _
                vbCritical, App.ProductName: txtNroTar.Focus: Exit Sub
    End If

    If Trim(txtNroTar.Text) = "" Then _
        MsgBox "El número de Tarjeta no puede ser vacío.", _
                vbCritical, App.ProductName: txtNroTar.Focus: Exit Sub

    If IsDate(mskVencimiento.Text) = False Then _
        MsgBox "La fecha de vencimiento no es válida.", _
                vbCritical, App.ProductName: mskVencimiento.SetFocus: Exit Sub
                
    If Now > Mes_Anterior(CDate(mskVencimiento.Text)) Then _
        MsgBox "La tarjeta está vencida.", vbCritical, _
                App.ProductName: mskVencimiento.SetFocus: Exit Sub

    If Val(Trim(txtNroCuota.Text)) <= 0 Then _
        MsgBox "El número de cuotas debe ser mayor que cero.", _
                vbCritical, App.ProductName: txtNumAutorizacion.Focus: Exit Sub

    If Trim(txtNumAutorizacion.Text) = "" Then _
        MsgBox "El número de autorización no puede ser vacío.", _
                vbCritical, App.ProductName: txtNumAutorizacion.Focus: Exit Sub

    If Val(Trim(txtImporte.Text)) <= 0 Then _
        MsgBox "El importe debe ser mayor que cero.", _
                vbCritical, App.ProductName: txtImporte.Focus: Exit Sub

    If Val(Trim(txtImporte.Text)) > objOFFVenta.TotalVenta Then _
        MsgBox "El importe no puede ser mayor que el importe total de la venta.", _
                vbCritical, App.ProductName: txtImporte.Focus: Exit Sub
                
                
                
                

    dblTipoCambio = 1

    objOFFVenta.AgregaPagoVenta strCodF, strCodMoneda, _
            Val(strPagaCon), _
            dblImpTot, _
            txtNroTar.Text, _
            mskVencimiento.Text, _
            Val(txtNroCuota.Text), _
            dblTipoCambio, _
            Val(left(cboTipoCuota.Text, 1)), _
            txtNumAutorizacion.Text

    Unload Me
    
    frm_OFF_FormaPago.grdListaFP.Rebind
    
    frm_OFF_Principal.MostrarTotales
    
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
    
    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0
    
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then Call cmdAceptar_Click
        ''''    DESDE ACA COPIA
        Case tmpCtrl And vbKeyQ And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyM And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyE And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyD
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyC
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF5
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF6 And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF7
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF8 And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyX And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyF
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
    End Select

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()

setteaFormulario Me

cboTipoCuota.ListIndex = 0

End Sub


Private Sub mskVencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImporte_Change()
    strPagaCon = txtImporte.Text
    Select Case strCodF
           Case "2"
                dblImpTot = Round(Val(strPagaCon), 2)
                strCodMoneda = "1"
    End Select

End Sub

Private Sub txtNroTar_GotFocus()
  intCount = 0
End Sub

Private Sub txtNroTar_KeyPress(KeyAscii As Integer)
Dim n As Integer

On Error GoTo CtrlErr


    If intCount > 1 And KeyAscii <> 13 Then KeyAscii = 0
    If Chr(KeyAscii) = "&" Then intCount = intCount + 1


    If KeyAscii = 13 Then
        
        n = fmod(txtNroTar.Text)
        If n = 1 Then SendKeys "{TAB}" 'txtNroTar.TABAuto = True
        
        If fmod(txtNroTar.Text) = 0 Then
            MsgBox "El número de Tarjeta no es válido", _
                    vbCritical, App.ProductName: txtNroTar.Focus: Exit Sub
        End If
        
        If Trim(txtNroTar.Text) = "" Then _
            MsgBox "El número de Tarjeta no puede ser vacío.", _
                    vbCritical, App.ProductName: txtNroTar.Focus: Exit Sub
    
    End If
    
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub txtNroTar_LostFocus()
txtNroTar.TABAuto = False
End Sub
