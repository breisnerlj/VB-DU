VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_CierreDiarioDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Diario"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7095
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   4680
         Picture         =   "frm_ADM_CierreDiarioDet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   615
         Left            =   5880
         Picture         =   "frm_ADM_CierreDiarioDet.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3120
         Width           =   1095
      End
      Begin vbp_Ventas.ctlTextBox txtVentas 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ColorDefault    =   12640511
         ColorDefault    =   12640511
         Tipo            =   4
         Alignment       =   2
         Enabled         =   0   'False
         MaxLength       =   8
         EnabledFoco     =   0   'False
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
      Begin vbp_Ventas.ctlTextBox txtPos 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Tipo            =   4
         Alignment       =   2
         MaxLength       =   8
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
      Begin vbp_Ventas.ctlTextBox txtCajeroExpress 
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ColorDefault    =   12640511
         ColorDefault    =   12640511
         Tipo            =   4
         Alignment       =   2
         Enabled         =   0   'False
         MaxLength       =   8
         EnabledFoco     =   0   'False
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
      Begin vbp_Ventas.ctlTextBox txtRemesas 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ColorDefault    =   12640511
         ColorDefault    =   12640511
         Tipo            =   4
         Alignment       =   2
         Enabled         =   0   'False
         MaxLength       =   8
         EnabledFoco     =   0   'False
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
      Begin VB.Label Label6 
         Caption         =   "Remesas:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cajero Express:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "POS:"
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Ventas:"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Tarjetas:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "&Calcular"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpFchCierre 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64880641
         CurrentDate     =   40962
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_ADM_CierreDiarioDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As oraDynaset
Dim objCElectronica As New clsCElectronica
Dim CodigoLocal As String

Public Sub carga(ByVal FechaCierre As String, ByVal CodLocal As String)
On Error GoTo handle
CodigoLocal = CodLocal
If FechaCierre = "" Then
    dtpFchCierre.Value = objUsuario.sysdate - 1
Else
    dtpFchCierre.Value = FechaCierre
    dtpFchCierre.Enabled = False
    cmdCalcular.Enabled = False
    cmdGrabar.Enabled = False
    txtPos.Enabled = False
    Set rs = objCElectronica.ListaCierreDiarioDetalle(FechaCierre, CodigoLocal)
    txtVentas.Text = rs("MTO_TARJETAS").Value
    txtPos.Text = rs("MTO_POS").Value
    txtCajeroExpress.Text = rs("CTD_CAJERO").Value
    txtRemesas.Text = rs("INDICADOR_REMESA").Value
End If

    Me.Show vbModal
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdCalcular_Click()
On Error GoTo handle
     Set rs = objCElectronica.ListaNuevoCierreDiarioDetalle(CStr(dtpFchCierre.Value), CodigoLocal)
     If rs.RecordCount > 0 Then
        txtVentas.Text = rs("MTO_TARJETAS").Value
        txtPos.Text = rs("MTO_POS").Value
        txtCajeroExpress.Text = rs("CTD_CAJERO").Value
        txtRemesas.Text = rs("INDICADOR_REMESA").Value
        cmdGrabar.Enabled = True
        cmdCalcular.Enabled = False
        txtPos.Enabled = True
        dtpFchCierre.Enabled = False
        txtPos.Focus
     Else
        MsgBox "No se encontraron Datos en dicha Fecha de Cierre", vbCritical + vbInformation, "Aviso"
     End If
     Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()
    txtPos.Enabled = False
    cmdGrabar.Enabled = False
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo handle
    Dim strMensaje As String
    
        If txtPos.Text = "" Then
           MsgBox "El Monto del POS no puede estar Vacio", vbCritical + vbInformation, "Aviso"
           Exit Sub
        End If
        
        If txtPos.Text = "0" Then
           If MsgBox("El Monto del POS es cero, Esta seguro de Grabar.. ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
           Else
               txtPos.Focus
               Exit Sub
           End If
        End If
        
        strMensaje = objCElectronica.GrabaCierreDiario(CStr(dtpFchCierre.Value), CodigoLocal, objUsuario.Codigo, CDbl(txtVentas.Text), CDbl(txtPos.Text), CInt(txtCajeroExpress.Text), IIf(txtRemesas.Text = "SI", "1", "0"))
        If strMensaje = "" Then
           MsgBox "Se grabo satisfactoriamente el Cierre Diario", vbInformation, App.ProductName
           Unload Me
        Else
            MsgBox strMensaje, vbCritical, App.ProductName
        End If
        Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCElectronica = Nothing
End Sub


