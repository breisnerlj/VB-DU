VERSION 5.00
Begin VB.Form frm_ADM_Moneda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantemiento de Moneda"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frm_ADM_Moneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmrGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin vbp_Ventas.ctlTextBox TXTDesMoneda 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   609
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
      Begin vbp_Ventas.ctlTextBox TxtCodMoneda 
         Height          =   345
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
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
      Begin vbp_Ventas.ctlTextBox TxtDesSmb 
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   609
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
      Begin vbp_Ventas.ctlTextBox TxtDesLng 
         Height          =   345
         Left            =   1200
         TabIndex        =   3
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   609
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
      Begin VB.Label Label4 
         Caption         =   "Longitud"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Simbolo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_ADM_Moneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMoneda As New clsMoneda
Dim odynMoneda As OraDynaset
Dim strActivo As String

Private Sub Form_Load()
    TxtCodMoneda.Enabled = False
    If frm_ADM_LstMoneda.pBlnMuestra = False Then
        Set odynMoneda = objMoneda.Lista
        TxtCodMoneda.Text = objMoneda.MaxCodMoneda
    Else
        Set odynMoneda = objMoneda.Lista(frm_ADM_LstMoneda.grdMoneda.Columns("COD_MONEDA").Value)
        CargaDatos
    End If
End Sub

Sub CargaDatos()
    TxtCodMoneda.Text = odynMoneda("COD_MONEDA").Value
    TXTDesMoneda.Text = odynMoneda("DES_MONEDA").Value
    TxtDesSmb.Text = odynMoneda("SMB_MONEDA").Value
    TxtDesLng.Text = odynMoneda("DES_LG_MONEDA").Value
    chkActivo.Value = odynMoneda("FLG_ACTIVO").Value
End Sub

Private Sub cmrGrabar_Click()
    Dim CtrlErr As String
    strActivo = IIf(chkActivo.Value = "1", "1", "0")
    
    CtrlErr = objMoneda.Graba(TxtCodMoneda.Text, _
                              TXTDesMoneda.Text, _
                              TxtDesSmb.Text, _
                              TxtDesLng.Text, _
                              strActivo, _
                              objUsuario.Codigo)
    
    If CtrlErr = "" Then
        MsgBox "Se Grabo con exito la Moneda", vbInformation, App.ProductName
        Set frm_ADM_LstMoneda.grdMoneda.DataSource = objMoneda.Lista
        Unload Me
    Else
        MsgBox CtrlErr, vbCritical, App.ProductName
    End If
End Sub
