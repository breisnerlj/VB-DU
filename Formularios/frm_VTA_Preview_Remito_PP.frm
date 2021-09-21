VERSION 5.00
Begin VB.Form frm_VTA_Preview_Remito_PP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Preview Impresión de Remito Prosegur"
   ClientHeight    =   8445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "$"
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
      Left            =   8160
      TabIndex        =   30
      Top             =   6480
      Width           =   135
   End
   Begin VB.Label LblCtdSob 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   29
      Top             =   3000
      Width           =   1530
   End
   Begin VB.Label LblPrecinto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1080
      TabIndex        =   28
      Top             =   6900
      Width           =   1815
   End
   Begin VB.Label LblSmb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "S/."
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
      Left            =   8160
      TabIndex        =   27
      Top             =   5805
      Width           =   300
   End
   Begin VB.Label LblTotOtros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E1EEFB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   300
      Left            =   8520
      TabIndex        =   26
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label LblTotDol 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E1EEFB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   300
      Left            =   8520
      TabIndex        =   25
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label LblTotEfectivo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E1EEFB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   300
      Left            =   8520
      TabIndex        =   24
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   10440
      X2              =   10440
      Y1              =   0
      Y2              =   8400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   10680
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   0
      X2              =   10440
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image5 
      Height          =   870
      Left            =   120
      Picture         =   "frm_VTA_Preview_Remito_PP.frx":0000
      Top             =   6840
      Width           =   10320
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "<Cancelar>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6480
      TabIndex        =   23
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   6945
      TabIndex        =   22
      Top             =   8040
      Width           =   285
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "<Grabar>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2865
      TabIndex        =   21
      Top             =   7800
      Width           =   1005
   End
   Begin VB.Label KeyF3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3225
      TabIndex        =   20
      Top             =   8040
      Width           =   285
   End
   Begin VB.Label LblOtros 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   6480
      Width           =   6735
   End
   Begin VB.Label LblEfectChq 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1080
      TabIndex        =   18
      Top             =   6120
      Width           =   6735
   End
   Begin VB.Label LblEfectivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1080
      TabIndex        =   17
      Top             =   5760
      Width           =   6735
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OTROS"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   6585
      Width           =   570
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHEQUES"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   6225
      Width           =   780
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EFECTIVO"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   5865
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   0
      Picture         =   "frm_VTA_Preview_Remito_PP.frx":31F1
      Top             =   0
      Width           =   3555
   End
   Begin VB.Image Image3 
      Height          =   915
      Left            =   3480
      Picture         =   "frm_VTA_Preview_Remito_PP.frx":9026
      Top             =   4800
      Width           =   6900
   End
   Begin VB.Label LblLocalidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   9000
      TabIndex        =   13
      Top             =   4455
      Width           =   1335
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOCALIDAD"
      Height          =   195
      Index           =   1
      Left            =   8040
      TabIndex        =   12
      Top             =   4560
      Width           =   900
   End
   Begin VB.Label LblNº 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7320
      TabIndex        =   11
      Top             =   4455
      Width           =   615
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº"
      Height          =   195
      Index           =   0
      Left            =   7080
      TabIndex        =   10
      Top             =   4560
      Width           =   180
   End
   Begin VB.Label LblCalle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4680
      TabIndex        =   9
      Top             =   4455
      Width           =   2295
   End
   Begin VB.Label LblPara 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4680
      TabIndex        =   8
      Top             =   4095
      Width           =   5655
   End
   Begin VB.Label LblRecibido 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4680
      TabIndex        =   7
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label LblCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCE7&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   4680
      TabIndex        =   6
      Top             =   3375
      Width           =   5655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CALLE"
      Height          =   195
      Left            =   3600
      TabIndex        =   5
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARA"
      Height          =   195
      Left            =   3600
      TabIndex        =   4
      Top             =   4200
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RECIBIDO DE"
      Height          =   195
      Left            =   3600
      TabIndex        =   3
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   3480
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   2145
      Left            =   3600
      Picture         =   "frm_VTA_Preview_Remito_PP.frx":C03D
      Top             =   1200
      Width           =   6870
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RECIBO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DE TRANSPORTE DE VALORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frm_VTA_Preview_Remito_PP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLetras As New clsLetras
Dim objRemito As New clsRemito
Dim ImporteLetras$
Dim strVarios$
Public blnSalir As Boolean

Public pCliente As String
Public pRecibido As String
Public pPara As String
Public pLugar As String
Public pLocalidad As String
Public pMtoSol As String
Public pMtoDol As String
Public pDesSol As String
Public pDesDol As String
Public pDirecc As String
Public pPrecinto As String
Public pCtdSobres As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF8
            frm_VTA_Remitos.pblnSalir = False
            Set objRemito = Nothing
            Unload Me
        Case vbKeyF3
            frm_VTA_Remitos.pblnSalir = True
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
   On Error GoTo Control

   Screen.MousePointer = vbHourglass
        
        LblCliente.Caption = frm_VTA_Remitos.LblBancoRemito.Caption
        pCliente = LblCliente.Caption
        '-----------------------------------------------------------------------------------'
        LblRecibido.Caption = "BTL" & " " & objUsuario.CodigoLocal
        pRecibido = LblRecibido.Caption
        '-----------------------------------------------------------------------------------'
        LblPara.Caption = frm_VTA_Remitos.pstrPortaValor
        pPara = LblPara.Caption
        '-----------------------------------------------------------------------------------'
        'LblCalle.Caption = objUsuario.Direccion
        LblCalle.Caption = objRemito.DireccLocal(objUsuario.CodigoLocal)
        pLugar = LblCalle.Caption
        '-----------------------------------------------------------------------------------'
        LblLocalidad.Caption = objUsuario.Departamento
        pLocalidad = LblLocalidad.Caption
        '-----------------------------------------------------------------------------------'
        LblCtdSob.Caption = frm_VTA_Remitos.pstrCtdSob
        pCtdSobres = LblCtdSob.Caption
        '-----------------------------------------------------------------------------------'
        LblPrecinto.Caption = frm_VTA_Remitos.pstrNumRecinto
        pPrecinto = LblPrecinto.Caption
        '-----------------------------------------------------------------------------------'
        If frm_VTA_Remitos.pdblSolx_Aux > 0 Then
            LblEfectivo.Caption = UCase(objLetras.NumAletra(frm_VTA_Remitos.pdblSolx_Aux)) & " " & "NUEVOS SOLES"
            pDesSol = LblEfectivo.Caption
            '--------------------------------------------------------------------------------------------------------'
            LblTotEfectivo.Caption = Format(frm_VTA_Remitos.pdblSolx_Aux, "###,###0.00")
            pMtoSol = LblTotEfectivo.Caption
        End If
        
        If frm_VTA_Remitos.pdblDolx_Aux > 0 Then
            LblOtros.Caption = UCase(objLetras.NumAletra(frm_VTA_Remitos.pdblDolx_Aux)) & " " & "DOLARES AMERICANOS"
            pDesDol = LblOtros.Caption
            '--------------------------------------------------------------------------------------------------------'
            LblTotOtros.Caption = Format(frm_VTA_Remitos.pdblDolx_Aux, "###,###0.00")
            pMtoDol = LblTotOtros.Caption
        End If
   Screen.MousePointer = vbDefault

   
   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

