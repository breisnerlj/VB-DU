VERSION 5.00
Begin VB.Form frm_VTA_Preview_Remito_PH 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   6375
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CTA CTE"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   645
         Width           =   675
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   555
      End
      Begin VB.Label LblCtaCte 
         Appearance      =   0  'Flat
         BackColor       =   &H00F3E7E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 999999999999999999"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label LblBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H00F3E7E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " xxxxxxxxxxxxxxxxxxxxx"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1920
      Y1              =   2400
      Y2              =   2520
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   1920
      X2              =   1800
      Y1              =   2400
      Y2              =   2520
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   840
      X2              =   960
      Y1              =   2400
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   960
      X2              =   840
      Y1              =   2400
      Y2              =   2520
   End
   Begin VB.Label LblDirEntrega 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3E7E0&
      Caption         =   "HERMES PROCESAMIENTO"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   21
      Top             =   4500
      Width           =   2130
   End
   Begin VB.Label LblLetrasMonto 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3E7E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " xxxxxxxxxxxxxxxxxxxxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   8535
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   11880
      X2              =   11880
      Y1              =   0
      Y2              =   8280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   -240
      X2              =   11880
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image4 
      Height          =   3645
      Left            =   120
      Picture         =   "frm_VTA_Preview_Remito_PH.frx":0000
      Top             =   3960
      Width           =   11880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTE DE SERVICIO"
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
      Left            =   4508
      TabIndex        =   19
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label KeyF3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "F2"
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
      Left            =   3945
      TabIndex        =   18
      Top             =   7920
      Width           =   285
   End
   Begin VB.Label Label17 
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
      Left            =   3585
      TabIndex        =   17
      Top             =   7680
      Width           =   1005
   End
   Begin VB.Label Label16 
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
      Left            =   7665
      TabIndex        =   16
      Top             =   7920
      Width           =   285
   End
   Begin VB.Label Label15 
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
      Left            =   7200
      TabIndex        =   15
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   6480
      Picture         =   "frm_VTA_Preview_Remito_PH.frx":8A23
      Top             =   2805
      Width           =   5415
   End
   Begin VB.Image Image2 
      Height          =   660
      Left            =   120
      Picture         =   "frm_VTA_Preview_Remito_PH.frx":A6DC
      Top             =   2160
      Width           =   11985
   End
   Begin VB.Label LblMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F3E7E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   8760
      TabIndex        =   9
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label LblEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3E7E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " xxxxxxxxxxxxxxxxxxxxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   8535
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENVASE"
      Height          =   195
      Left            =   10440
      TabIndex        =   7
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label LblEntrega 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F3E7E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   9720
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTREGA"
      Height          =   195
      Left            =   8760
      TabIndex        =   5
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label LblEnvase 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F3E7E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   11160
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      Height          =   195
      Left            =   6600
      TabIndex        =   3
      Top             =   1185
      Width           =   525
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F3E7E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dd/mm/yyyy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   7200
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LUGAR"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1185
      Width           =   555
   End
   Begin VB.Label LblLugar 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3E7E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " xxxxxxxxxxxxxxxxxxxxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   120
      Picture         =   "frm_VTA_Preview_Remito_PH.frx":CD34
      Top             =   120
      Width           =   2730
   End
End
Attribute VB_Name = "frm_VTA_Preview_Remito_PH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLetras As New clsLetras
Dim objRemito As New clsRemito

Public pLugar As String
Public pFecha As String
Public pEmpresa As String
Public pDesMonto As String
Public pDirEntrega As String
Public pEntrega As String
Public pEnvase As String
Public pBanco As String
Public pCtaCte As String
Public pMonto As String
Public pPrecinto As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            frm_VTA_Remitos.pblnSalir = False
            Unload Me
        Case vbKeyF3
            frm_VTA_Remitos.pblnSalir = True
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
        lblFecha.Caption = frm_VTA_Remitos.pstrFecha
        pFecha = lblFecha.Caption
        '------------------------------------------------------------------------------------'
        LblLugar.Caption = UCase(objUsuario.NombreLocal) & " " & UCase(objUsuario.Localidad)
        pLugar = LblLugar.Caption
        '------------------------------------------------------------------------------------'
        LblEmpresa.Caption = objRemito.DireccLocal(objUsuario.CodigoLocal) 'UCase(objUsuario.Direccion)
        pEmpresa = LblEmpresa.Caption
        '------------------------------------------------------------------------------------'
        LblBanco.Caption = frm_VTA_Remitos.LblBancoRemito.Caption
        pBanco = LblBanco.Caption
        '------------------------------------------------------------------------------------'
        LblEntrega.Caption = "1"
        pEntrega = LblEntrega.Caption
        '------------------------------------------------------------------------------------'
        pEnvase = LblEnvase.Caption
        '------------------------------------------------------------------------------------'
        pDirEntrega = LblDirEntrega.Caption
        '------------------------------------------------------------------------------------'
        LblCtaCte.Caption = objRemito.CtaBtl(objUsuario.CodigoLocal, _
                                             frm_VTA_Remitos.pstrMon)
        pCtaCte = LblCtaCte.Caption
        '------------------------------------------------------------------------------------'
        If frm_VTA_Remitos.pstrMon = "S" Then     'SOLES'
            LblMonto.Caption = "S/." & " " & Format(frm_VTA_Remitos.pdblSolx_Aux, "#,###,##0.00")
            pMonto = LblMonto.Caption
            '-------------------------------------------------------------------------------------------------------'
            LblLetrasMonto.Caption = UCase(objLetras.NumAletra(frm_VTA_Remitos.pdblSolx_Aux)) & " " & "NUEVOS SOLES"
            pDesMonto = LblLetrasMonto.Caption
        ElseIf frm_VTA_Remitos.pstrMon = "D" Then 'DOLARES'
            LblMonto.Caption = "$" & " " & Format(frm_VTA_Remitos.pdblDolx_Aux, "#,###,##0.00")
            pMonto = LblMonto.Caption
            '-------------------------------------------------------------------------------------------------------'
            LblLetrasMonto.Caption = UCase(objLetras.NumAletra(frm_VTA_Remitos.pdblDolx_Aux)) & " " & "DOLARES AMERICANOS"
            pDesMonto = LblLetrasMonto.Caption
        End If
   Screen.MousePointer = vbDefault
End Sub
