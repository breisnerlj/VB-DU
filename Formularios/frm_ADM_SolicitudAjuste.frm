VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_SolicitudAjuste 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrilla grdIncidencia 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5318
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   6855
      Begin vbp_Ventas.ctlDataCombo cboEstado 
         Height          =   315
         Left            =   5280
         TabIndex        =   2
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   40004
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   40004
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado :"
         Height          =   195
         Left            =   4560
         TabIndex        =   9
         Top             =   330
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Fin"
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   330
         Width           =   390
      End
   End
   Begin vbp_Ventas.ctlGrilla grdDetalle 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3625
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_SolicitudAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objControl As New clsAjuste

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
   On Error GoTo Control

    Select Case Index
        Case 1:
            frm_VTA_Solicitud_Ajuste.Show vbModal
            BuscaValores

        Case 2:
            ValoresIniciales

        Case 3:
            ValoresIniciales

        Case 4:
            grdIncidencia.MostrarImprimir

        Case 5:
            grdIncidencia.MostrarExcel

        Case 6:
            grdIncidencia.MostrarEmail

        Case 7:
            Unload Me

        Case Else
            MsgBox "No se encuentra implementado", vbCritical, App.ProductName
    End Select
   Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number

End Sub

Private Sub Form_Load()
   On Error GoTo Control

    ctlToolBar1.Buttons(2).Visible = False

    setteaFormulario Me
    SeteaGrilla
    Inicio

    Set cboEstado.RowSource = objControl.ListaEstados
        cboEstado.BoundColumn = "COD"
        cboEstado.ListField = "DES"
        cboEstado.BoundText = "*"

    BuscaValores

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objControl = Nothing
End Sub

Sub BuscaValores()
    Set grdIncidencia.DataSource = objControl.Lista(objUsuario.CodigoLocal, dtpInicio.Value, dtpFin.Value, cboEstado.BoundText)
End Sub

Sub ValoresIniciales()
    'BuscaValores
    SeteaGrilla
    BuscaValores
End Sub

Sub Inicio()
    dtpInicio.Value = objUsuario.sysdate
    dtpFin.Value = objUsuario.sysdate
End Sub

Private Sub SeteaGrilla()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim i As Integer

    arrCampos = Array("NRO_SOLICITUD", "COD_USUARIO", "DES_USUARIO", "EST_SOLICITUD", "FCH_EMISION")
    arrCaption = Array("N. Solicitud", "Usuario", "Nombre Usuario", "Estado", "Fecha Registro")
    arrAncho = Array(1000, 700, 2200, 800, 1700)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter)
    grdIncidencia.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdIncidencia.Columns("EST_SOLICITUD").FetchStyle = True

    arrCampos = Array("FLG_APROBADO", "COD_PRODUCTO", "DES_PRODUCTO", "COD_TIPO_AJUSTE", "DES_MOTIVO", _
                      "CTD_PRODUCTO_FIS", "CTD_PRODUCTO_FRAC_FIS", "CTD_PRODUCTO_APROBADO", "CTD_PRODUCTO_FRAC_APROBADO", "DES_OBSERVACION")
    arrCaption = Array("Aprobado", "Codigo", "Descripción", "Tipo", "Motivo", _
                       "Ctd Unt", "Ctd. Frac", "Ctd Apr", "Ctd Frac Apr", "Observaciones")
    arrAncho = Array(800, 700, 1000, 500, 2500, 800, 800, 800, 800, 3500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft)
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDetalle.Columns(0).Visible = False
    
    For i = 0 To 9
        grdDetalle.Columns(i).FetchStyle = True
    Next i
End Sub

Private Sub grdDetalle_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
On Error GoTo Control

    Select Case col
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            Select Case grdDetalle.Columns("FLG_APROBADO").CellValue(Bookmark)
                Case 1
                    CellStyle.ForeColor = vbRed
                    CellStyle.Font.Bold = True
                    CellStyle.BackColor = RGB(251, 242, 183)
            End Select
        End Select

   Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

'Private Sub grdDetalle_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
'On Error GoTo Control
'
'    If grdDetalle.Columns("FLG_APROBADO").CellValue(Bookmark) = "0" Then
'        RowStyle.BackColor = RGB(252, 207, 213)
'        RowStyle.ForeColor = RGB(0, 0, 0)
'        grdDetalle.CambiaSeleccionado RGB(252, 207, 213)
'        Exit Sub
'    Else
'        RowStyle.BackColor = RGB(251, 242, 183)
'        RowStyle.ForeColor = RGB(0, 0, 0)
'        grdDetalle.CambiaSeleccionado RGB(251, 242, 183)
''    End If
'
'Exit Sub
'Control:
'    MsgBox Err.Description, vbCritical, App.ProductName
'
'End Sub

Private Sub grdIncidencia_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
On Error GoTo Control
    
    Select Case col
        Case grdIncidencia.Columns("EST_SOLICITUD").ColIndex
            Select Case grdIncidencia.Columns("EST_SOLICITUD").CellValue(Bookmark)
                Case "ATE"
                    CellStyle.BackColor = RGB(50, 175, 50)
                    CellStyle.ForeColor = vbWhite
                Case "EMI"
                    CellStyle.BackColor = RGB(50, 50, 175)
                    CellStyle.ForeColor = vbWhite
                Case "ANU"
                    CellStyle.BackColor = RGB(175, 50, 50)
                    CellStyle.ForeColor = vbWhite
            End Select
    End Select

   Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdIncidencia_RegistroSeleccionado(ByVal DatoColumna0 As String)
    Set grdDetalle.DataSource = objControl.ListaDetalle(grdIncidencia.Columns("NRO_SOLICITUD"))
End Sub
