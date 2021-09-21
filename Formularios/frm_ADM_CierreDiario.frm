VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_CierreDiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cierre Diario"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8415
      Begin VB.Frame Frame4 
         Caption         =   "&Estado"
         Height          =   735
         Left            =   4440
         TabIndex        =   7
         Top             =   120
         Width           =   2775
         Begin vbp_Ventas.ctlDataCombo cboEstado 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            MatchEntry      =   1
         End
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   40004
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   40004
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Fin"
         Height          =   195
         Left            =   2400
         TabIndex        =   3
         Top             =   450
         Width           =   390
      End
   End
   Begin vbp_Ventas.ctlToolBar tlbMenu 
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
   Begin vbp_Ventas.ctlGrilla grdCierreDiario 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7646
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_CierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim objCElectronica As New clsCElectronica

Private Sub Form_Load()
On Error GoTo Handle
    inicio
    Consulta
    SeteaGrilla
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub inicio()
    dtpInicio.Value = objUsuario.sysdate - 7
    dtpFin.Value = objUsuario.sysdate

    Set cboEstado.RowSource = objCElectronica.ListaEstado("Todos")
        cboEstado.BoundColumn = "COD"
        cboEstado.ListField = "DES"
        cboEstado.Text = "Todos"
End Sub

Sub Consulta()
    Set grdCierreDiario.DataSource = objCElectronica.ListaCierres(objUsuario.CodigoLocal, cboEstado.BoundText, dtpInicio.Value, dtpFin.Value)
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant

    arrCampos = Array("DiaCierre", "Usuario", "FCH_PROCESO", "FCH_ARCHIVO", "DesEstado")
    arrCaption = Array("Día de Cierre", "Usuario", "Fecha Proceso", "Fecha Archivo", "Estado")
    arrAncho = Array(1100, 2000, 2000, 2000, 1500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    grdCierreDiario.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objCElectronica = Nothing
End Sub

Private Sub TlbMenu_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
 On Error GoTo Control

    Select Case Index
        Case 1:
                frm_ADM_CierreDiarioDet.carga "", objUsuario.CodigoLocal
                Consulta
                
        Case 2:
                If grdCierreDiario.ApproxCount > 0 Then
                      frm_ADM_CierreDiarioDet.carga grdCierreDiario.Columns(0).Value, objUsuario.CodigoLocal
                End If
        Case 3:
                Consulta

        Case 4:
                Consulta

        Case 5:
                grdCierreDiario.MostrarImprimir
                
        Case 6:
                grdCierreDiario.MostrarExcel
        
        Case 7:
                grdCierreDiario.MostrarEmail
            
        Case 8:
            Unload Me

        Case Else
            MsgBox "No se encuentra implementado", vbCritical, App.ProductName
    End Select
   Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub
