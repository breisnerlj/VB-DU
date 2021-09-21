VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_cnt_competencia 
   BorderStyle     =   0  'None
   Caption         =   "Reporte de Control de Competencia"
   ClientHeight    =   6930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrilla grdDetalle 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3625
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
      TabIndex        =   5
      Top             =   600
      Width           =   6855
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   40004
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   40004
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Fin"
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   330
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Inicio"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   330
         Width           =   555
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
   Begin vbp_Ventas.ctlGrilla grdIncidencia 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5741
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_cnt_competencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objControl As New clsCntCompetencia
Dim RSClone As oraDynaset

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
   On Error GoTo Control

    Select Case Index
        Case 1:
            'frm_ADM_RegControl.CodigoIncidencia = ""
            frm_ADM_RegControl.Show vbModal
            BuscaValores
        Case 2:
            'Exit Sub
            If grdIncidencia.ApproxCount > 0 Then
                'frm_ADM_RegControl.CodigoIncidencia = ""
                'frm_ADM_RegControl.CodigoIncidencia = grdIncidencia.DataSource("NUM_CONTROL").Value
                frm_CNT_Estadistica.NumeroControl = "" & grdIncidencia.Columns("NUM_CONTROL").Value
                frm_CNT_Estadistica.Show vbModal
            End If
        Case 3:
            ValoresIniciales
        Case 4:
            grdIncidencia.MostrarImprimir
        Case 5:
            grdIncidencia.MostrarExcel
        Case 6:
            grdIncidencia.MostrarEmail
        Case 8:
            Unload Me
        Case 7:
            ''BuscaValores
            Anula
            
        Case Else
            MsgBox "No se encuentra implementado", vbCritical, App.ProductName
    End Select

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Private Sub Form_Load()
   On Error GoTo Control

    setteaFormulario Me
    SeteaGrilla
    inicio
    BuscaValores

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objControl = Nothing
End Sub

Sub BuscaValores()
        Set grdIncidencia.DataSource = objControl.Lista(objUsuario.CodigoLocal, "", dtpInicio.Value, dtpFin.Value)
        Set RSClone = grdIncidencia.DataSource
End Sub

Sub ValoresIniciales()
    'BuscaValores
    SeteaGrilla
    BuscaValores
End Sub

Sub inicio()
    ctlToolBar1.Buttons(2).Caption = "Gráfico"
    ctlToolBar1.Buttons(4).Visible = False
    ctlToolBar1.Buttons(13).Visible = True
    ctlToolBar1.Buttons(13).Caption = "Anular"
    dtpInicio.Value = objUsuario.sysdate - 7
    dtpFin.Value = objUsuario.sysdate
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim o As Integer
    
    arrCampos = Array("NUM_CONTROL", "COD_USUARIO", "DES_USUARIO", "COD_PERIODO", "FCH_REGISTRO")
    arrCaption = Array("N. Control", "Cod. Usuario", "Nombre Usuario", "Periodo", "Fecha Registro")
    arrAncho = Array(800, 800, 2500, 800, 1700)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    grdIncidencia.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    
    arrCampos = Array("NOM_PROVEEDOR", "DES_DIRECCION", "NUM_DOC1", "NUM_DOC2", "VENTA_DIA")
    arrCaption = Array("Competencia", "Dirección", "Doc. Inical", "Doc. Final", "Venta")
    arrAncho = Array(2500, 2500, 1000, 1000, 800)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

End Sub

Private Sub grdIncidencia_RegistroSeleccionado(ByVal DatoColumna0 As String)
    Set grdDetalle.DataSource = objControl.ListaEstadistica(grdIncidencia.Columns("NUM_CONTROL"), "")
End Sub

Sub Anula()
    If MsgBox("Desea Anular el control de la competencia ", vbYesNo + vbDefaultButton2 + vbInformation, App.ProductName) = vbNo Then Exit Sub
    Dim mensaje As String
    mensaje = objControl.Anula(objUsuario.Codigo, grdIncidencia.Columns("NUM_CONTROL"))
    If mensaje = "" Then
        MsgBox "Se grabo satisfactoriamente", vbExclamation, App.ProductName
        ValoresIniciales
        Exit Sub
    Else
        MsgBox mensaje, vbCritical, App.ProductName
        Exit Sub
    End If
End Sub
