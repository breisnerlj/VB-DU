VERSION 5.00
Begin VB.Form frm_VTA_DatoReceta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del médico"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frm_VTA_DatoReceta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla grdMedico 
      Height          =   1095
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   1931
   End
   Begin vbp_Ventas.ctlTextBox txtCMP 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   5400
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código CMP :"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   840
   End
End
Attribute VB_Name = "frm_VTA_DatoReceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMedico As clsMedico

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    txtCMP.Tipo = Entero
    
    
    SeteaGrilla
    
End Sub



Private Sub grdMedico_DblClick()
    On Error GoTo CtrlErr
    If grdMedico.ApproxCount = 0 Then
        Exit Sub
    End If
    
    objVenta.CodMedico = grdMedico.Columns("COD_MEDICO").Value
    Unload Me
    

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbInformation, App.ProductName
End Sub

Private Sub grdMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdMedico_DblClick
    End Select
End Sub




Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim columna As TrueDBGrid70.Column
  
  
    arrCampos = Array("COD_MEDICO", "NUM_CMP", "NOM_MEDICO", "ESPECIALIDAD", "DES_TIPO_COLEGIO")
                      
    arrCaption = Array("Codigo", "NºCMP", "Nombre", "Especialidad", "Tipo Colegio")
    
    arrAncho = Array(0, 0, 2700, 0, 2300)
    
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    
    grdMedico.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    
            For Each columna In grdMedico.Columns
                columna.AllowSizing = False
            Next
    
    
    grdMedico.Columns(0).Visible = False
    grdMedico.Columns(1).Visible = False
    grdMedico.Columns(3).Visible = False
    


End Sub

Private Sub txtCMP_KeyPress(KeyAscii As Integer)
On Error GoTo CtrlErr

    If KeyAscii = 13 Then
        
        Set objMedico = New clsMedico
        
        Set grdMedico.DataSource = objMedico.ListaCMP(txtCMP.Text)
        
        Set objMedico = Nothing
    
        If grdMedico.ApproxCount < 2 Then objVenta.CodMedico = grdMedico.Columns("COD_MEDICO").Value: Unload Me
    End If
    Exit Sub
    
CtrlErr:
    Select Case Err.Number
        Case 94
            MsgBox "No se encontro el número de CMP", vbInformation, App.ProductName
        Case Else
            MsgBox Err.Description, vbCritical, App.ProductName
    End Select
    
End Sub
