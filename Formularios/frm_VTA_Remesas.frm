VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_VTA_Remesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Remesas"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frm_VTA_Remesas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraRemesa 
      BackColor       =   &H80000004&
      Caption         =   "Remesa"
      ForeColor       =   &H00800000&
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6925
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   465
         Left            =   4193
         TabIndex        =   20
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton CmOkRemesa 
         Height          =   300
         Left            =   4920
         Picture         =   "frm_VTA_Remesas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   400
         Width           =   380
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         Caption         =   "Sobres"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4680
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   180
         Begin VB.Label LblSobreSolRemesa 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFBFA&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1080
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label LblSobreDolRemesa 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFBFA&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1080
            TabIndex        =   7
            Top             =   525
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000004&
            Caption         =   "Soles"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000004&
            Caption         =   "Dolares"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   580
            Width           =   615
         End
      End
      Begin vbp_Ventas.ctlTextBox TxtObsRemesa 
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   4800
         Width           =   6735
         _ExtentX        =   11880
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
      Begin vbp_Ventas.ctlGrillaArray grdRemesa 
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   3836
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboCajeroRemesa 
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin MSComCtl2.DTPicker dtpFchIniRemesa 
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   16908289
         CurrentDate     =   39016
      End
      Begin MSComCtl2.DTPicker dtpFchFinRemesa 
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   16908289
         CurrentDate     =   39016
      End
      Begin VB.CommandButton CmdGrabaRemesa 
         Caption         =   "&Grabar"
         Height          =   465
         Left            =   1793
         TabIndex        =   1
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label LblLiquidacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2220
         TabIndex        =   23
         Top             =   1410
         Width           =   3200
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F1"
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " F3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   4560
         TabIndex        =   21
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         Caption         =   "Hasta"
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000004&
         Caption         =   "Cajero"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   910
         Width           =   615
      End
      Begin VB.Label LblNomCajeroRemesa 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFBFA&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   310
         Left            =   720
         TabIndex        =   16
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label LblCajaRemesa 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFBFA&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   310
         Left            =   720
         TabIndex        =   15
         Top             =   1410
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " F2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   2160
         TabIndex        =   13
         Top             =   5880
         Width           =   285
      End
   End
End
Attribute VB_Name = "frm_VTA_Remesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objRemesa As New clsRemesa
Dim objMaquina As New clsMaquina
Dim odynR1 As oraDynaset
'Dim odynr2 As oraDynaset
Public pColR As String

Private Sub Form_Load()
    dtpFchIniRemesa.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    dtpFchFinRemesa.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    SeteaGrilla
End Sub

Sub Nuevo()
    grdRemesa.Delete
    ctlCboCajeroRemesa.BoundText = ""
    If ctlCboCajeroRemesa.BoundText = "" Then LblCajaRemesa.Caption = "": LblNomCajeroRemesa.Caption = ""
    LblSobreSolRemesa.Caption = ""
    LblSobreDolRemesa.Caption = ""
    TxtObsRemesa.Text = ""
    grdRemesa.Columns(5).FooterText = "": grdRemesa.Columns(6).FooterText = ""
End Sub

Private Sub CmOkRemesa_Click()
    CargaCombo
End Sub

Sub CargaCombo()
    Set odynR1 = objLiquidacion.ListaUsurioxArqueo(objUsuario.CodigoEmpresa, _
                                                                objUsuario.CodigoLocal, _
                                                                Format(dtpFchIniRemesa.Value, "dd/mm/yyyy"), _
                                                                Format(dtpFchFinRemesa.Value, "dd/mm/yyyy"))
                                                                
    ''Set ctlCboCajeroRemesa.RowSource = objLiquidacion.ListaUsurioxArqueo(objUsuario.CodigoEmpresa, _
                                                                objUsuario.CodigoLocal, _
                                                                Format(dtpFchIniRemesa.Value, "dd/mm/yyyy"), _
                                                                Format(dtpFchFinRemesa.Value, "dd/mm/yyyy"))
                                                                
    Set ctlCboCajeroRemesa.RowSource = odynR1
    ctlCboCajeroRemesa.ListField = "NOMB"
    ctlCboCajeroRemesa.BoundColumn = "COD"
    ctlCboCajeroRemesa.ListField2 = "COD_LIQUIDACION"
    'ctlCboCajeroRemesa.BoundText = "NOMB"

End Sub

Private Sub ctlCboCajeroRemesa_Change()
    If odynR1.RecordCount <= 0 Then Exit Sub
    If ctlCboCajeroRemesa.Text = "" Then Exit Sub
                                                                        
'    Dim LongCaja As String
'    Dim Longit As String
'
'    LongCaja = Len(Right(ctlCboCajeroRemesa.Text, 8))
'    Longit = Len(ctlCboCajeroRemesa.Text)
'    LblNomCajeroRemesa.Caption = Mid(ctlCboCajeroRemesa.Text, 1, Longit - LongCaja)
'    LblCajaRemesa.Caption = Right(ctlCboCajeroRemesa.Text, 8)
                                                           
    
    '----- Para sacar la caja -----'
    Dim strMaquina As String
    Dim k As Integer
    Dim Valor As String
    
    k = 1
    Valor = "": strMaquina = ""
    For k = 1 To Len(ctlCboCajeroRemesa.Text)
        Valor = Mid(ctlCboCajeroRemesa.Text, k, 1)
        If Valor <> " " Then
            strMaquina = strMaquina & Valor
          Else
            GoTo abajo
        End If
    Next k
abajo:
    strMaquina = Len(strMaquina)
    LblCajaRemesa.Caption = left(ctlCboCajeroRemesa.Text, strMaquina)
    
    
    '---- Para sacar el nombre ----'
    Dim strNombini As String
    Dim strNombfin As String
    Dim strNombini_aux As String
    
    strNombini = Val(strMaquina) + 2
    strNombini_aux = Val(strMaquina) + 1
    strNombfin = Len(ctlCboCajeroRemesa.Text) - Val(strNombini_aux)
    
    LblNomCajeroRemesa.Caption = Mid(Trim(ctlCboCajeroRemesa.Text), strNombini, strNombfin)
    
    'Aca le pasa al arreglo'
    objLiquidacion.CargaConceptoRemesa
    grdRemesa.Array1 = objLiquidacion.Remesa
    
    LblLiquidacion.Caption = "Liquidacion Nº" & "  " & ctlCboCajeroRemesa.BoundText2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set odynR1 = Nothing
End Sub

Private Sub grdRemesa_DblClick()
On Error GoTo handle
    If objLiquidacion.Remesa.UpperBound(1) = -1 Then Exit Sub
    pColR = grdRemesa.Col
    
    'Soles'
    If (pColR = 5) And (objLiquidacion.Remesa(grdRemesa.row, 7) = "1") Then
        frm_VTA_Montos_Remesa.Show vbModal
    End If
    
    'Dolares'
    If (pColR = 6) And (objLiquidacion.Remesa(grdRemesa.row, 7) = "2") Then
        frm_VTA_Montos_Remesa.Show vbModal
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            grdRemesa.SetFocus
        Case vbKeyF2
            CmdGrabaRemesa_Click
        Case vbKeyReturn
            grdRemesa_DblClick
        Case vbKeyF3
            Unload Me
    End Select
End Sub

Private Sub CmdGrabaRemesa_Click()
Dim gvarError As String
    On Error GoTo CtrlErr
    
    gvarError = objRemesa.Graba(gclsOracle.ODataBase, _
                                right(Trim(ctlCboCajeroRemesa.BoundText), 5), _
                                objUsuario.Codigo, _
                                LblCajaRemesa.Caption, _
                                ctlCboCajeroRemesa.BoundText2)
     If gvarError = "" Then
        
        objRemesa.Impresion objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, ctlCboCajeroRemesa.BoundText2

        Nuevo
        psub_IniciaArray
    Else
        MsgBox gvarError, vbCritical, Caption
    End If
    Exit Sub
CtrlErr:
    'Err.Raise Err.Number, "Error al grabar remesa", Err.Description
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Public Sub psub_IniciaArray()
On Error GoTo handle
    Set objRemesa = Nothing
    Set objRemesa = New clsRemesa
    objLiquidacion.Remesa.ReDim 0, -1, 0, 11
    grdRemesa.Array1 = objLiquidacion.Remesa
    grdRemesa.Rebind

    Exit Sub
handle:
    Set objRemesa = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Sub SeteaGrilla()
On Error GoTo handle
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant

    arrCampos = Array("", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Cod", "Concepto", "BBVA", "SCOTIABANK", "BCP", "S/.", "$", "TipoCta", "FPP", "FPH")
    arrAncho = Array(600, 1800, 800, 1100, 800, 900, 900, 900, 900, 900)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgLeft, dbgLeft)
    arrFoco = Array(False, False, False, False, False, True, True, False, False, False)
    grdRemesa.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    
    grdRemesa.Columns(0).Visible = False
    grdRemesa.Columns(7).Visible = False
    grdRemesa.Columns(8).Visible = False
    grdRemesa.Columns(9).Visible = False
    
    grdRemesa.ColumnFooter = True
    
    grdRemesa.Columns(5).FooterText = 0
    grdRemesa.Columns(6).FooterText = 0

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub grdRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 5, 6
            grdRemesa_DblClick
    End Select
End Sub

Private Sub grdRemesa_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo handle
    If grdRemesa.ApproxCount <= 0 Then Exit Sub
    If objLiquidacion.Remesa.UpperBound(1) = -1 Then Exit Sub

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub
