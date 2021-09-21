VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_Lista_Caja_PreCerradas 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   Icon            =   "frm_Lista_Caja_PreCerradas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Info. Liqui."
      Height          =   555
      Left            =   1920
      Picture         =   "frm_Lista_Caja_PreCerradas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Detalle de Liquidación"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdImpDetLiquidacion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imp.Detalle"
      Height          =   555
      Left            =   480
      Picture         =   "frm_Lista_Caja_PreCerradas.frx":064D
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Imprime Detalle de Liquidación"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdDetLiquidacion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle Liqui."
      Height          =   555
      Left            =   3300
      Picture         =   "frm_Lista_Caja_PreCerradas.frx":0BD7
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Detalle de Liquidación"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdImpresion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Re Impresión"
      Height          =   555
      Left            =   4620
      Picture         =   "frm_Lista_Caja_PreCerradas.frx":1161
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Re Impresión de Liquidación"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnular 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Activa Caja"
      Height          =   555
      Left            =   5940
      Picture         =   "frm_Lista_Caja_PreCerradas.frx":16EB
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Activa Caja"
      Top             =   3240
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1111
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            Key             =   "Close"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar"
            Key             =   "Buscar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Anular"
            Key             =   "Anular"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   2880
         TabIndex        =   13
         Top             =   -120
         Width           =   3975
         Begin MSComCtl2.DTPicker dtpFchIni 
            Height          =   375
            Left            =   600
            TabIndex        =   14
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   67698689
            CurrentDate     =   39013
         End
         Begin MSComCtl2.DTPicker dtpFchFin 
            Height          =   375
            Left            =   2520
            TabIndex        =   15
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   67698689
            CurrentDate     =   39013
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Fin"
            Height          =   435
            Left            =   2040
            TabIndex        =   17
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Inicio"
            Height          =   435
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Seleccionados"
      Height          =   1280
      Left            =   0
      TabIndex        =   1
      Top             =   5950
      Width           =   7215
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
         Left            =   5640
         TabIndex        =   24
         Top             =   885
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label LblEstCaja 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5640
         TabIndex        =   11
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label LblFecApertura 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5640
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LblNomCja 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   9
         Top             =   892
         Width           =   3375
      End
      Begin VB.Label LblNomQF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   8
         Top             =   563
         Width           =   3375
      End
      Begin VB.Label LblCaja 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   7
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Estado Caja"
         Height          =   195
         Left            =   4320
         TabIndex        =   6
         Top             =   623
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fec. Apertura"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   293
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cajero"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   945
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quimico"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   616
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Caja"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   308
         Width           =   315
      End
   End
   Begin vbp_Ventas.ctlGrilla grdCajas 
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3969
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   5520
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":1C75
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":220F
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":27A9
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":2D43
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":32DD
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":3877
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":3E11
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":43AB
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":4945
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":4EDF
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":5479
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Lista_Caja_PreCerradas.frx":5A13
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin vbp_Ventas.ctlGrilla grdArqAnula 
      Height          =   2055
      Left            =   0
      TabIndex        =   18
      Top             =   3840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3625
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "frm_Lista_Caja_PreCerradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objImpresion As New clsImpresiones
Dim odynLiq As oraDynaset
Dim strValor As String

Private Sub CmdDetLiquidacion_Click()
    'x = grdCajas.Columns("COD_LIQUIDACION").Value
    'x = grdArqAnula.Columns("COD_LIQUIDACION").Value
    On Error GoTo CtrlErr
    
    If (grdCajas.ApproxCount <= 0 And grdArqAnula.ApproxCount <= 0) Then Exit Sub
    frm_VTA_Detalle_Liquidacion.pCodLiq = Trim(LblLiquidacion.Caption)
    '21/04/09 comentado por PHERRERA No es necesario volver a cargar la informacion, si no hay venta, que se muestre en blanco.
'    Set odynLiq = objLiquidacion.fnCalFormaPagos(objUsuario.CodigoEmpresa, _
'                                                 objUsuario.CodigoLocal, _
'                                                 grdCajas.Columns("COD_MAQUINA").Value, _
'                                                 strValor, _
'                                                 grdCajas.Columns("COD_LIQUIDACION").Value)
    
'    If odynLiq.RecordCount <= 0 Then MsgBox "La liquidación no tiene venta", vbCritical, Caption: Exit Sub
    frm_VTA_Detalle_Liquidacion.Show vbModal
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub cmdImpDetLiquidacion_Click()
    On Error GoTo CtrlErr
    
    Dim lstrCodMaquina As String
    Dim lstrLiquidacion As String
    
    If (grdCajas.ApproxCount <= 0 And grdArqAnula.ApproxCount <= 0) Then Exit Sub
    lstrLiquidacion = Trim(LblLiquidacion.Caption)
    
    If lstrLiquidacion = "" Then
        Exit Sub
    End If
    
    If grdCajas.ApproxCount > 0 Then
        If grdCajas.Columns("COD_LIQUIDACION").Value = lstrLiquidacion Then
            lstrCodMaquina = grdCajas.Columns("COD_MAQUINA").Value
        End If
    End If
    If grdArqAnula.ApproxCount > 0 Then
        If grdArqAnula.Columns("COD_LIQUIDACION").Value = lstrLiquidacion Then
            lstrCodMaquina = grdArqAnula.Columns("COD_MAQUINA").Value
        End If
    End If
    
    ' invocar el proceso que imprime el detalle de la liquidacion
    If MsgBox("Esta seguro de Imprimir la Liquidación Nº" & "  " & lstrLiquidacion, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    objImpresion.Imprime_Detalle_Liquidacion objUsuario.CodigoEmpresa, _
                                             objUsuario.CodigoLocal, _
                                             lstrCodMaquina, _
                                             lstrLiquidacion
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub cmdInfo_Click()
    If (grdCajas.ApproxCount <= 0 And grdArqAnula.ApproxCount <= 0) Then Exit Sub
    frm_ADM_Liquidacion.pCodLiq = Trim(LblLiquidacion.Caption)
    frm_ADM_Liquidacion.Show
End Sub

Private Sub Form_Load()
    setteaFormulario Me
    SeteaGrilla
    SeteaGrilla_Anula
    dtpFchIni.Value = "01/" & Month(Now) & "/" & Year(Now) 'Format(objUsuario.sysdate, "dd/mm/yyyy")
    dtpFchFin.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    LblNomCja.Caption = objLiquidacion.fnDevNomUsu(objUsuario.Codigo)
    
    strValor = objUsuario.Parametros("COD_MONEDA")(0, 2)
End Sub

Private Sub CmdAnular_Click()
    If grdArqAnula.ApproxCount <= 0 Then Exit Sub
    Dim CodLiqui As String
    CodLiqui = grdArqAnula.Columns("COD_LIQUIDACION").Value
    If MsgBox("Esta seguro de anular la liquidación Nº " & Trim(CodLiqui) & " ", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        Anula CodLiqui
        Set grdCajas.DataSource = objLiquidacion.ListaCajasPrecerradas(objUsuario.CodigoEmpresa, _
                                                                           objUsuario.CodigoLocal, _
                                                                           Format(dtpFchIni.Value, "dd/mm/yyyy"))
                                                                           
        Set grdArqAnula.DataSource = objLiquidacion.fndevLiqCerrados(objUsuario.CodigoEmpresa, _
                                                                         objUsuario.CodigoLocal, _
                                                                         Format(dtpFchIni.Value, "dd/mm/yyyy"))
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
                grdCajas.SetFocus
        Case vbKeyF2
                grdArqAnula.SetFocus
        Case vbKeyEscape
                Unload Me
    End Select
End Sub

Private Sub cmdImpresion_Click()
    If grdArqAnula.ApproxCount <= 0 Then Exit Sub
      'MsgBox "Se Re Imprimira la Liquidación Nª " & grdArqAnula.Columns("COD_LIQUIDACION").Value & "", vbInformation, Caption
    If MsgBox("Esta seguro de Imprimir la Liquidación Nº" & "  " & grdArqAnula.Columns("COD_LIQUIDACION").Value, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        objImpresion.Imprime_Liquidacion objUsuario.CodigoEmpresa, _
                                         objUsuario.CodigoLocal, _
                                         grdArqAnula.Columns("COD_MAQUINA").Value, _
                                         grdArqAnula.Columns("COD_LIQUIDACION").Value
      Else
        Exit Sub
    End If
                                         
End Sub

Private Sub grdArqAnula_DblClick()
    'cmdImpresion_Click
    If grdArqAnula.ApproxCount <= 0 Then Exit Sub
    frm_VTA_LiquidacionCaja.Mostrar True, grdArqAnula.Columns("COD_MAQUINA").Value, grdArqAnula.Columns("COD_LIQUIDACION").Value
End Sub

Private Sub grdArqAnula_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
                grdArqAnula.SetFocus
    End Select
End Sub

Private Sub grdArqAnula_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdArqAnula.ApproxCount <= 0 Then Exit Sub
    LblLiquidacion.Caption = "" & grdArqAnula.Columns("COD_LIQUIDACION").Value
End Sub

Private Sub grdCajas_DblClick()
    If grdCajas.ApproxCount <= 0 Then Exit Sub
    
'''    Set odynLiq = objLiquidacion.fnCalFormaPagos(objUsuario.CodigoEmpresa, _
'''                                                 objUsuario.CodigoLocal, _
'''                                                 grdCajas.Columns("COD_MAQUINA").Value, _
'''                                                 strValor, _
'''                                                 grdCajas.Columns("COD_LIQUIDACION").Value)
    
'''    If odynLiq.RecordCount > 0 Then
'''       'frm_VTA_LiquidacionCaja.Show vbModal
'''       'Set frm_VTA_LiquidacionCaja.podynLiq = odynLiq
       frm_VTA_LiquidacionCaja.Mostrar False, grdCajas.Columns("COD_MAQUINA").Value, grdCajas.Columns("COD_LIQUIDACION").Value
'''     Else
'''       MsgBox "La caja no tiene documentos emitidos", vbCritical, App.ProductName
'''       Exit Sub
'''    End If
End Sub

Private Sub grdCajas_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdCajas_DblClick
    End Select
End Sub

Private Sub grdCajas_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdCajas.ApproxCount <= 0 Then Exit Sub
    LblCaja.Caption = "" & grdCajas.Columns("COD_MAQUINA").Value
    LblNomCja.Caption = "" & grdCajas.Columns("USU_TEC").Value & " " & grdCajas.Columns("NOM_TEC").Value
    LblFecApertura.Caption = "" & grdCajas.Columns("FCH_INICIO").Value
    LblEstCaja.Caption = "" & grdCajas.Columns("FLG_ESTADO_CAJA").Value
    LblLiquidacion.Caption = "" & grdCajas.Columns("COD_LIQUIDACION").Value
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error GoTo Control

    Select Case Button.Key
        Case "Nuevo"
            Nuevo
        Case "Buscar"
            Set grdCajas.DataSource = objLiquidacion.ListaCajasPrecerradas(objUsuario.CodigoEmpresa, _
                                                                           objUsuario.CodigoLocal, _
                                                                           Format(dtpFchIni.Value, "dd/mm/yyyy"))
                                                                           
            Set grdArqAnula.DataSource = objLiquidacion.fndevLiqCerrados(objUsuario.CodigoEmpresa, _
                                                                         objUsuario.CodigoLocal, _
                                                                         Format(dtpFchIni.Value, "dd/mm/yyyy"))

        Case "Anular"
'            If grdArqAnula.ApproxCount <= 0 Then Exit Sub
'            Dim CodLiqui As String
'            CodLiqui = grdArqAnula.Columns("COD_LIQUIDACION").Value
'            If MsgBox("Esta seguro de anular la liquidación Nº " & Trim(CodLiqui) & " ", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
'                Anula CodLiqui
'            End If
             If grdCajas.ApproxCount <= 0 Then Exit Sub
             Dim strError As String
             Dim CodLiqui As String
             CodLiqui = grdCajas.Columns("COD_LIQUIDACION").Value
             strError = objLiquidacion.CierraCajaNoDoc(gclsOracle.ODataBase, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, CodLiqui, objUsuario.Codigo)
              
              If strError = "" Then
                 MsgBox "Se cerro la caja de liquidacion Nº " & Trim(CodLiqui) & "", vbInformation, Caption
               Else
                 MsgBox strError, vbCritical, App.ProductName
              End If
              
        Case "Salir"
            Unload Me
    End Select

  
Exit Sub
Control:

    MsgBox Err.Descripcion, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Sub Anula(ByVal vstrCodLiqui As String)
Dim gvarError As String



  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Autor : Arturo Escate
    'Fecha : 10/11/2009
    'Proposito: Esto es para validar si necesita autorizacion previa
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim ObjValidacion As New clsAprobacion
    Dim strNumeroSolicitud As String
    Dim strAccion As String
    Dim strMensaje As String
    Dim strCodigoAutorizacion As String
    Dim srtCodigoAUTH As String
    Dim strStore As String
    srtCodigoAUTH = ""
valida:
'TxtImpTot.text
    If srtCodigoAUTH = "" Then frm_VTA_ObservaAutorizacion.Show vbModal
    
    strStore = ObjValidacion.Solicita("4", strAccion, strMensaje, srtCodigoAUTH, objUsuario.CodigoLocal, vstrCodLiqui, "", "", "", "", "", "1", "", "", objUsuario.Codigo, frm_VTA_ObservaAutorizacion.OutObservacion, strCodigoAutorizacion, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", frm_VTA_ObservaAutorizacion.OutNumeroId)
    If Not strStore = "" Then
        MsgBox strStore, vbCritical, App.ProductName
        Exit Sub
    Else
        Select Case strAccion
            Case 0
                    MsgBox strMensaje, vbInformation, App.ProductName
            Case 1
                   MsgBox strMensaje, vbCritical, App.ProductName
                   Exit Sub
            Case 2
                   MsgBox strMensaje, vbInformation, App.ProductName
                   Exit Sub
            Case 3
                If MsgBox(strMensaje & Chr(13) & "¿Desea ingresar el codigo de autorización?", vbYesNo + vbInformation, App.ProductName) = vbYes Then
                    srtCodigoAUTH = frmAprobacion.Carga
                    If Not srtCodigoAUTH = "" Then
                        GoTo valida
                        Exit Sub
                    End If
                   Exit Sub
                Else
                    Exit Sub
                End If
            Case Else
                   MsgBox "no esta implementado", vbInformation, App.ProductName
                   Exit Sub
        End Select
    End If

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    










    gvarError = objLiquidacion.Anula(gclsOracle.ODataBase, _
                                     objUsuario.CodigoEmpresa, _
                                     objUsuario.CodigoLocal, _
                                     vstrCodLiqui, "2", _
                                     objUsuario.Codigo _
                                     )
                         
    If gvarError = "" Then
        MsgBox "Se anulo la liquidacion de caja", vbInformation, App.ProductName
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    End If
End Sub

Sub Nuevo()
    grdCajas.Limpiar
    LblCaja.Caption = "": LblEstCaja.Caption = ""
    LblFecApertura.Caption = "": LblNomCja.Caption = ""
    LblNomQF.Caption = ""
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_MAQUINA", "FCH_INICIO", _
                      "USU_TEC", "NOM_TEC", _
                      "FLG_ESTADO_CAJA", "COD_LIQUIDACION")
                      
    arrCaption = Array("Maquina", "Fch Inicio", _
                       "Código", "Nombre Depen", _
                       "Estado", "Liquidación")
    
    arrAncho = Array(900, 1000, _
                     600, 2150, _
                     850, 1450)
    
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft)
    
    grdCajas.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
End Sub

Private Sub SeteaGrilla_Anula()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_MAQUINA", "COD_LIQUIDACION", _
                      "COD_DEPENDIENTE", "NOMBRE", _
                      "FECHA")
                      
    arrCaption = Array("Maquina", "Liquidación", _
                       "Código", "Nombre Depen", _
                       "Fecha")
    
    arrAncho = Array(900, 1600, _
                     600, 2200, _
                     1600)
    
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft)
    
    grdArqAnula.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdArqAnula.Caption = "Liquidaciones Cerradas"
    grdArqAnula.FColorCaption = vbWhite
    grdArqAnula.BColorCaption = vbBlack
End Sub
