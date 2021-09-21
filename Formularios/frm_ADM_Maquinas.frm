VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ADM_Maquinas 
   Caption         =   "Mantenimiento de Maquinas"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "frm_ADM_Maquinas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   7095
   Begin vbp_Ventas.ctlGrillaArray grdMaquinas 
      Height          =   4575
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8070
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Maquina"
      Height          =   1190
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   3375
      Begin vbp_Ventas.ctlDataCombo ctlCboMaquina 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F1"
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   225
      End
   End
   Begin VB.Frame FraMaq 
      Caption         =   "Registre Maquina"
      Height          =   1190
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3450
      Begin VB.CommandButton CmdAdd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2880
         Picture         =   "frm_ADM_Maquinas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   480
      End
      Begin vbp_Ventas.ctlTextBox TxtIP 
         Height          =   330
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin vbp_Ventas.ctlDataCombo ctlTipoMaquina 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Maquina"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   280
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "IP :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   790
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Maquina"
            Key             =   "Machine"
            Object.ToolTipText     =   "Agregar Maquina"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Graba"
            Key             =   "Save"
            Object.ToolTipText     =   "Agrega Documentos"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Exit"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Command1"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   2640
         Width           =   255
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Maquinas.frx":0894
            Key             =   "Machine"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Maquinas.frx":0E2E
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Maquinas.frx":13C8
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Maquinas.frx":1962
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Maquinas.frx":1EFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      BackColor       =   &H00CFF9FE&
      Caption         =   "Nota:  Hacer doble click sobre el correlativo para cambiarlo"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   6600
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CFF9FE&
      Caption         =   "Hacer DblClik sobre el correlativo para cambiarlo"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frm_ADM_Maquinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMaquina As New clsMaquina
Dim odynMaquina As oraDynaset
Public strDatoNum As String
Dim odynR1 As oraDynaset

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub ctlCboMaquina_Click(Area As Integer)
On Error GoTo Control
If Area = 0 Then Exit Sub
    psubIniArray
    Set odynMaquina = objMaquina.DocumentoxLocal(objUsuario.CodigoEmpresa, _
                                                 objUsuario.CodigoLocal, _
                                                 ctlCboMaquina.BoundText)
    
    If odynMaquina.RecordCount <= 0 Then MsgBox "El Local no registra documentos para emitir", vbCritical, App.ProductName:  Exit Sub
    
    
    odynMaquina.MoveFirst
    While Not odynMaquina.EOF
        objVenta.AgregaMaquina odynMaquina("COD_TIPO_DOCUMENTO").Value, _
                               "" & odynMaquina("COD_MAQUINA").Value, _
                               "" & odynMaquina("NUM_ACTUAL").Value, _
                               "" & odynMaquina("FLG_COLA_IMPRESION").Value, _
                               "" & odynMaquina("COD_MAQUINA_REL").Value, _
                               "" & odynMaquina("COD_SERIE_REL").Value, _
                               "" & odynMaquina("COD_FORMATO").Value
                               
        odynMaquina.MoveNext
    Wend
    grdMaquinas.Array1 = objVenta.Maquina
    grdMaquinas.Rebind
    
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub ctlTipoMaquina_Change()
    If ctlCboMaquina.BoundText = "" Then MsgBox "Seleccione la Maquina", vbCritical, App.ProductName: Exit Sub
    'TxtIP.Text = objMaquina.NumIP
'''    objMaquina.fnIP_x_Maquina(objUsuario.CodigoEmpresa, _
'''                                           objUsuario.CodigoLocal, _
'''                                           ctlCboMaquina.BoundText, _
'''                                           ctlTipoMaquina.BoundText)
                                               
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1: grdMaquinas.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    setteaFormulario Me
    psubIniArray
    SeteaGrilla
    
    '-- Lista Tipo Maquina --'
    Set ctlTipoMaquina.RowSource = objMaquina.ListaTipoMaquina
    ctlTipoMaquina.ListField = "DES"
    ctlTipoMaquina.BoundColumn = "COD"

    psubCargaCboMaq
End Sub

Sub psubCargaCboMaq()
    '-- Lista Maquinas --'
    Set ctlCboMaquina.RowSource = objMaquina.MaquinaLocal(objUsuario.CodigoEmpresa, _
                                                          objUsuario.CodigoLocal)
    ctlCboMaquina.ListField = "DES"
    ctlCboMaquina.BoundColumn = "COD"
End Sub

Public Sub psubIniArray()
    '''Set objVenta = Nothing
    objVenta.Maquina.ReDim 0, -1, 0, 7
    grdMaquinas.Array1 = objVenta.Maquina
    grdMaquinas.Rebind
End Sub

Private Sub CmdAdd_Click()
    If ctlTipoMaquina.BoundText = "" Then MsgBox "Seleccione el Tipo de Maquina", vbCritical, App.ProductName: ctlTipoMaquina.SetFocus: Exit Sub
    If TxtIP.Text = "" Then MsgBox "Ingrese el Nº IP de la maquina", vbCritical, App.ProductName: TxtIP.SetFocus: Exit Sub
    GrabaIP
    FraMaq.Visible = False
End Sub

Sub GrabaIP()
    Dim gvarError  As String
    
    Screen.MousePointer = vbHourglass
    gvarError = objMaquina.GrabaIP(objUsuario.CodigoEmpresa, _
                                   objUsuario.CodigoLocal, _
                                   ctlCboMaquina.BoundText, _
                                   ctlTipoMaquina.BoundText, _
                                   Trim(TxtIP.Text))
                                  
    If gvarError = "" Then
        MsgBox "Se grabó el Nº IP de Máquina", vbInformation, App.ProductName
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub grdMaquinas_DblClick()
    If grdMaquinas.ApproxCount <= 0 Then Exit Sub
    With frm_ADM_Correlativo
        .input_TipoDocumento = grdMaquinas.Columns(0).Value
        .input_CodigoMaquina = grdMaquinas.Columns(1).Value
        .input_NumeroDocumento = grdMaquinas.Columns(2).Value
        .input_CodigoMaquinaRel = grdMaquinas.Columns(4).Value
        .input_CodigoFlagEncola = grdMaquinas.Columns(5).Value
        .input_Ticketera = grdMaquinas.Columns(6).Value
        .input_CodFormato = grdMaquinas.Columns(7).Value
        .Show vbModal
    End With
End Sub

Private Sub grdMaquinas_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn: grdMaquinas_DblClick
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
        Case "Machine"
            MsgBox "Esta opción se encuentre Temporalmente deshabilitada"
            Exit Sub
            'FraMaq.Visible = True
        Case "Save"
            Graba
        Case "Exit"
            Unload Me
    End Select
End Sub

Private Sub TxtIP_KeyPress(KeyAscii As Integer)
    'TxtIP.Tipo = Real
End Sub

Sub Graba()
    Screen.MousePointer = vbHourglass
    Dim i%
    Dim strCadCodTipoDoc As String
    Dim strCadNumActual As String
    Dim strCadMaquinaRel As String
    Dim strCadFlag As String
    Dim strCadTicketera As String
    Dim strCadCodFormato As String
    Dim gvarError  As String

    If objVenta.Maquina.UpperBound(1) = -1 Then MsgBox "No existe datos para grabar", vbCritical, App.ProductName: Exit Sub
    
    If MsgBox("¿Desea grabar los datos actualizados?", vbQuestion + vbYesNo + vbDefaultButton2, "Grabar") = vbNo Then
        Exit Sub
    End If
    
    
    For i = 0 To objVenta.Maquina.UpperBound(1)
        If Not IsNull(objVenta.Maquina(i, 2)) And Not IsEmpty(objVenta.Maquina(i, 2)) And objVenta.Maquina(i, 2) <> vbNullString Then
            gvarError = "" & objMaquina.fn_Verifica_Sec_Correlativo(objUsuario.CodigoEmpresa, _
                                                                    objUsuario.CodigoLocal, _
                                                                    objVenta.Maquina(i, 0), _
                                                                    objVenta.Maquina(i, 2))
            If LenB(Trim(gvarError)) > 0 Then
                MsgBox gvarError, vbOKOnly + vbInformation, App.ProductName
                gvarError = vbNullString
            End If
        End If
        
        strCadCodTipoDoc = strCadCodTipoDoc & objVenta.Maquina(i, 0) & "|"
        strCadNumActual = strCadNumActual & objVenta.Maquina(i, 2) & "|"
        strCadMaquinaRel = strCadMaquinaRel & objVenta.Maquina(i, 4) & "|"
        strCadFlag = strCadFlag & objVenta.Maquina(i, 5) & "|"
        strCadTicketera = strCadTicketera & objVenta.Maquina(i, 6) & "|"
        strCadCodFormato = strCadCodFormato & objVenta.Maquina(i, 7) & "|"
    Next i
    
    gvarError = objMaquina.GrabaCorXDoc(objUsuario.CodigoEmpresa, _
                                        objUsuario.CodigoLocal, _
                                        ctlCboMaquina.BoundText, _
                                        strCadCodTipoDoc, _
                                        strCadNumActual, _
                                        objUsuario.Codigo, _
                                        strCadMaquinaRel, _
                                        strCadFlag, _
                                        strCadTicketera, _
                                        strCadCodFormato)
                                  
    If gvarError = "" Then
        MsgBox "Se grabó los correlativos", vbInformation, App.ProductName
        'grdMaquinas.Delete
        'psubIniArray
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Sub SeteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("", "", "", "", "", "", "", "")
    arrCaption = Array("Tipo Documento", "Maquina", "Nº Actual", "M.Alternativa", "Encola", "Flag", "Ticketera", "Formato")
    arrAncho = Array(450, 1200, 1000, 2500, 1000, 500, 1300, 800)
    arrAlineacion = Array(dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft)
    
    grdMaquinas.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdMaquinas.Columns(1).Visible = False
    grdMaquinas.Columns(5).Visible = False
        
    Dim i%
    For i = 0 To grdMaquinas.Columns.Count - 1
        grdMaquinas.Columns(i).AllowFocus = False
    Next i
        grdMaquinas.Columns(2).AllowFocus = True
        grdMaquinas.Columns(6).AllowFocus = True
        
        
    Dim objFormato As New clsFormato
    Dim odyTmp As oraDynaset
    Set odyTmp = objFormato.Lista
    Set objFormato = Nothing
    
    While Not odyTmp.EOF
        psub_Grilla_Traslate grdMaquinas, 7, odyTmp("COD_FORMATO").Value, odyTmp("DES_FORMATO").Value
        odyTmp.MoveNext
    Wend
    
End Sub
