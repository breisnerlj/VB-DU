VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_OFF_Documentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Correlativos"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frm_OFF_Documentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboDocDefecto 
      Height          =   315
      ItemData        =   "frm_OFF_Documentos.frx":1C9A2
      Left            =   120
      List            =   "frm_OFF_Documentos.frx":1C9A4
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4245
      Width           =   2535
   End
   Begin TrueOleDBGrid70.TDBGrid grdDocumentos 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4895
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "C. Doc"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Documento"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Número Actual"
      Columns(2).DataField=   ""
      Columns(2).DataWidth=   10
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1535"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3836"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3757"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2990"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2910"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.locked=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "&Continuar >>"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H80000018&
      Caption         =   "En Número Actual, debe ingresar el siguiente correlativo al ultimo documento emitido."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   5295
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
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Documento por defecto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese el número actual de los documentos a emitir:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4680
   End
End
Attribute VB_Name = "frm_OFF_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objArchivoIni As New cls_ArchivoIni
Dim tmpDocumento As New XArrayDB
Dim arrSeries As New XArrayDB

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdContinuar_Click()
Dim objDocumento As cls_OFF_Documento
    
    If Not ValidarCorrelativos Then
        Exit Sub
    End If
    If Not ValidarDocumento Then
        cboDocDefecto.SetFocus
        Exit Sub
    End If
    ActualizarCorrelativos
    
    objArchivoIni.GuardarIni gstrIni, "general", "FLG_CONTINGENCIA", "1"
    
    objOFFVenta.CodDocDefault = tmpDocumento.Value(cboDocDefecto.ListIndex, 0)
    Set objDocumento = New cls_OFF_Documento
    objOFFVenta.NumDocDefault = objDocumento.UltimoCorrelativo(objOFFVenta.CodDocDefault)
    Set objDocumento = Nothing
    objArchivoIni.GuardarIni gstrIni, "general", "COD_DOC_DEFAULT", tmpDocumento.Value(cboDocDefecto.ListIndex, 0)
    
    Unload Me
    
    frm_OFF_Principal.Show

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        grdDocumentos.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    Dim objDocumento As cls_OFF_Documento
    Dim I As Integer
    objArchivoIni.GuardarIni gstrIni, "general", "FLG_CONTINGENCIA", "0"
    Set objDocumento = New cls_OFF_Documento
    Set tmpDocumento = objDocumento.ListaTipoDocumento
    Set objDocumento = Nothing
    Set grdDocumentos.Array = tmpDocumento
    '''Set arrSeries = tmpDocumento
    arrSeries.ReDim 0, -1, 0, tmpDocumento.Count(2)
    For I = tmpDocumento.LowerBound(1) To tmpDocumento.UpperBound(1)
        arrSeries.AppendRows 1
        arrSeries.Value(I, 0) = tmpDocumento(I, 0)
        arrSeries.Value(I, 1) = tmpDocumento(I, 1)
        arrSeries.Value(I, 2) = tmpDocumento(I, 2)
        arrSeries.Value(I, 3) = tmpDocumento(I, 3)
        arrSeries.Value(I, 4) = tmpDocumento(I, 4)
    Next
    CargaCombo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objArchivoIni = Nothing
    Set tmpDocumento = Nothing
End Sub

Private Sub grdDocumentos_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    
    If ColIndex = 2 Then
        If Len(grdDocumentos.Columns(2).Text) > 0 And Len(grdDocumentos.Columns(2).Text) <> 10 Then
            MsgBox "El correlativo debe tener 10 caracteres de longitud", vbExclamation + vbOKOnly, App.ProductName
            Cancel = True
        End If
    End If
    
End Sub

'Private Sub grdDocumentos_Change()
'    objDocumento.ActualizaCorrelativo grdDocumentos.Columns(0), grdDocumentos.Columns(2)
'End Sub

Private Sub grdDocumentos_KeyPress(KeyAscii As Integer)

    ' Don't disable the Esc or Backspace keys
    If (KeyAscii = 27) Or (KeyAscii = 8) Then Exit Sub

    ' Cancel user key input if it is not a digit
    If (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If

End Sub

Private Sub CargaCombo()
    Dim I As Integer
    Dim de As Integer
    I = 0
    While I <= tmpDocumento.UpperBound(1)
        cboDocDefecto.AddItem tmpDocumento(I, 1), I
        If objArchivoIni.LeerIni(gstrIni, "general", "COD_DOC_DEFAULT", "") = tmpDocumento(I, 0) Then
            de = I
        End If
        I = I + 1
    Wend
    cboDocDefecto.ListIndex = de
End Sub

Private Function ValidarCorrelativos() As Boolean
On Error GoTo Handle

    Dim I As Integer
    Dim strCorrelativo As String
    Dim strDocumento As String
    Dim a As Boolean
    
    ValidarCorrelativos = True

    For I = tmpDocumento.LowerBound(1) To tmpDocumento.UpperBound(1)
        strDocumento = Trim(tmpDocumento(I, 1))
        strCorrelativo = Trim(tmpDocumento(I, 2))
        If Len(strCorrelativo) > 0 And Len(strCorrelativo) <> 10 Then
            MsgBox "El correlativo del documento " & strDocumento & " debe tener 10 caracteres de longitud", vbExclamation + vbOKOnly, App.ProductName
            ValidarCorrelativos = False
            Exit For
        End If
        If strCorrelativo <> "" Then
            If Not arrSeries(I, 2) = "" Then
                If Mid(arrSeries(I, 2), 1, 4) <> Mid(strCorrelativo, 1, 4) Then
                    MsgBox "El correlativo del documento " & strDocumento & " no coincide con la Serie", vbExclamation + vbOKOnly, App.ProductName
                    ValidarCorrelativos = False
                    Exit For
                End If
            End If
        End If

    Next I

    Exit Function
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Function

Public Sub ActualizarCorrelativos()
Dim objDocumento As cls_OFF_Documento
On Error GoTo Handle

    Dim I As Integer
    Dim strCorrelativo As String
    Dim strDocumento As String
    'objArchivoIni.LeerIni (gstrIni, "general", "SER_DOCUMENTO", "")
    Set objDocumento = New cls_OFF_Documento
    
    For I = tmpDocumento.LowerBound(1) To tmpDocumento.UpperBound(1)
        strDocumento = Trim(tmpDocumento(I, 0))
        strCorrelativo = Trim(tmpDocumento(I, 2))
        
        If Len(strDocumento) = 3 Then
            objDocumento.ActualizaCorrelativo strDocumento, strCorrelativo, True
        End If
    Next I
    Set objDocumento = Nothing
    

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdDocumentos_LostFocus()

    grdDocumentos.Update

End Sub

Private Function ValidarDocumento() As Boolean
On Error GoTo Handle

    Dim I As Integer
    Dim strCorrelativo As String
    Dim strDocumento As String
    
    ValidarDocumento = True
    
    For I = tmpDocumento.LowerBound(1) To tmpDocumento.UpperBound(1)
        strDocumento = Trim(tmpDocumento(I, 1))
        strCorrelativo = Trim(tmpDocumento(I, 2))
        
        If strDocumento = Trim(cboDocDefecto.Text) And Len(strCorrelativo) < 1 Then
            MsgBox strDocumento & " no puede ser el Documento por Defecto porque no tiene correlativo", vbExclamation + vbOKOnly, App.ProductName
            ValidarDocumento = False
            Exit For
        End If
    Next I

    Exit Function
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Function

