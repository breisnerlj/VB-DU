VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_VTA_FormsSoat 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   Icon            =   "frm_VTA_FormsSoat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_FormsSoat.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1095
   End
   Begin TrueDBGrid70.TDBGrid grdSoat 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7435
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "#"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Inicio"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fin"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=4101"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4022"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=4128"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4048"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=67,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&HDDE4FF&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FormsSoat.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin vbp_Ventas.ctlTextBox TxtNumInicio 
         Height          =   375
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Alignment       =   2
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
      Begin vbp_Ventas.ctlTextBox TxtNumFin 
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Alignment       =   2
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
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
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
      Index           =   11
      Left            =   6097
      TabIndex        =   10
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
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
      Index           =   12
      Left            =   4380
      TabIndex        =   8
      Top             =   6900
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_FormsSoat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Dim xdbArray As New XArrayDB
Dim strCadNumIni$
Dim strCadNumFin$
Dim strCodSoat As String
Dim strCtdFrac As String

Private Sub cmdAceptar_Click()
    
    If xdbArray.UpperBound(1) = "-1" Then MsgBox "No existe certificados SOAT a registar", vbCritical, App.ProductName: Exit Sub
                                  
    Screen.MousePointer = vbHourglass
                                  
    Dim i%
    strCadNumIni = "": strCadNumFin = ""
    For i = 0 To xdbArray.UpperBound(1)
        strCadNumIni = strCadNumIni & xdbArray(i, 1) & "|"
        strCadNumFin = strCadNumFin & xdbArray(i, 2) & "|"
    Next i
    Dim gvarError  As String
    
    gvarError = objServicio.GrabaAsignaSoat(objUsuario.CodigoLocal, _
                                  strCadNumIni, _
                                  strCadNumFin, _
                                  objUsuario.Codigo)
                                  
    If gvarError = "" Then
        MsgBox "Se Generó con satisfación los certificados", vbInformation, App.ProductName
        TxtNumInicio.Text = "": TxtNumFin.Text = "": grdSoat.Delete
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call Xarray(xdbArray)
    Me.top = 0
    Me.left = 0
    setteaFormulario Me
    SeteaGrilla

End Sub

Private Sub Xarray(ByVal vxdArray As XArrayDB)
    vxdArray.ReDim 0, -1, 0, 3
    grdSoat.Array = vxdArray
    grdSoat.Rebind
End Sub

Private Sub CmdAdd_Click()
         
    '-- 04/09/2008 Atributos del SOAT ---'
    strCodSoat = objServicio.Codigo_SOAT
    strCtdFrac = objServicio.Ctd_Fraccion
    
    If (Trim(TxtNumInicio.Text) = "") Or (Trim(TxtNumFin.Text) = "") Then MsgBox "Coloque Rango de Certificados para la Asignación", vbCritical, "CERTIFICADOS SOAT": TxtNumInicio.selection: Exit Sub
    On Error GoTo CtrlErr
    
    If Abs(Trim(TxtNumInicio.Text) - Trim(TxtNumFin.Text)) > Val(strCtdFrac) Then MsgBox "El rango del correlativo no puede superar la cantidad" & Chr(13) & _
                                                                                         "de la que tiene el Pack de formularios que son" & " " & strCtdFrac, vbCritical, "Asignación de formularios SOAT": Exit Sub
    Add_Secuencia TxtNumInicio.Text, _
                  TxtNumFin.Text, _
                  xdbArray
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "", App.FileDescription
End Sub

Private Sub sub_Items(ByVal vxdb As XArrayDB)
    Dim i%
    For i = 0 To vxdb.UpperBound(1)
        vxdb(i, 0) = i + 1
    Next i
End Sub

Private Sub Add_Secuencia(ByVal vstrNumIni As String, _
                          ByVal vstrNumFin As String, _
                          ByVal vxdbArray As XArrayDB)

    Dim aux As Integer
    Dim i As Integer
    Dim encontro As Boolean
    
    aux = vxdbArray.Count(1)
    While i < aux
       If (vxdbArray(i, 1) = vstrNumIni) Or (vxdbArray(i, 2) = vstrNumFin) Then
            encontro = True
         Else
            encontro = False
       End If
       i = i + 1
    Wend
    
    If encontro = False Then
      vxdbArray.AppendRows 1
      Call sub_Items(vxdbArray)
      vxdbArray(vxdbArray.UpperBound(1), 1) = vstrNumIni
      vxdbArray(vxdbArray.UpperBound(1), 2) = vstrNumFin
    End If
    
    grdSoat.Rebind
End Sub


Private Sub grdSoat_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
          On Error GoTo CtrlErr
            grdSoat.Delete
CtrlErr:
    On Error GoTo 0
            
    End Select
End Sub

Private Sub TxtNumFin_KeyPress(KeyAscii As Integer)
    TxtNumFin.Tipo = Entero
End Sub

Private Sub TxtNumInicio_KeyPress(KeyAscii As Integer)
    TxtNumInicio.Tipo = Entero
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub SeteaGrilla()
    grdSoat.AllowUpdate = False
    grdSoat.AllowRowSizing = False
    grdSoat.Columns(0).Alignment = dbgLeft
    
    Dim i%
    For i = 0 To grdSoat.Columns.Count - 1
        grdSoat.Columns(i).AllowSizing = False
    Next i
    grdSoat.Style.VerticalAlignment = dbgVertCenter
    grdSoat.AllowColSelect = False
    grdSoat.AllowRowSelect = True
    grdSoat.AllowRowSizing = False
    
    grdSoat.MarqueeStyle = dbgHighlightRow
    grdSoat.HighlightRowStyle.BackColor = RGB(255, 230, 230)
    grdSoat.HighlightRowStyle.ForeColor = RGB(0, 0, 0)
    grdSoat.HighlightRowStyle.Font.Bold = True
    grdSoat.HighlightRowStyle.Font.Size = 9
    grdSoat.RowHeight = 1.5 * grdSoat.RowHeight
    grdSoat.Font.Size = 10
End Sub

