VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form frm_CNT_Estadistica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadistica de Control de Competencias"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6015
      Left            =   120
      OleObjectBlob   =   "frm_CNT_Estadistica.frx":0000
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10215
      Begin MSComCtl2.DTPicker dtpPeriodo 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyyMM"
         Format          =   16449539
         CurrentDate     =   40007
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_CNT_Estadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objControl As New clsCntCompetencia
Public NumeroControl As String

Private Sub dtpPeriodo_Change()
    BuscaDatos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
   On Error GoTo Control

dtpPeriodo.Value = objUsuario.sysdate

BuscaDatos

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Sub BuscaDatos()
Dim rsControl As oraDynaset
    Set rsControl = objControl.ListaEstadistica("", Format(dtpPeriodo.Value, "YYYYmm"))
    
   With MSChart1
      ' Muestra un gráfico 3d con 8 columnas y 8 filas
      ' de datos.
      .chartType = VtChChartType2dLine
      Dim h As Integer
      .ColumnLabelCount = rsControl.RecordCount
      .ColumnCount = rsControl.RecordCount
      .RowCount = 7
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        While Not rsControl.EOF
                '.ColumnLabelIndex = H + 1
                .Column = h + 1
                .ColumnLabel = CStr("" & rsControl("NOM_PROVEEDOR").Value)
                '.Legend = CStr("" & rsControl("NOM_PROVEEDOR").Value)
                Debug.Print .ColumnLabel
                    For i = 1 To 7
                      .row = i
                      .RowLabel = rsControl.FieldName(i)
                      .data = "" & rsControl(i).Value
                    Next
                    
                h = h + 1
          rsControl.MoveNext
        Wend
        

      
''      For i = 1 To 7
''        .row = i
''        .RowLabel = rsControl.FieldName(i)
''        While Not rsControl.EOF
''                .ColumnLabelIndex = H + 1
''                .data = "" & rsControl(i).Value
''                .ColumnLabel = CStr("" & rsControl("NOM_PROVEEDOR").Value)
''                H = H + 1
''          rsControl.MoveNext
''        Wend
''      Next
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
   End With

End Sub


