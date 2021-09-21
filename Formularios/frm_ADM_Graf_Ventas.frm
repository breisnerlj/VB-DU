VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_ADM_Graf_Ventas 
   AutoRedraw      =   -1  'True
   Caption         =   "Estadisticas de ventas "
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   Icon            =   "frm_ADM_Graf_Ventas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart graVentas 
      Height          =   5175
      Left            =   0
      OleObjectBlob   =   "frm_ADM_Graf_Ventas.frx":000C
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frm_ADM_Graf_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrDatos As Variant
Private strTitulo As String

Public Property Get Titulo() As String
    Titulo = strTitulo
End Property

Public Property Let Titulo(ByVal vNewValue As String)
    strTitulo = vNewValue
End Property

Public Property Get Datos() As Variant
    Datos = arrDatos
End Property

Public Property Let Datos(ByVal vNewValue As Variant)
    arrDatos = vNewValue
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub


Private Sub Form_Resize()
    graVentas.Width = Me.Width - 90 '- 3 * graVentas.Left
    graVentas.Height = Me.Height - 400 '- (4 * graVentas.Top) '* graVentas.Left)
End Sub

Public Sub Mostrar()
    Unload Me
        
    Me.Width = 8000
    Me.Height = 5000
    
    sub_Dibuja
    
    Me.Caption = Titulo
    
    Me.Show vbModal
End Sub

Private Sub sub_Dibuja()
Dim i%, j%, k%
Dim xdbGrafica As New XArrayDB
        
    
    Dim dblCtdVendida#, dblPNC_Stock#, dblPNC_Venta#, dblImpVenta#
    Dim strPeriodo$, varDaTos As Variant, strDatos$
    Dim strPeriodoAct$
                                                                             
    Screen.MousePointer = vbHourglass
                
    graVentas.Visible = True
            
    Screen.MousePointer = vbHourglass
        
    
    With graVentas
        .chartType = VtChChartType2dBar
        .ChartData = Datos
        .AllowDithering = True
        
        For i = LBound(Datos, 1) + 1 To UBound(Datos, 1)
                .Row = i
            For j = LBound(Datos, 2) + 1 To UBound(Datos, 2)
                
                .Column = j
                .ColumnLabelCount = 2

                .ColumnLabelIndex = 2
                .ColumnLabel = Format(Val(Datos(i, j)), "Standard")
                
            Next j
        Next i
                      
        .Legend.Location.LocationType = VtChLocationTypeBottom
        .ShowLegend = True
    End With
    Screen.MousePointer = vbDefault
End Sub

