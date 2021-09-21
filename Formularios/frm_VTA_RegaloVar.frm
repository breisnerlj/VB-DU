VERSION 5.00
Begin VB.Form frm_VTA_RegaloVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione un producto de regalo"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrillaArray ctlGrillaArray1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5953
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label lblMensjae 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frm_VTA_RegaloVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim xTmp As New XArrayDB
Dim TotalEscoge As Integer

Private Sub ctlGrillaArray1_AfterColUpdate(ByVal ColIndex As Integer)
Dim i As Integer
i = Val(ctlGrillaArray1.Bookmark)

objVenta.xProductoRegalo.Value(i, 2) = ctlGrillaArray1.Columns(2).Value
objVenta.xProductoRegalo.Value(i, 4) = Val(ctlGrillaArray1.Columns(2).Value) * xTmp.Value(i, 25)
xTmp.Value(i, 2) = ctlGrillaArray1.Columns(2).Value
xTmp.Value(i, 4) = Val(ctlGrillaArray1.Columns(2).Value) * xTmp.Value(i, 25)
End Sub

Private Sub ctlGrillaArray1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
     ctlGrillaArray1.MovePrevious
    ctlGrillaArray1.MoveNext

    Dim i As Integer
    Dim escogio As Integer
    i = ctlGrillaArray1.Bookmark
    With objVenta
    i = 0
    'TotalEscoge
    While i < xTmp.Count(1)
        escogio = escogio + Val(xTmp(i, 2))
        i = i + 1
    Wend
    If escogio <> TotalEscoge Then
        MsgBox "tiene que seleccionar " & TotalEscoge & " y selecciono " & escogio, vbCritical, App.ProductName
        Exit Sub
    End If
    i = 0
    While i < xTmp.Count(1)
    If xTmp(i, 2) > 0 Then
    .AgregaProducto xTmp(i, 0), _
    xTmp(i, 1), _
    Val(xTmp(i, 2)), _
    xTmp(i, 3), _
    xTmp(i, 4), _
    xTmp(i, 5), _
    xTmp(i, 6), _
    xTmp(i, 7), _
    xTmp(i, 8), _
    xTmp(i, 9), _
    xTmp(i, 10), _
    xTmp(i, 11), _
    xTmp(i, 12), _
    xTmp(i, 13), _
    xTmp(i, 14), _
    xTmp(i, 15), _
    xTmp(i, 16), _
    xTmp(i, 17), _
    xTmp(i, 18), _
    xTmp(i, 19), _
    xTmp(i, 20), _
    xTmp(i, 21), _
    xTmp(i, 22)
    
    .AGREGAREGALOBK xTmp(i, 0), _
    xTmp(i, 1), _
    Val(xTmp(i, 2)), _
    xTmp(i, 3), _
    xTmp(i, 4), _
    xTmp(i, 5), _
    xTmp(i, 6), _
    xTmp(i, 7), _
    xTmp(i, 8), _
    xTmp(i, 9), _
    xTmp(i, 10), _
    xTmp(i, 11), _
    xTmp(i, 12), _
    xTmp(i, 13), _
    xTmp(i, 14), _
    xTmp(i, 15), _
    xTmp(i, 16), _
    xTmp(i, 17), _
    xTmp(i, 18), _
    xTmp(i, 19), _
    xTmp(i, 20), _
    xTmp(i, 21), _
    xTmp(i, 22)
    Else
    .AGREGAREGALOBK xTmp(i, 0), _
    xTmp(i, 1), _
    0, _
    xTmp(i, 3), _
    xTmp(i, 4), _
    xTmp(i, 5), _
    xTmp(i, 6), _
    xTmp(i, 7), _
    xTmp(i, 8), _
    xTmp(i, 9), _
    xTmp(i, 10), _
    xTmp(i, 11), _
    xTmp(i, 12), _
    xTmp(i, 13), _
    xTmp(i, 14), _
    xTmp(i, 15), _
    xTmp(i, 16), _
    xTmp(i, 17), _
    xTmp(i, 18), _
    xTmp(i, 19), _
    xTmp(i, 20), _
    xTmp(i, 21), _
    xTmp(i, 22)
    End If
    i = i + 1
Wend
    End With

    Unload Me
 End If
End Sub

Private Sub Form_Load()
SetGrd


End Sub

Function Carga(strPromocionActual As String) As String

xTmp.ReDim 0, -1, 0, 29
Dim t As Integer
Dim CantSeleccionada As Integer
While t < objVenta.xProductoRegalo.Count(1)
Dim p As Integer
    If objVenta.xProductoRegalo(t, 18) = strPromocionActual Then
        xTmp.AppendRows
    xTmp(p, 0) = objVenta.xProductoRegalo(t, 0)
    xTmp(p, 1) = objVenta.xProductoRegalo(t, 1)
        Dim KO As Integer
        xTmp(p, 2) = 0
    While KO < xProductoRegaloBK.Count(1)
    
    If objVenta.xProductoRegalo(KO, 0) = xProductoRegaloBK(KO, 0) And objVenta.xProductoRegalo(KO, 6) = xProductoRegaloBK(KO, 6) Then
        xTmp(p, 2) = Val(xProductoRegaloBK(KO, 2))
        Else
        xTmp(p, 2) = Val(objVenta.xProductoRegalo(t, 2))
        End If
    KO = KO + 1
    Wend

    xTmp(p, 3) = objVenta.xProductoRegalo(t, 3)
    xTmp(p, 4) = objVenta.xProductoRegalo(t, 4)
    xTmp(p, 5) = objVenta.xProductoRegalo(t, 5)
    xTmp(p, 6) = objVenta.xProductoRegalo(t, 6)
    xTmp(p, 7) = objVenta.xProductoRegalo(t, 7)
    xTmp(p, 8) = objVenta.xProductoRegalo(t, 8)
    xTmp(p, 9) = objVenta.xProductoRegalo(t, 9)
    xTmp(p, 10) = objVenta.xProductoRegalo(t, 10)
    xTmp(p, 11) = objVenta.xProductoRegalo(t, 11)
    xTmp(p, 12) = objVenta.xProductoRegalo(t, 12)
    xTmp(p, 13) = objVenta.xProductoRegalo(t, 13)
    xTmp(p, 14) = objVenta.xProductoRegalo(t, 14)
    xTmp(p, 15) = objVenta.xProductoRegalo(t, 15)
    xTmp(p, 16) = objVenta.xProductoRegalo(t, 16)
    xTmp(p, 17) = objVenta.xProductoRegalo(t, 17)
    xTmp(p, 18) = objVenta.xProductoRegalo(t, 18)
    xTmp(p, 19) = objVenta.xProductoRegalo(t, 19)
    xTmp(p, 20) = objVenta.xProductoRegalo(t, 20)
    xTmp(p, 21) = objVenta.xProductoRegalo(t, 21)
    xTmp(p, 22) = objVenta.xProductoRegalo(t, 22)
    xTmp(p, 23) = objVenta.xProductoRegalo(t, 23)
    xTmp(p, 24) = objVenta.xProductoRegalo(t, 24)
    xTmp(p, 25) = objVenta.xProductoRegalo(t, 25)
    
    CantSeleccionada = CantSeleccionada + Val(objVenta.xProductoRegalo(t, 23))
    p = p + 1
    End If
    t = t + 1
Wend

ctlGrillaArray1.Array1 = xTmp
ctlGrillaArray1.Rebind
TotalEscoge = objVenta.ObtieneCuentaMaxima(strPromocionActual)  'objVenta.CUANTAMAXIMA 'ECASTILLO 04.07.2020
'Debug.Print objVenta.CUANTAMAXIMA

lblMensjae.Caption = "Tiene que escoger la cantidad de " & TotalEscoge & " productos"
Carga = objVenta.CUANTOTENGOREGALO(strPromocionActual) - xTmp(p - 1, 24)
End Function


Private Sub SetGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("#", "Descripción", "Cantidad", "Precio", "Precio", "Medida", "Base", "Cant.", "% Margen", "Precio", "Sub Total", "Cod Unico", "Cod Und Med")
    arrAncho = Array(0, 5000, 700, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0)
    arrAlineacion = Array(dbgGeneral, dbgGeneral, dbgLeft, dbgGeneral, dbgLeft, dbgLeft, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgGeneral, dbgGeneral)
    ctlGrillaArray1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    ctlGrillaArray1.Columns(4).AllowFocus = True
    
    
    ctlGrillaArray1.Columns(1).Locked = True
    ctlGrillaArray1.Columns(1).AllowFocus = False
    ctlGrillaArray1.Columns(2).Locked = False
    ctlGrillaArray1.AllowUpdate = True
    
    ctlGrillaArray1.Columns(4).Locked = True
    ctlGrillaArray1.Columns(4).AllowFocus = False
    ctlGrillaArray1.EditActive = True
    ctlGrillaArray1.Columns(2).EditBackColor = RGB(255, 255, 204)
    
    ctlGrillaArray1.Columns(0).Visible = False
    ctlGrillaArray1.Columns(3).Visible = False
    ctlGrillaArray1.Columns(5).Visible = False
    ctlGrillaArray1.Columns(6).Visible = False
    ctlGrillaArray1.Columns(7).Visible = False
    ctlGrillaArray1.Columns(8).Visible = False
    ctlGrillaArray1.Columns(9).Visible = False
    ctlGrillaArray1.Columns(10).Visible = False
    ctlGrillaArray1.Columns(11).Visible = False
    ctlGrillaArray1.Columns(12).Visible = False
    
    

End Sub

