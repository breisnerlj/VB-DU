VERSION 5.00
Begin VB.Form frm_VTA_Tipo_NotaCredito 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1875
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraDev 
      Caption         =   "Devolución"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3015
      Begin vbp_Ventas.ctlTextBox TxtCtdUnd 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
      Begin vbp_Ventas.ctlTextBox TxtCtdFra 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label1 
         Caption         =   "Fracción"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame FraDscto 
      Caption         =   "Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton opt 
         Caption         =   "Monto"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opt 
         Caption         =   "Porcentaje"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin vbp_Ventas.ctlTextBox TxtDato 
         Height          =   330
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label LblSoles 
         Caption         =   "S/."
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
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblPct 
         Caption         =   "%"
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
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFBFA&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Importe Dscto es aplicado al subtotal    de venta del item"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESC =>"
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
      Index           =   2
      Left            =   788
      TabIndex        =   11
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   1508
      TabIndex        =   10
      Top             =   1080
      Width           =   765
   End
End
Attribute VB_Name = "frm_VTA_Tipo_NotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnConcepto As Boolean
Dim objProducto As New clsProducto
Dim intCant As Integer
Dim dblPrecio As Double

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo handle
  If blnConcepto = False Then 'Devolucion'
        
        If (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3) > gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4)) Then
            TxtCtdUnd.Text = "0"
            TxtCtdFra.Text = "0"
        ElseIf (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4) > gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3)) Then
            TxtCtdUnd.Text = "0"
            TxtCtdFra.Text = "0"
        ElseIf (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4) > 0 And gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3)) > 0 Then
            TxtCtdUnd.Text = "0"
            TxtCtdFra.Text = "0"
        End If
  Else 'Porcentaje'
        TxtDato.Text = IIf(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = "", "0", gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7))
  End If
  Dim strIdFrac  As String
  strIdFrac = ""
   strIdFrac = objProducto.ListaDevFracciona(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1), objUsuario.CodigoLocal, "")
   If strIdFrac = "1" Then
    TxtCtdUnd.Enabled = True
    TxtCtdFra.Enabled = True
   Else
    TxtCtdUnd.Enabled = True
    TxtCtdFra.Enabled = False
   End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
  
End Sub
    
Private Sub opt_Click(Index As Integer)
On Error GoTo handle
    ' blnConcepto = true "Porcentaje"
    ' blnConcepto = false "Devolucion"
    If opt(0).Value = True Then
           'Porcentaje
           blnConcepto = True
           SendKeys "{TAB}"
       Else
           'Devolucion
           blnConcepto = False
           SendKeys "{TAB}"
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub TxtCtdUnd_KeyPress(KeyAscii As Integer)
    '----------------------'
    '*** POR DEVOLUCION ***'
    '----------------------'
    TxtCtdUnd.Tipo = Entero
    objVenta.FlgNcrDev = "1"
    If KeyAscii = 13 Then
        If Val(TxtCtdUnd.Text) > 0 And Val(TxtCtdFra.Text) = 0 Then
        'Unidades'
            If Val(TxtCtdUnd.Text) > Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3)) Then MsgBox "Ctd Devuelta no puede ser mayor a ctd doc", vbCritical, "Revisar": Exit Sub
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = Val(TxtCtdUnd.Text)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), 2)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "0"     'Flg Fraccion
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
            
        Else
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = Val(TxtCtdUnd.Text)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), 2)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "0"     'Flg Fraccion
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
            
        End If
        If TxtCtdFra.Enabled = False Then TxtCtdFra_KeyPress 13
    End If
End Sub

Private Sub TxtCtdFra_KeyPress(KeyAscii As Integer)
    '----------------------'
    '*** POR DEVOLUCION ***'
    '----------------------'
    
    Dim Indicador As String
    Dim PctComi As Double
    On Error GoTo handle
    TxtCtdFra.Tipo = Entero
    objVenta.FlgNcrDev = "1"
    If KeyAscii = 13 Then
        If Val(TxtCtdFra.Text) > 0 And Val(TxtCtdUnd.Text) = 0 Then
        'Fracciones'
            If Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4)) >= Val(TxtCtdFra.Text) Then 'MsgBox "Ctd Devuelta no puede ser mayor a ctd doc", vbCritical, "Revisar": Exit Sub
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = Val(TxtCtdFra.Text)
                'gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(((Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) / Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9)))) * ((Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3)) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9))) + Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7))), 2)
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(((Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) / Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9)))) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), 2)
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "1" 'Flg Fraccion
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
              Else
                Dim dblConvertFra As Double
                dblConvertFra = Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3)) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9))
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = Val(TxtCtdFra.Text)
                If gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) > dblConvertFra Then MsgBox "Ctd Devuelta no puede ser mayor a ctd doc", vbCritical, "Revisar": Exit Sub
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(Val((gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) / Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9))) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), 2)
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "1" 'Flg Fraccion
                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
            End If
            
        ElseIf Val(TxtCtdUnd.Text) > 0 And Val(TxtCtdFra.Text) > 0 Then
        'Unidades y Fracciones'
            If (Val(TxtCtdUnd.Text) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9)) + Val(TxtCtdFra.Text)) > (Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3)) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9)) + Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4))) Then MsgBox "Ctd Devuelta no puede ser mayor a ctd doc", vbCritical, "Revisar": Exit Sub
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = Val(TxtCtdUnd.Text) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9)) + Val(TxtCtdFra.Text)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(Val((gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) / Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9))) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), 2)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "1"  'Flg Fraccion
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
            
        ElseIf Val(TxtCtdUnd.Text) > 0 And Val(TxtCtdFra.Text) = 0 Then
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = Val(TxtCtdUnd.Text)
            'gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(Val((gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) / Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 9))) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), 2)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 5)) * Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), 2)
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "0"  'Flg Fraccion
            gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
        
        End If
        
        Indicador = objProducto.CodIndicadorReceta(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1))
        PctComi = objProducto.pctComision(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1), objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
        
        
        objVenta.AgregaProducto gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1), _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 2), _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7), _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10), _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8), _
                                Venta_Regular, _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 12), _
                                , , , , , Indicador, PctComi
        ''arturo escate 12/11/2009 es para q limpie el producto antes de agregarlo
        If Val(TxtCtdFra.Text) + Val(TxtCtdUnd.Text) = 0 Then objVenta.EliminaProductoNew gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1), Venta_Regular
        CalFooter
        
        frm_VTA_NotaCredito.grdNC.Rebind
        Unload Me
        End If
        Exit Sub
handle:
        MsgBox Err.Description, vbCritical, App.ProductName
End Sub
    
Private Sub txtDato_KeyPress(KeyAscii As Integer)
    '----------------------'
    '*** POR DESCUENTO ***'
    '----------------------'
    Dim Indicador As String
    Dim PctComi  As Double
    On Error GoTo handle
    TxtDato.Tipo = Real
    If KeyAscii = 13 Then
            If blnConcepto = True Then
              'Porcentaje'
              objVenta.FlgNcrDev = "0"   'Permite saber que la NC es por descuento'
              gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = ""
              If Val(TxtDato.Text) = 0 Then
                    gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = "0"
                Else
                    gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = Round(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 6), 2) ' * ((TxtDato.Text) / 100), 2)
              End If
              If Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8)) > Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 6)) Then MsgBox "Descuento mayor al IMPORTE del Doc", vbCritical, Caption: Exit Sub
         Else
              'Devolucion'
              objVenta.FlgNcrDev = "1"   'Permite saber que la NC es por devolución'
              gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = ""
              gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8) = TxtDato.Text
              If Val(TxtDato.Text) > Val(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 6)) Then MsgBox "Monto de Devolución mayor al IMPORTE del Doc", vbCritical, Caption: Exit Sub
        End If
            
        If (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4) > 0) And (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3) <= 0) Then  'Fracciones'
           gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "1"
           gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "1"
        'Unidades'
        ElseIf (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3) > 0) And (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4) <= 0) Then
           gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "0"
           gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
        'Undidades y Fracciones'
        ElseIf (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 3) > 0) And (gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 4) > 0) Then
           gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10) = "1"
           gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 11) = "0"
        End If
        
        Indicador = objProducto.CodIndicadorReceta(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1))
        PctComi = objProducto.pctComision(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1), objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
            objVenta.AgregaProducto gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1), _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 2), _
                                IIf(gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7) = "", "0", gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 7)), _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 10), _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 8), _
                                Venta_Regular, _
                                gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 12), _
                                , , , , , Indicador, PctComi
        ''arturo escate 12/11/2009 es para q limpie el producto antes de agregarlo
        If Not blnConcepto = True Then
            If Val(TxtCtdFra.Text) + Val(TxtCtdUnd.Text) = 0 Then objVenta.EliminaProductoNew gxdbNC(frm_VTA_NotaCredito.grdNC.Bookmark, 1), Venta_Regular
        End If
        CalFooter
        
        frm_VTA_NotaCredito.grdNC.Rebind
        Unload Me
    
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub CalFooter()
    Dim k%
    intCant = 0: dblPrecio = 0
    For k = 0 To gxdbNC.UpperBound(1)
        If gxdbNC(k, 7) > 0 Or gxdbNC(k, 7) <> "" Then
            intCant = intCant + Val(gxdbNC(k, 7))
        End If
        If gxdbNC(k, 8) > 0 Or gxdbNC(k, 8) <> "" Then
            dblPrecio = dblPrecio + Val(gxdbNC(k, 8))
        End If
    Next k
    frm_VTA_NotaCredito.grdNC.Columns(7).FooterText = intCant
    frm_VTA_NotaCredito.grdNC.Columns(8).FooterText = Format(dblPrecio, "#,###,##0.00")
End Sub
    
