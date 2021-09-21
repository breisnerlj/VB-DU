VERSION 5.00
Begin VB.Form frm_ADM_TranspGuia 
   Caption         =   "Guías Pendientes"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "[F5] Todos"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "[F11] Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "[Esc] Cerrar"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "[F3] Ver Detalle"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtidentrega 
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin vbp_Ventas.ctlTextBox txtBuscar 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
      Begin VB.Label lblGuiasSelec 
         Alignment       =   1  'Right Justify
         Caption         =   "0 Seleccionada(s)"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin vbp_Ventas.ctlGrillaArray ctlgrdguias 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6376
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_TranspGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objEntrega As New clsEntrega
Dim xDetalle As New XArrayDB
Dim strIdEntrega As String
Public flgEsAdmEntrega As Integer

Public Sub carga(ByVal idEntrega As String, Optional esAdmEntrega As Integer = 1)
    strIdEntrega = idEntrega
    flgEsAdmEntrega = esAdmEntrega
    Me.txtidentrega = strIdEntrega
    xDetalle.ReDim 0, -1, 0, 7
    cargaDetalle "", ""
    Me.Caption = "Guias Pendientes de Recepcionar"
    Me.Show vbModal
End Sub

Sub cargaDetalle(idEntrega As String, numGuia As String)
    Dim i As Integer
    Dim rs As oraDynaset
    Set rs = objEntrega.ListaPendiente(objUsuario.CodigoLocal, idEntrega, numGuia)
    i = 0
    While Not rs.EOF
    xDetalle.AppendRows
        xDetalle(i, 0) = rs("FLG_SELECCIONADO").Value * (-1)
        xDetalle(i, 1) = rs("NUM_GUIA").Value
        xDetalle(i, 2) = "" & rs("NUM_ENTREGA").Value
        xDetalle(i, 3) = rs("FCH_RECEPCION").Value
        xDetalle(i, 4) = rs("FCH_EMISION").Value
        xDetalle(i, 5) = rs("INDICADOR").Value
        xDetalle(i, 6) = rs("NUM_FACTURA_SAP").Value
        i = i + 1
        rs.MoveNext
    Wend
    SeteaGrilla
    Me.ctlgrdguias.Array1 = xDetalle
End Sub

Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant

    arrCampos = Array("FLG_SELECCIONADO", "NUM_GUIA", "NUM_ENTREGA", "FCH_RECEPCION", "FCH_EMISION", "INDICADOR", "NUM_FACTURA_SAP")
    arrCaption = Array("X", "Nº Guía", "Nº Entrega", "Fec. Recepcion", "Fec. Emision", "F", "Nº Factura")
    arrAncho = Array(800, 1200, 1300, 2000, 2000, 500, 1200)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgLeft)
    arrFoco = Array(True, False, False, False, False, False, False)
    Me.ctlgrdguias.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    Me.ctlgrdguias.AllowUpdate = True
    
    Me.ctlgrdguias.Columns(0).ValueItems.Presentation = dbgCheckBox
    Me.ctlgrdguias.Columns(4).Merge = False
    Me.ctlgrdguias.Columns(1).Merge = False
    Me.ctlgrdguias.Columns(2).Merge = False
    Me.ctlgrdguias.Columns(3).Merge = False
    Me.ctlgrdguias.Columns(5).Merge = False
    Me.ctlgrdguias.Columns(6).Merge = False
    
    ctlgrdguias.EditorStyle.BackColor = vbWhite
    ctlgrdguias.EditorStyle.ForeColor = RGB(180, 0, 180)
    ctlgrdguias.EditorStyle.Font.Bold = True
    
    ctlgrdguias.Columns(1).BackColor = vbInfoBackground
    ctlgrdguias.Columns(2).BackColor = vbInfoBackground
    ctlgrdguias.Columns(3).Visible = False
    ctlgrdguias.Columns(5).Visible = False
End Sub

Private Sub cmdAceptar_Click()
Dim strDatos As String

    If BuscarDatosSeleccionados = False Then
        MsgBox "No se ha seleccionado Guía.", vbCritical, "Error"
        Exit Sub
    End If
    Dim msbo As Variant
    msbo = MsgBox("¿Seguro que desea asociar las guías seleccionadas?", vbYesNo + vbInformation, App.ProductName)
    If msbo = vbYes Then
        strDatos = Graba
        If strDatos <> "" Then
            GoTo salir
        End If
        If flgEsAdmEntrega = 1 Then
        frm_ADM_Entrega.Consulta
        frm_ADM_Entrega.grdRecepcion.DataSource.FindFirst "ID_ENTREGA='" & Trim(Me.txtidentrega.Text) & "'"
        Else
        frm_ADM_Entrega.Consulta
        frm_ADM_Entrega.grdRecepcion.DataSource.FindFirst "ID_ENTREGA='" & Trim(Me.txtidentrega.Text) & "'"
        frm_ADM_GuiaAsociada.strIdEntrega = strIdEntrega
        frm_ADM_GuiaAsociada.cargaCabGuias
        End If
        Unload Me
    End If
salir:


End Sub

Private Function BuscarDatosSeleccionados()
Dim j As Integer
Me.ctlgrdguias.Update
Me.ctlgrdguias.MoveNext
Me.ctlgrdguias.MovePrevious
    BuscarDatosSeleccionados = False
    For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
        If xDetalle(j, 0) = -1 Then
           j = xDetalle.UpperBound(1) + 1
           BuscarDatosSeleccionados = True
        End If
    Next
End Function

Function Graba() As String

'On Error GoTo CtrlErr
Dim i As Integer
Dim Entrega As String
Dim arrGuias As String
Dim strDatos As String

arrGuias = ""
While i < xDetalle.Count(1)
    If Val(xDetalle(i, 0)) <> 0 Then
        arrGuias = arrGuias & xDetalle(i, 1) & "|"
    End If
    i = i + 1
Wend
strDatos = objEntrega.GrabaGuias(Trim(Me.txtidentrega.Text), arrGuias)

If strDatos <> "" Then
    MsgBox strDatos
    Graba = strDatos
End If


'CtrlErr:
    'Err.Raise Err.Description, "clsLocal.GrabaGuias", Err.Description
End Function



Private Sub cmdCancelar_Click()
     If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
       Unload Me
    End If
End Sub

Private Sub cmdDetalle_Click()
    VerDetGuia
End Sub


Sub VerDetGuia()
    Dim mensaje As String
    If Me.ctlgrdguias.ApproxCount <= 0 Then
        Exit Sub
    End If
    mensaje = Me.ctlgrdguias.Columns(1).Value
    frm_ADM_DetGuias.numGuia = mensaje
    frm_ADM_DetGuias.Show vbModal
End Sub

Private Sub Command1_Click()
    Dim j As Integer
    Me.ctlgrdguias.Update
    If Me.ctlgrdguias.ApproxCount <= 0 Then
        MsgBox "No exiten guias que seleccionar.", vbCritical, "Error"
        Exit Sub
    End If
    If xDetalle(xDetalle.LowerBound(1), 0) = 0 Then
        For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
            xDetalle(j, 0) = -1
        Next
        Me.lblGuiasSelec.Caption = CStr(xDetalle.UpperBound(1)) + " Seleccionado(s)."
        Me.Command1.Caption = "[F5] Ninguno"
    Else
        For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
            xDetalle(j, 0) = 0
        Next
        Me.lblGuiasSelec.Caption = "0 Seleccionado(s)."
        Me.Command1.Caption = "[F5] Todos"
    End If
    Me.ctlgrdguias.Rebind
End Sub

Private Sub ctlgrdguias_AfterColUpdate(ByVal ColIndex As Integer)
    CalcularSeleccionados
'    If VerificarFechas Then
'        CalcularSeleccionados
'    Else
'        If xDetalle(ctlgrdguias.Bookmark, 0) = -1 Then
'           xDetalle(ctlgrdguias.Bookmark, 0) = -1
'        Else
'           xDetalle(ctlgrdguias.Bookmark, 0) = 0
'        End If
'        Me.ctlgrdguias.Rebind
'    End If
End Sub

Private Sub CalcularSeleccionados()
    Me.ctlgrdguias.Update
    Dim j As Integer
    Dim acum As Integer
    acum = 0
    For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
            If xDetalle(j, 0) = -1 Then
               acum = acum + 1
            End If
    Next
    Me.lblGuiasSelec.Caption = CStr(acum) + " Seleccionado(s)."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then
    cmdAceptar_Click
End If
If KeyCode = vbKeyF3 Then
    VerDetGuia
End If
If KeyCode = vbKeyF5 Then
    Me.ctlgrdguias.Update
    Dim j As Integer
    If xDetalle(xDetalle.LowerBound(1), 0) = 0 Then
        For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
            xDetalle(j, 0) = -1
        Next
        Me.Command1.Caption = "[F5] Ninguno"
        Me.lblGuiasSelec.Caption = CStr(xDetalle.UpperBound(1)) + " Seleccionado(s)."
    Else
        For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
            xDetalle(j, 0) = 0
        Next
        Me.lblGuiasSelec.Caption = "0 Seleccionado(s)."
        Me.Command1.Caption = "[F5] Todos"
    End If
    Me.ctlgrdguias.Rebind
End If
If KeyCode = vbKeyF9 Then
    Me.txtBuscar.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    Dim INDEX As Long
    Dim colbus As Integer
    
    If KeyAscii = 13 Then
        If xDetalle.Count(1) > 0 Then
            colbus = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "COLBUSREC")
                INDEX = xDetalle.Find(0, colbus, Me.txtBuscar.Text)
                If INDEX <> -1 Then
                    xDetalle(INDEX, 0) = -1
                    ctlgrdguias.Rebind
                    Me.ctlgrdguias.Bookmark = INDEX
                    CalcularSeleccionados
                    SendKeys "{TAB}"
                    SendKeys "{TAB}"
                    SendKeys "{TAB}"
                    SendKeys "{TAB}"
                    SendKeys "{TAB}"
                End If
        End If
    End If
End Sub

Private Sub ctlgrdguias_KeyPress(KeyAscii As Integer)
       If KeyAscii = 13 Then
          'If VerificarFechas Then
          If xDetalle.Count(1) > 0 Then
             If xDetalle(ctlgrdguias.Bookmark, 0) = -1 Then
                xDetalle(ctlgrdguias.Bookmark, 0) = 0
             Else
                xDetalle(ctlgrdguias.Bookmark, 0) = -1
             End If
             Me.ctlgrdguias.Rebind
             CalcularSeleccionados
          'Else
              'MsgBox "No se Puede Mezclar Guias con Diferente Corte", vbCritical, "Error"
          'End If
          End If
       End If
End Sub

Private Function VerificarFechas() As Boolean
Dim j As Integer
VerificarFechas = True
For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
     If j <> ctlgrdguias.Bookmark Then
        If xDetalle(j, 0) = -1 Then
            If xDetalle(ctlgrdguias.Bookmark, 5) <> xDetalle(j, 5) Then
               VerificarFechas = False
               GoTo Termina
            End If
        End If
     End If
 Next
Termina:
End Function
