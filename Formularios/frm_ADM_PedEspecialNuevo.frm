VERSION 5.00
Begin VB.Form frm_ADM_PedEspecialNuevo 
   Caption         =   "Nuevo Pedido Especial"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFiltro 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10815
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "&Añadir"
         Height          =   615
         Left            =   9480
         Picture         =   "frm_ADM_PedEspecialNuevo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Añade Productos a Detalle"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Li&mpiar"
         Height          =   615
         Left            =   9480
         Picture         =   "frm_ADM_PedEspecialNuevo.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Elimina los No Solicitados"
         Top             =   960
         Width           =   1095
      End
      Begin vbp_Ventas.ctlDataCombo cboLaboratorio 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboLinea 
         Height          =   315
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox TxtProducto 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   661
         Tipo            =   8
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
         Caption         =   "&Laboratorio"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Lí&nea"
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "&Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ACT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7800
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCod_Producto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "77700"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BROCODUAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1080
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   6615
      End
   End
   Begin vbp_Ventas.ctlGrillaArray grdDetalle 
      Height          =   4335
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7646
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1058
      ModoBotones     =   10
      EnabledEfecto   =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_PedEspecialNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public v_accion As Integer
Public v_numPedido As String
Private objSPVM As New clsPedEspecial

Public Property Get solicitud() As clsPedEspecial
    Set solicitud = objSPVM
End Property

Public Property Set solicitud(ByVal vstrNewValue As clsPedEspecial)
    Set objSPVM = vstrNewValue
End Property

Private Sub cboLaboratorio_Change()
    On Error GoTo CtrlErr
    
    CargaLinea cboLaboratorio.BoundText
    LimpiaProducto
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdAñadir_Click()
    Dim odynTemp As oraDynaset
    
    On Error GoTo CtrlErr
    'Dim Valor As String
    'Valor = "00"
    
'''    If Me.grdDetalle.ApproxCount > 0 Then
'''        Valor = Me.grdDetalle.Columns(9).Value
'''    End If
    
'''    If validaIngresado(Valor) Then
        If lblCod_Producto.Caption = "" And cboLaboratorio.BoundText = "*" Then
            MsgBox "Seleccione criterio de búsqueda, ya sea línea o Producto", vbExclamation, "Aviso"
            Exit Sub
        End If
        
        If cboLaboratorio.BoundText <> "*" And cboLinea.BoundText = "*" Then
            MsgBox "Seleccione línea", vbExclamation, "Aviso"
            cboLinea.SetFocus
            Exit Sub
        End If
        
        Set odynTemp = solicitud.Filtro(solicitud.CodLocal, _
                                        lblCod_Producto.Caption, _
                                        cboLaboratorio.BoundText, _
                                        cboLinea.BoundText)
        If odynTemp.RecordCount = 0 Then
            MsgBox "No se ubicaron productos activos con criterio de búsqueda", vbExclamation, "Aviso"
            
            If lblCod_Producto.Caption <> "" Then
                TxtProducto.SetFocus
            Else
                cboLinea.SetFocus
            End If
            
        Else
            AdicionaDetalle odynTemp
        End If
'''    End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdLimpiar_Click()
    On Error GoTo CtrlErr
'''    Dim Valor As String
'''    Valor = "00"
'''    If Me.grdDetalle.ApproxCount > 0 Then
'''        Valor = Me.grdDetalle.Columns(9).Value
'''    End If
'''    If validaIngresado(Valor) Then
    
        If grdDetalle.ApproxCount < 1 Then Exit Sub
        
        If MsgBox("¿Seguro(a) de Eliminar Items que no tienen cantidades de solicitud?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
        
        grdDetalle.MoveFirst

        While Not grdDetalle.EOF
            Screen.MousePointer = vbHourglass
            If "" & grdDetalle.Columns(9).Value = "" Then
                grdDetalle.Delete
            Else
                grdDetalle.MoveNext
            End If
            Screen.MousePointer = vbDefault
        Wend
            
        Me.grdDetalle.MoveFirst
'''    End If
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    On Error GoTo CtrlErr
    Select Case boton
        Case Nuevo
        Case Modificar
        Case Buscar
        Case tb_Actualizar
        Case Imprimir
        Case tb_Excel
        Case tb_email
        Case Grabar
            If v_accion = 1 Then
            Graba
            ElseIf v_accion = 2 Then
            Actualiza
            End If
        Case Cancelar
            Cancela
        Case Eliminar
        Case salir

    End Select
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Cancela
    End If
End Sub

Public Sub Form_Load()
    On Error GoTo CtrlErr
    SeteaGrilla
    CargaLaboratorio
    LimpiaProducto
    'CargaMotivos
    
    Dim objEstVta As New clsEstadistica
    
    objEstVta.Limpia
    
    Set objEstVta = Nothing
    
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdDetalle_AfterColUpdate(ByVal ColIndex As Integer)
    grdDetalle.MoveNext
    grdDetalle.MovePrevious
'''    validaGrilla
End Sub

Private Sub grdDetalle_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Select Case ColIndex
        Case 9
            Cancel = False
        Case Else
            Cancel = True
    End Select
End Sub

'''Function validaIngresado(Valor As String) As Boolean
'''Dim Cancel As Boolean
'''If Valor <> "00" Then
'''    Cancel = True
'''    If (Not IsNumeric(Valor) And Valor <> "") Or (Len(Valor) > 4) Then
'''       MsgBox "El valor ingresado no es válido", vbExclamation, "Error"
'''       Me.grdDetalle.Columns(9).Value = ""
'''       Cancel = False
'''    ElseIf Trim(Valor) <> "" Then
'''        Cancel = True
'''    End If
'''Else
'''    Cancel = True
'''End If
'''validaIngresado = Cancel
'''End Function

Private Sub grdDetalle_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'''    Dim objCol As TrueDBGrid70.Column
'''
'''    Set objCol = Me.grdDetalle.Columns(ColIndex)
'''    Select Case ColIndex
'''        Case 9
'''            If (Not IsNumeric(objCol.Text) And objCol.Text <> "") Or (Len(objCol.Text) > 4) Then
'''                MsgBox "El valor ingresado no es válido", vbExclamation, "Error"
'''                Cancel = True
'''            ElseIf Trim(objCol.Text) <> "" Then
'''                objCol.Text = Trim(objCol.Text)
'''            End If
'''    End Select
    Select Case ColIndex
        Case 9
            If Not IsNumeric(Trim(grdDetalle.Columns(ColIndex).Value)) And _
                    Trim(grdDetalle.Columns(ColIndex).Value) <> "" Then
                MsgBox "El valor no es valido", vbCritical, "Error"
                Cancel = True
                Exit Sub
            End If
    End Select
End Sub

Private Sub grdDetalle_ButtonClick(ByVal ColIndex As Integer)
    
    On Error GoTo CtrlErr
    
    Select Case ColIndex
        Case 7 'ventas
            
        Dim frm As New frm_ADM_Graf_Ventas
        Dim objDatos As New clsEstadistica
        Dim odynDatos As oraDynaset
        Dim varDaTos(0 To 1, 0 To 9) As Variant
    
            
        Set odynDatos = objDatos.Lista(solicitud.CodLocal, grdDetalle.Columns(0).Value)
        
        odynDatos.MoveFirst
        
        While Not odynDatos.EOF
                    
            varDaTos(1, odynDatos("ORDEN").Value) = odynDatos("CTD_VENTA").Value
            varDaTos(0, odynDatos("ORDEN").Value) = odynDatos("PERIODO").Value
            odynDatos.MoveNext
            
        Wend
        
        varDaTos(1, 0) = "Cantidad de Venta"
        frm.Titulo = grdDetalle.Columns(0).Value & ": " & grdDetalle.Columns(1).Value
        frm.Datos = varDaTos
        frm.Mostrar
            
        Set frm = Nothing
            
            
    End Select
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdDetalle_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    On Error GoTo CtrlErr
        
    Select Case Col
        Case 1
            If grdDetalle.Columns("sel").CellValue(Bookmark) = "1" Then
                CellStyle.ForeColor = vbRed
                CellStyle.Font.Bold = True
            End If
            
        Case 9 'Campos a ingresar
            Select Case Condition
                Case CellStyleConstants.dbgMarqueeRow
                    CellStyle.Font.Bold = True
                Case CellStyleConstants.dbgMarqueeRow + CellStyleConstants.dbgCurrentCell
                    CellStyle.BackColor = vbCyan  'RGB(60, 255, 255)  'vbInfoBackground
                    CellStyle.Font.Bold = True
            End Select
        End Select
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error"

End Sub

Private Sub grdDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo CtrlErr
    
    If grdDetalle.ApproxCount < 1 Then Exit Sub
    
    If Not grdDetalle.EditActive Then
    
        If KeyCode = vbKeyDelete Then
            If MsgBox("¿Desea eliminar Item de la lista?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Item") = vbYes Then
                grdDetalle.Delete
                grdDetalle.Refresh
            End If
        End If
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdDetalle_LostFocus()
    grdDetalle.Refresh
End Sub

''''Private Sub grdDetalle_Validate(Cancel As Boolean)
''''validaGrilla
''''End Sub
''''
''''Sub validaGrilla()
''''    If Len(grdDetalle.Columns(9).Text) > 4 Then
''''        MsgBox "Ingrese Nº menor de 4 dígitos", vbCritical + vbInformation, "Error"
''''        grdDetalle.Columns(9).Text = ""
''''
''''    End If
''''
''''    Dim Valor As String
''''    Dim i As Integer
''''    Valor = grdDetalle.Columns(9).Text
''''    If Len(Valor) > 0 Then
''''        For i = 1 To Len(Valor)
''''            If Mid(Valor, i, 1) <> "0" And Mid(Valor, i, 1) <> "1" _
''''            And Mid(Valor, i, 1) <> "2" And Mid(Valor, i, 1) <> "3" _
''''            And Mid(Valor, i, 1) <> "4" And Mid(Valor, i, 1) <> "5" _
''''            And Mid(Valor, i, 1) <> "6" And Mid(Valor, i, 1) <> "7" _
''''            And Mid(Valor, i, 1) <> "8" And Mid(Valor, i, 1) <> "9" Then
''''            MsgBox "Ingrese Nº correcto", vbCritical + vbInformation, "Error"
''''            grdDetalle.Columns(9).Text = ""
''''
''''            Exit For
''''            End If
''''        Next
''''    End If
''''End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Handle
    
    If KeyAscii = vbKeyReturn Then
        
        If Len(Trim(TxtProducto.Text)) < 3 Then
            MsgBox "use por lo menos 3 caracteres", vbExclamation, "Aviso"
            Exit Sub
        End If
        
        Dim frm As New frm_ADM_ProductoDatos
        frm.Dato = Trim(TxtProducto.Text)
        frm.Show vbModal
        
        If frm.Salida(1) <> "" Then
            cboLaboratorio.BoundText = "*"
        End If
                
        lblCod_Producto.Caption = frm.Salida(1)
        lblProducto.Caption = frm.Salida(2)
        lblEstado.Caption = frm.Salida(3)
                
        Set frm = Nothing
        
        cmdAñadir_Click
        
    End If

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub Graba()
    Dim i As Integer
'''    Dim Valor As String
'''    Valor = "00"
    
    If grdDetalle.EditActive Then
        grdDetalle.MoveNext
        grdDetalle.MovePrevious
    End If
    
'''    If Me.grdDetalle.ApproxCount > 0 Then
'''        Valor = Me.grdDetalle.Columns(9).Value
'''    End If
'''
'''    If validaIngresado(Valor) Then
        
        If grdDetalle.ApproxCount < 1 Then Exit Sub
    
        If MsgBox("¿Seguro(a) de registrar la Solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, "Grabar") = vbNo Then Exit Sub
    
    On Error GoTo Handle
        For i = solicitud.Detalle.LowerBound(1) To solicitud.Detalle.UpperBound(1)
            If "" & solicitud.Detalle(i, 9) = "" Then
                MsgBox "Se han encontrado items que no cuentan con cantidad a Solicitar" & Chr(13) & _
                       "Limpiar primero estos items con la opción ""Limpiar""", vbExclamation, "Grabar"
                Exit Sub
            End If
            If "" & solicitud.Detalle(i, 9) = "0" Then
                MsgBox "Se han encontrado items cuya cantidad a Solicitar es CERO" & Chr(13) & _
                       "Agregar una cantidad distinta de CERO", vbExclamation, "Grabar"
                Exit Sub
            End If
        Next i
        
        solicitud.CodUsuario = objUsuario.codigo
    
        solicitud.Grabar
        MsgBox "Se registró la Solicitud Número " & solicitud.Numero, vbInformation, "Grabar"
        frm_ADM_PedEspecial.Form_Load
        frm_ADM_PedEspecial.FLAG = 0
        frm_ADM_PedEspecial.spBuscar
        Unload Me
'''    End If
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Actualiza()
    Dim i As Integer
'''    Dim Valor As String
'''    Valor = "00"
'''    If Me.grdDetalle.ApproxCount > 0 Then
'''        Valor = Me.grdDetalle.Columns(9).Value
'''    End If
'''    If validaIngresado(Valor) Then
        If grdDetalle.EditActive Then
            grdDetalle.MoveNext
            grdDetalle.MovePrevious
        End If

        If grdDetalle.ApproxCount < 1 Then Exit Sub

        If MsgBox("¿Seguro(a) de actualizar la Solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, "Actualizar") = vbNo Then Exit Sub

    On Error GoTo Handle
        For i = solicitud.Detalle.LowerBound(1) To solicitud.Detalle.UpperBound(1)
            If "" & solicitud.Detalle(i, 9) = "" Then
                MsgBox "Se han encontrado items que no cuentan con cantidad a Solicitar" & Chr(13) & _
                       "Limpiar primero estos items con la opción ""Limpiar""", vbExclamation, "Grabar"
                Exit Sub
            End If
            If "" & solicitud.Detalle(i, 9) = "0" Then
                MsgBox "Se han encontrado items cuya cantidad a Solicitar es CERO" & Chr(13) & _
                       "Agregar una cantidad distinta de CERO", vbExclamation, "Grabar"
                Exit Sub
            End If
        Next i

        solicitud.CodUsuario = objUsuario.codigo
        solicitud.Numero = v_numPedido
        solicitud.Actualizar
        MsgBox "Se actualizó la Solicitud Número " & solicitud.Numero, vbInformation, "Actualizar"
        frm_ADM_PedEspecial.Form_Load
        Unload Me
'''    End If
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Cancela()
    Dim blnPosible As Boolean
    '''Dim intEncontrado As Integer
    
    blnPosible = False
    
    If grdDetalle.ApproxCount > 1 Then
        blnPosible = True
    End If
    

    If blnPosible Then
        If MsgBox("¿Seguro(a) de Salir sin grabar?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelar") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
    
End Sub

Private Sub CargaLinea(ByVal vstrCodLab As String)
    Dim objLinea As New clsLinea
    Set cboLinea.RowSource = objLinea.Lista(vstrCodLab, "1", "[SELECCIONAR]")       '''objProducto.ListaLinea(vstrCodLab)
    Set objLinea = Nothing
    
    cboLinea.ListField = "DES"
    cboLinea.BoundColumn = "COD"
    cboLinea.BoundText = "*"
    
End Sub

Private Sub CargaLaboratorio()
    Dim objLaboratorio As New clsLaboratorio
    Set cboLaboratorio.RowSource = objLaboratorio.Lista("1", "[SELECCIONAR]")
    Set objLaboratorio = Nothing
    
    cboLaboratorio.ListField = "DES"
    cboLaboratorio.BoundColumn = "COD"
    cboLaboratorio.BoundText = "*"
         
End Sub

Private Sub LimpiaProducto()
    lblCod_Producto.Caption = ""
    lblProducto.Caption = ""
    lblEstado.Caption = ""
End Sub


Public Sub AdicionaDetalle(ByRef rodynTemp As oraDynaset)
Dim intRow As Integer
Dim strMensaje As String
Dim blnAdd  As Boolean
Dim btnActualizar As Boolean
Dim inTf As Integer

    rodynTemp.MoveFirst
    strMensaje = ""
    btnActualizar = False
    intRow = 0
        
    
    While Not rodynTemp.EOF
        Screen.MousePointer = vbHourglass
        'Permite Ubicar en producto en la grilla o de lo contrario encontralo lo pinta'
        blnAdd = False
        If solicitud.Detalle.UpperBound(1) > solicitud.Detalle.LowerBound(1) - 1 Then
             inTf = solicitud.Detalle.Find(0, 0, CStr("" & rodynTemp("COD_PRODUCTO").Value), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
             If inTf = -1 Then
                   blnAdd = True
             End If
        Else
            blnAdd = True
        End If
        
                      
        If blnAdd Then
            btnActualizar = True
            solicitud.Detalle.InsertRows intRow
            solicitud.Detalle(intRow, 0) = "" & rodynTemp("COD_PRODUCTO").Value
            solicitud.Detalle(intRow, 1) = "" & rodynTemp("DES_PRODUCTO").Value
            solicitud.Detalle(intRow, 2) = "" & rodynTemp("LABORATORIO").Value
            solicitud.Detalle(intRow, 3) = "" & rodynTemp("LINEA").Value
            solicitud.Detalle(intRow, 4) = "" & rodynTemp("FLG_SELECCIONADO").Value
            solicitud.Detalle(intRow, 5) = "" '& rodynTemp("COD_EST_ABAST").Value
            solicitud.Detalle(intRow, 6) = "" '& rodynTemp("PVM_APROBADO").Value
            solicitud.Detalle(intRow, 7) = "" '& rodynTemp("VENTAS").Value
            solicitud.Detalle(intRow, 8) = "" '& rodynTemp("PVM_ACTUAL").Value
            solicitud.Detalle(intRow, 9) = "" & rodynTemp("CTD_SOLICITADA").Value
            solicitud.Detalle(intRow, 10) = "" & rodynTemp("STOCK").Value
            solicitud.Detalle(intRow, 11) = "" & rodynTemp("CATEGORIA").Value
            intRow = intRow + 1
        Else
             strMensaje = strMensaje & CStr("" & rodynTemp("COD_PRODUCTO").Value) & " - " & CStr("" & rodynTemp("DES_PRODUCTO").Value) & Chr(13)
        End If
        
        rodynTemp.MoveNext
        Screen.MousePointer = vbDefault
    Wend
        
    
    If btnActualizar = True Then
        grdDetalle.Rebind
        grdDetalle.MoveFirst
        grdDetalle.Col = 9
        'grdDetalle.SetFocus
    End If
    
    If strMensaje <> "" Then
        Me.Refresh
        MsgBox "Los siguientes productos ya se encontraban en la lista: " & Chr(13) & strMensaje, vbInformation, "Aviso"
        'grdDetalle.Bookmark = inTf
    End If
    
    
    
End Sub

Public Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
      
    '---------------------------------------------------------------
    '-- Detalle
    '---------------------------------------------------------------
    arrCampos = Array("Código", "Descripción", "Laboratorio", _
                       "Línea", "Sel", "est_abast", _
                       "pvm_aprob", "Vnt", _
                       "pvm_actual", "Cant. Solic", "Stock", "Categoria Comercial")
    
    arrCaption = Array("Código", "Descripción", "Laboratorio", _
                       "Línea", "Sel", "est_abast", _
                       "pvm_aprob", "Vnt", _
                       "pvm_actual", "Cant. Solic", "Stock", "Categoria Comercial")
    
    arrAncho = Array(650, 3800, 1500, _
                     1500, 600, 600, _
                     600, 600, _
                     600, 800, 800, 1500)
    
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, _
                          dbgLeft, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgLeft)
                              
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdDetalle.HeadLines = 2
    grdDetalle.RowHeight = 0
    grdDetalle.RowHeight = grdDetalle.RowHeight * 2
    grdDetalle.Columns("Vnt").ButtonText = True
    grdDetalle.Columns("Vnt").ButtonAlways = True
    grdDetalle.AllowUpdate = True
    
    grdDetalle.EditorStyle.BackColor = vbWhite 'RGB(242, 242, 252)
    grdDetalle.EditorStyle.ForeColor = RGB(180, 0, 180)
    grdDetalle.EditorStyle.Font.Bold = True
    
    'Columnas editables
    grdDetalle.Columns(9).BackColor = vbInfoBackground
    grdDetalle.Columns(9).DataWidth = 4
            
    grdDetalle.Array1 = solicitud.Detalle
    solicitud.Detalle.ReDim 0, -1, 0, 15
       
    grdDetalle.Columns("Descripción").FetchStyle = True
    grdDetalle.Columns("Sel").Visible = False
    grdDetalle.Columns("est_abast").Visible = False
    grdDetalle.Columns("pvm_aprob").Visible = False
    grdDetalle.Columns("pvm_actual").Visible = False
    grdDetalle.Columns("Vnt").Visible = False
End Sub





