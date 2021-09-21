VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_Entrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion de Mercaderia"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   ClipControls    =   0   'False
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbVerGuiasSob 
      Caption         =   "Ver Guias SOB (Rpta)"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "[Esc] Salir"
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "[F5] Conteo Productos"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdF1 
      Caption         =   "[F1] Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdF3 
      Caption         =   "[F3] Asociar Entregas"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5280
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFechaInicio 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59703297
         CurrentDate     =   40675
      End
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59703297
         CurrentDate     =   40675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Inicio"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Fin"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   11
         Top             =   330
         Width           =   390
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "[F9] Reimprimir Transportista"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "[F8] Ver Reportes"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   5880
      Width           =   1815
   End
   Begin vbp_Ventas.ctlGrilla grdRecepcion 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
   End
End
Attribute VB_Name = "frm_ADM_Entrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objEntrega As New clsEntrega
Dim permiso As Integer

Private Sub cmbVerGuiasSob_Click()
 frm_ADM_Guias_Pend_Sob.carga
 
End Sub

Private Sub cmdBuscar_Click()

 If dtpFechaInicio.Value > dtpFechaFin.Value Then
    MsgBox "La Fecha de Inicio no puede ser mayor a la fecha Final", vbExclamation, "Atención"
    dtpFechaInicio.SetFocus
    GoTo Salida
 End If
 Consulta
 grdRecepcion.SetFocus
Salida:
End Sub

Private Sub cmdEsc_Click()
    If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
       Unload Me
    End If
End Sub

Private Sub cmdF1_Click()
    Nuevo
End Sub

Private Sub cmdF3_Click()
On Error GoTo Control

If Me.grdRecepcion.Columns("COD_ESTADO") <> "CRC" Then
    If Me.grdRecepcion.Columns("NUM_GUIAS").Value = "0" Then
        If (Me.grdRecepcion.Columns("COD_ESTADO").Value <> "CRC") Then
            frm_ADM_TranspGuia.carga (Me.grdRecepcion.Columns("ID_ENTREGA").Value)
        Else
            MsgBox "Para agregar guías el estado debe ser distinto de AFECT. TOTAL", vbInformation, "Aviso"
        End If
    Else
        frm_ADM_GuiaAsociada.strIdEntrega = Me.grdRecepcion.Columns("ID_ENTREGA").Value
        frm_ADM_GuiaAsociada.Show vbModal
    End If
Else
    MsgBox "No se puede Asociar Guía, es estado es AFECT. TOTAL", vbCritical + vbInformation, "Aviso"
End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdF5_Click()
On Error GoTo Control
    If permiso = 0 Then
            If Me.grdRecepcion.Columns("NUM_GUIAS").Value <> "0" Then
                If Me.grdRecepcion.Columns("COD_ESTADO").Value = "EMI" Then
                    frm_ADM_Conteo1.strIdEntrega = Me.grdRecepcion.Columns("ID_ENTREGA").Value
                    frm_ADM_Conteo1.Caption = "Conteo - Entrega Nº " & Me.grdRecepcion.Columns("ID_ENTREGA").Value
                    frm_ADM_Conteo1.Show vbModal
                Else
                    MsgBox "Solo tiene permisos para efectuar el primer conteo", vbInformation, "Aviso"
                End If
            Else
                MsgBox "Debe asociar guías a la entrega", vbCritical, "Aviso"
            End If
    Else
        If Me.grdRecepcion.Columns("NUM_GUIAS").Value <> "0" Then
            If Me.grdRecepcion.Columns("COD_ESTADO").Value = "EMI" Then
              frm_ADM_Conteo1.strIdEntrega = Me.grdRecepcion.Columns("ID_ENTREGA").Value
              frm_ADM_Conteo1.Caption = "Conteo - Entrega Nº " & Me.grdRecepcion.Columns("ID_ENTREGA").Value
              frm_ADM_Conteo1.Show vbModal
            ElseIf Me.grdRecepcion.Columns("COD_ESTADO").Value = "CER" Then
              frm_ADM_Conteo2.strEntrega = Me.grdRecepcion.Columns("ID_ENTREGA").Value
              frm_ADM_Conteo2.carga "" & Me.grdRecepcion.Columns("ID_ENTREGA").Value, "1"
            End If
        Else
            MsgBox "Debe asociar guías a la entrega", vbCritical, "Aviso"
        End If
    End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Command1_Click()
On Error GoTo Control

    If (Me.grdRecepcion.Columns("COD_ESTADO").Value = "CRC") Then
        frm_ADM_Sobrantes.idEntrega = Me.grdRecepcion.Columns("ID_ENTREGA").Value
        frm_ADM_Sobrantes.Show
    Else
        MsgBox "Para ver el reporte el estado debe ser AFECT. TOTAL", vbInformation, "Aviso"
    End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Command2_Click()
  frm_ADM_Guias_Pend_Sob.carga
  
  
End Sub

Private Sub Command3_Click()
On Error GoTo Control

    If Me.grdRecepcion.Columns("COD_ESTADO").Value <> "CRC" Then
        fnImprimeTransportista ("" & Me.grdRecepcion.Columns("ID_ENTREGA").Value)
    Else
        MsgBox "No se Puede Inprimir debido a que el estado es AFECT. TOTAL", vbCritical + vbInformation, "Aviso"
    End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Sub justifica_printer(x0, xf, y0, txt)
' x0, xf = posicion de los margenes izquierdo y derecho
' y0 = posicion vertical donde se desea empezar a escribir
' txt = texto a escribir

Dim x, Y, k, Ancho
Dim s As String, ss As String
Dim x_spc

s = txt
x = x0
Y = y0
Ancho = (xf - x0)

While s <> ""

ss = ""
While (s <> "") And (Printer.TextWidth(ss) <= Ancho)
ss = ss & left$(s, 1)
s = right$(s, Len(s) - 1)
Wend
If (Printer.TextWidth(ss) > Ancho) Then
s = right$(ss, 1) & s
ss = left$(ss, Len(ss) - 1)
End If
' aqui tenemos en ss lo maximo que cabe en una linea
If right$(ss, 1) = " " Then
ss = left$(ss, Len(ss) - 1)
Else
If (InStr(ss, " ") > 0) And (left$(s & " ", 1) <> " ") Then
While right$(ss, 1) <> " "
s = right$(ss, 1) & s
ss = left$(ss, Len(ss) - 1)
Wend
ss = left$(ss, Len(ss) - 1)
End If
End If
x_spc = 0
x = x0
If (Len(ss) > 1) And (s & "" <> "") Then
x_spc = (Ancho - Printer.TextWidth(ss)) / (Len(ss) - 1)
End If
Printer.CurrentX = x
Printer.CurrentY = Y

If x_spc = 0 Then
Printer.Print ss;
Else
For k = 1 To Len(ss)
Printer.CurrentX = x
Printer.Print Mid$(ss, k, 1);
x = x + Printer.TextWidth("*" & Mid$(ss, k, 1) & "*") - Printer.TextWidth("**")
x = x + x_spc
Next
End If

Y = Y + Printer.TextHeight(ss)
While left$(s, 1) = " "
s = right$(s, Len(s) - 1)
Wend
Wend

End Sub

Public Function fnImprimeTransportista(strIdEntrega As String)
    
    Dim rs As oraDynaset

    Set rs = objEntrega.ImprimeTransportista(strIdEntrega)
    Printer.Font.Size = 10
    Printer.Print Space(6) & "CONSTANCIA DE TRANSPORTISTA"
    Printer.Print ""
    Printer.Print Space(20) & rs("COD_LOCAL") & " - " & rs("DES_LOCAL")
    Printer.Print ""
    Printer.Print "Fecha : " & rs("FCH_REGISTRA")
    Printer.Print ""
    Dim str As String
    str = "Yo, " & rs("DES_NOMBRE") & " " & rs("APE_PAT_USUARIO") & " " & rs("APE_MAT_USUARIO") & ", Adm. del Local, " & " CONFIRMO LA RECEPCION DE LA MERCADERIA " & _
    "entregada por el Sr. Transportista " & rs("DES_CHOFER") & ", en la unidad con Placa: " & rs("DES_PLACA") & "."
    justifica_printer 20, 4000, Printer.CurrentY, str
    str = "La recepcion Consta de: " & rs("CTD_BULTOS") & " Bulto(s) y " & rs("CTD_PRECINTOS") & " Precinto(s)."
    Printer.Print ""
    justifica_printer 20, 4000, Printer.CurrentY, str
    str = "GLOSA : " & rs("DES_GLOSA") & "."
    Printer.Print ""
    justifica_printer 20, 4000, Printer.CurrentY, str
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print "Firma Transportista: __________________"
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print "Firma Adm. Local: ____________________"
    Printer.EndDoc
End Function


'Private Sub dtpFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        cmdBuscar.SetFocus
'    End If
'End Sub
'
'Private Sub dtpFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        dtpFechaFin.SetFocus
'    End If
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control

If permiso = 1 Then
    If KeyCode = vbKeyF1 Then
        Nuevo
    End If
    If KeyCode = vbKeyF8 Then
        Command1_Click
    End If
    If KeyCode = vbKeyF3 Then
        cmdF3_Click
    End If
    If KeyCode = vbKeyF9 Then
        Command3_Click
    End If
    If KeyCode = 13 Or KeyCode = vbKeyF5 Then
        cmdF5_Click
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
ElseIf permiso = 0 Then
    If KeyCode = 13 Or KeyCode = vbKeyF5 Then
        If Me.grdRecepcion.Columns("NUM_GUIAS").Value <> "0" Then
            If Me.grdRecepcion.Columns("COD_ESTADO").Value = "EMI" Then
                frm_ADM_Conteo1.strIdEntrega = Me.grdRecepcion.Columns("ID_ENTREGA").Value
                frm_ADM_Conteo1.Caption = "Conteo - Entrega Nº " & Me.grdRecepcion.Columns("ID_ENTREGA").Value
                frm_ADM_Conteo1.Show vbModal
            Else
                MsgBox "Solo tiene permisos para efectuar el primer conteo", vbInformation, "Aviso"
            End If
        Else
            MsgBox "Debe asociar guías a la entrega", vbCritical, "Aviso"
        End If
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Public Sub Form_Load()
On Error GoTo handle
    permiso = 0
    HabilitaPermisos
    If permiso = 1 Then
        Me.cmdEsc.Enabled = True
        Me.cmdF1.Enabled = True
        Me.cmdF3.Enabled = True
        Me.cmdF5.Enabled = True
        Me.Command1.Enabled = True
        Me.Command3.Enabled = True
    Else
        Me.cmdEsc.Enabled = True
        Me.cmdF1.Enabled = False
        Me.cmdF3.Enabled = False
        Me.cmdF5.Enabled = True
        Me.Command1.Enabled = False
        Me.Command3.Enabled = False
    End If
    dtpFechaInicio.Value = objUsuario.sysdate - 15
    dtpFechaFin.Value = objUsuario.sysdate
    Consulta
    SeteaGrilla
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

'Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
'On Error GoTo Handle
'
'Select Case Index
'    Case 1
'        Nuevo
''    Case 2
''        edita
'    Case 2
'        Consulta
''    Case 4
''        Actualiza
'    Case 3
'        grdRecepcion.MostrarImprimir
'    Case 4
'        grdRecepcion.MostrarExcel
'    Case 5
'        grdRecepcion.MostrarEmail
'    Case 6
'        Unload Me
'    Case Else
'        MsgBox "Esta opción no esta implementada", vbCritical, App.ProductName
'End Select
'Exit Sub
'Handle:
'    MsgBox Err.Description, vbCritical, App.ProductName
'End Sub

Sub Nuevo()
On Error GoTo Control
    frm_ADM_Transportista.carga ""
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Sub edita()
If Me.grdRecepcion.ApproxCount > 0 Then
        frm_ADM_EntregaDet.carga "" & grdRecepcion.Columns("ID_ENTREGA").Value
        Actualiza
Else
    MsgBox "No se encontraron Items en la Grilla.", vbCritical, App.ProductName
End If


End Sub

Public Sub Consulta()
On Error GoTo Control
    Set grdRecepcion.DataSource = objEntrega.Lista(objUsuario.CodigoLocal, dtpFechaInicio.Value, dtpFechaFin.Value, "", "")
    cargaDetalle
    'grdRecepcion.SetFocus
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Sub Actualiza()
    Consulta
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant

    arrCampos = Array("ID_ENTREGA", "FCH_REGISTRA", "DES_USUARIO", "NUM_GUIAS", "COD_ESTADO", "DES_CODIGO")
    arrCaption = Array("Nº Ingreso", "Fecha Ingreso", "Usuario Creación", "Cant. Guías", "Estado", "Estado")
    arrAncho = Array(1200, 1200, 3500, 1000, 1200, 1500)
    arrAlineacion = Array(dbgRight, dbgCenter, dbgLeft, dbgRight, dbgLeft, dbgLeft)
    grdRecepcion.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Me.grdRecepcion.Columns(4).Visible = False
End Sub

'Private Sub grdRecepcion_RegistroSeleccionado(ByVal DatoColumna0 As String)
'    cargaDetalle
'End Sub

Private Sub cargaDetalle()
'    If grdRecepcion.ApproxCount > 0 Then
'        If grdRecepcion.Columns("COD_ESTADO").Value <> "EMI" Then
'            Set grdRecepcionDET.DataSource = objEntrega.ListaDetalle("" & grdRecepcion.Columns("ID_ENTREGA").Value)
'            SeteaGrillaDet
'        Else
'            grdRecepcionDET.Limpiar
'        End If
'    Else
'         grdRecepcionDET.Limpiar
'    End If
End Sub

Private Sub SeteaGrillaDet()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant

    arrCampos = Array("NUM_GUIA", "NUM_ITEM", "DES_PRODUCTO")
    arrCaption = Array("Num_Guia", "Item", "Producto")
    arrAncho = Array(1200, 800, 7000)
    arrAlineacion = Array(dbgLeft, dbgCenter, dbgLeft)
    'grdRecepcionDET.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objEntrega = Nothing
End Sub

Sub HabilitaPermisos()
On Error GoTo Control

    Dim objPermisos As New clsAutorizacion
    Dim rsPermisos As oraDynaset
    Set rsPermisos = objPermisos.ListaPermisos(objUsuario.Aplicacion, objUsuario.Codigo, "001")
    rsPermisos.MoveFirst
    While Not rsPermisos.EOF
         If rsPermisos("COD_MENU") = "104" Then
            permiso = 1
            Set objPermisos = Nothing
            Exit Sub
         End If
         rsPermisos.MoveNext
    Wend
    Set objPermisos = Nothing
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub
