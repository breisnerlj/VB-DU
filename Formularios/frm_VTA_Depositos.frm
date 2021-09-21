VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_VTA_Depositos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registo de Depósitos"
   ClientHeight    =   7905
   ClientLeft      =   2265
   ClientTop       =   2250
   ClientWidth     =   10170
   Icon            =   "frm_VTA_Depositos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10170
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   6120
      TabIndex        =   16
      Top             =   6000
      Width           =   2175
      Begin VB.Label Label5 
         Caption         =   "<Ctrl><Del> Elimina Registro"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "<Insert> Adiciona Registro"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
   End
   Begin vbp_Ventas.ctlGrillaArray grdDepositos 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8070
   End
   Begin MSComCtl2.DTPicker dtpFchDeposito 
      Height          =   315
      Left            =   7800
      TabIndex        =   3
      Top             =   2445
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65142785
      CurrentDate     =   40080
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   8640
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones"
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   4800
      Width           =   6015
      Begin VB.TextBox txtObservaciones 
         Height          =   735
         Left            =   120
         MaxLength       =   199
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   5775
      End
   End
   Begin vbp_Ventas.ctlTextBox txtImporte 
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Tipo            =   4
      Alignment       =   2
      MaxLength       =   8
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
   Begin vbp_Ventas.ctlTextBox txtNroOperacion 
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Tipo            =   4
      Alignment       =   2
      MaxLength       =   10
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
   Begin vbp_Ventas.ctlDataCombo cboBancos 
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlDataCombo cboCtaCte 
      Height          =   315
      Left            =   6360
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlGrillaArray ctlOpciones 
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3201
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label lblADepositar 
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4440
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblFchDep 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Despósito:"
      Height          =   195
      Left            =   6360
      TabIndex        =   14
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label lblSoles 
      AutoSize        =   -1  'True
      Caption         =   "Importe :"
      Height          =   195
      Left            =   6480
      TabIndex        =   12
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblNroOpe 
      Caption         =   "N° Operación :"
      Height          =   435
      Left            =   6480
      TabIndex        =   11
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label lblEntFin 
      AutoSize        =   -1  'True
      Caption         =   "Entidad Financiera"
      Height          =   195
      Left            =   6360
      TabIndex        =   10
      Top             =   240
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblctacte 
      AutoSize        =   -1  'True
      Caption         =   "Cta Cte."
      Height          =   195
      Left            =   6360
      TabIndex        =   9
      Top             =   1320
      Width           =   570
   End
End
Attribute VB_Name = "frm_VTA_Depositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ixdbDepositos As New XArrayDB
Private lintFila As Integer
Dim objDepositos As New clsDepositos
Dim objRemito As New clsRemito
Dim vImporte, dblImporte, dblDiferencia, TotalDif  As Double
Dim strMoneda As String
Dim objOpciones As New clsOpciones
Dim lodynConsultaArea As oraDynaset

Public godbOraDatabase As OraDatabase
Dim lxdbOpcion As New XArrayDB
Dim strCadCodDif As String
Dim strCadMonDif As String



Private Sub cboBancos_Change()
On Error GoTo Control
    Set cboCtaCte.RowSource = objDepositos.ListaCtaCte(cboBancos.BoundText)
        cboCtaCte.BoundColumn = "COD"
        cboCtaCte.ListField = "DES"
        cboCtaCte.BoundText = "*"
   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cboCtaCte_Change()
Dim rstDespositos As oraDynaset

On Error GoTo Control
    
    strMoneda = ""
    ixdbDepositos.ReDim 0, -1, 0, 5
    grdDepositos.Array1 = ixdbDepositos
    grdDepositos.Rebind
    
    If cboCtaCte.BoundText = "*" Then
        Exit Sub
    Else
    
        strMoneda = IIf(Mid(cboCtaCte.BoundText, 5, 5) = "S", "1", "2")
        
        Select Case strMoneda
            
            Case "1"
                 Set rstDespositos = objDepositos.ListaRemitos(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, _
                                                               strMoneda, gclsOracle.Fecha_Servidor)
            Case "2"
                 Set rstDespositos = objDepositos.ListaRemitos(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, _
                                                               strMoneda, gclsOracle.Fecha_Servidor)
    
        End Select
    
        If rstDespositos.RecordCount > 0 Then
            lintFila = 0
            ixdbDepositos.LoadRows rstDespositos.GetRows
        End If
    End If

    grdDepositos.Rebind
    grdDepositos.MoveFirst

   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdGrabar_Click()
Dim strMensaje As String
Dim dblDif, dblHoy, dblFecha As Double
Dim larrRemito() As String
Dim larrImpRemito() As String
Dim i, vconta, conta As Integer

On Error GoTo Control
TotalDif = 0#
strCadMonDif = ""
strCadCodDif = ""

    dblFecha = Val(Format(dtpFchDeposito.Value, "YYYYMMDD"))
    dblHoy = Val(Format(Now, "YYYYMMDD"))
    vconta = 0
    'MODIFICADO POR MLEVANO 16/11/2012
                        For i = 0 To ctlOpciones.ApproxCount - 1
                            If objOpciones.Opcionesxdb(i, 1) <> "" Then
                                TotalDif = TotalDif + objOpciones.Opcionesxdb(i, 1)
                                strCadCodDif = strCadCodDif & objOpciones.Opcionesxdb(i, 0) & "|"
                                strCadMonDif = strCadMonDif & objOpciones.Opcionesxdb(i, 1) & "|"
                            End If
                        Next
    If grdDepositos.ApproxCount < 1 And vImporte = "" Then
       MsgBox "Debe seleccionar el o los remitos a grabar", vbCritical, "Error": Exit Sub
       ElseIf cboCtaCte.BoundText = "*" Then
              MsgBox "Debe seleccionar el Nro de Cuenta Corriente", vbCritical, "Error": Exit Sub
              ElseIf dblFecha > dblHoy Then
                     MsgBox "La fecha del depósito no puede ser mayor al día de hoy.", vbCritical, "Error"
                     dtpFchDeposito.Value = gclsOracle.Fecha_Servidor
                     Exit Sub
                     ElseIf txtNroOperacion.Text = "" Then
                            MsgBox "Debe ingresar el Nro de Operación", vbCritical, "Error": Exit Sub
                            ElseIf Val(lblADepositar.Caption) = 0 Then
                                MsgBox "Debe seleccionar un remito", vbCritical, "Error": Exit Sub
                                ElseIf txtImporte.Text = "0.00" Or txtImporte.Text = "0" Then
                                       MsgBox "El importe a grabar no puede ser 0", vbCritical, "Error": Exit Sub
                                       ElseIf vImporte < (TotalDif + Val(txtImporte.Text)) Then
                                            MsgBox "El importe a depositar no puede ser menor que la suma del importe depositado más el importe no depositado", vbCritical, "Error": Exit Sub
                                            ElseIf vImporte > (TotalDif + Val(txtImporte.Text)) Then
                                               MsgBox "El importe a depositar no puede ser mayor que la suma del importe depositado más el importe no depositado", vbCritical, "Error": Exit Sub
    
    End If
    
    If TotalDif <> 0 Then
        If MsgBox("El total del importe no depositado es de " & TotalDif & " ¿Desea continuar?", vbYesNo + vbCritical, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
        
'MsgBox "Grabo", vbCritical, "Aviso": Exit Sub


    ReDim larrRemito(0 To 0)
    ReDim larrImpRemito(0 To 0)
    
    ReDim larrDifCod(0 To 0)
    ReDim larrDifMonto(0 To 0)

        For i = ixdbDepositos.LowerBound(1) To ixdbDepositos.UpperBound(1)
            If Abs(Val(ixdbDepositos(i, 4))) = 1 Then
                vconta = vconta + 1

                larrRemito(UBound(larrRemito)) = ixdbDepositos(i, 0)
                
                If strMoneda = 1 Then
                    larrImpRemito(UBound(larrImpRemito)) = ixdbDepositos(i, 2)
                Else
                    larrImpRemito(UBound(larrImpRemito)) = ixdbDepositos(i, 3)
                End If
                ReDim Preserve larrRemito(UBound(larrRemito) + 1)
                ReDim Preserve larrImpRemito(UBound(larrImpRemito) + 1)
            End If
        Next i
        
        i = 0
        For i = objOpciones.Opcionesxdb.LowerBound(1) To objOpciones.Opcionesxdb.UpperBound(1)
            'If Abs(Val(objOpciones.Opcionesxdb(i, 1))) <> 0 Then
                conta = conta + 1

                larrDifCod(UBound(larrDifCod)) = objOpciones.Opcionesxdb(i, 0)
                
                larrDifMonto(UBound(larrDifMonto)) = objOpciones.Opcionesxdb(i, 1)
                
                ReDim Preserve larrDifCod(UBound(larrDifCod) + 1)
                ReDim Preserve larrDifMonto(UBound(larrDifMonto) + 1)
            'End If
        Next i


        If vconta = 0 Then
            MsgBox "Debe seleccionar el o los remitos a grabar", vbCritical, Caption
            Exit Sub
        End If
    
    ReDim Preserve larrRemito(UBound(larrRemito) - 1)
    ReDim Preserve larrImpRemito(UBound(larrImpRemito) - 1)
    
    ReDim Preserve larrDifCod(UBound(larrDifCod) - 1)
    ReDim Preserve larrDifMonto(UBound(larrDifMonto) - 1)

    'dblDiferencia = dblImporte - Val(txtImporte.Text)
    dblDiferencia = Val(lblADepositar.Caption) - dblImporte
    
    If MsgBox("¿Seguro(a) de Grabar el Deposito?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        Exit Sub
    End If

'    strMensaje = objDepositos.Graba(objUsuario.CodigoEmpresa, _
'                                    objUsuario.CodigoLocal, _
'                                    Mid(cboCtaCte.BoundText, 1, 3), _
'                                    Trim(txtNroOperacion.Text), _
'                                    dblImporte, _
'                                    dblDiferencia, _
'                                    CStr(Format(dtpFchDeposito.Value, "dd/mm/yyyy")), _
'                                    Trim(txtObservaciones.Text), _
'                                    objUsuario.Codigo, _
'                                    strMoneda, _
'                                    larrRemito, larrImpRemito, _
'                                    objUsuario.CodigoLocal)
'MODIFICADO POR MLEVANO 16/11/2012
    strMensaje = objDepositos.Graba(objUsuario.CodigoEmpresa, _
                                    objUsuario.CodigoLocal, _
                                    Mid(cboCtaCte.BoundText, 1, 3), _
                                    Trim(txtNroOperacion.Text), _
                                    dblImporte, _
                                    dblDiferencia, _
                                    CStr(Format(dtpFchDeposito.Value, "dd/mm/yyyy")), _
                                    Trim(txtObservaciones.Text), _
                                    objUsuario.Codigo, _
                                    strMoneda, _
                                    larrRemito, larrImpRemito, _
                                    objUsuario.CodigoLocal, larrDifCod, larrDifMonto)

   If strMensaje = "" Then
        MsgBox "Se grabo satisfactoriamente", vbExclamation, App.ProductName
        Unload Me
   Else
        MsgBox strMensaje, vbCritical, App.ProductName
   End If

   Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim strBanco As String
   On Error GoTo Control

    IniciaArray
    SetDepositos
    SetOpciones
    
    txtImporte.Text = "0.00"

    strBanco = objRemito.BancoxBtl(objUsuario.CodigoLocal)
    dtpFchDeposito.Value = gclsOracle.Fecha_Servidor

    Set cboBancos.RowSource = objDepositos.ListaBancos
        cboBancos.BoundColumn = "COD"
        cboBancos.ListField = "DES"
        cboBancos.BoundText = Trim(Mid(strBanco, 1, 3))
    
    Set cboCtaCte.RowSource = objDepositos.ListaCtaCte(cboBancos.BoundText, objUsuario.CodigoEmpresa)
        cboCtaCte.BoundColumn = "COD"
        cboCtaCte.ListField = "DES"
        cboCtaCte.BoundText = "*"
   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objDepositos = Nothing
    Set ixdbDepositos = Nothing
End Sub

Private Sub SetDepositos()

  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim i As Integer
  
    arrCampos = Array("REMITO", "FCH_REMITO", "IMP_SOLES", "IMP_DOLARES", "")

    arrCaption = Array("Remito", "Fch.Registro", "Total S/.", "Total $", "CHK")

    arrAncho = Array(1100, 1500, 900, 900, 800)

    arrAlineacion = Array(dbgCenter, dbgCenter, dbgRight, dbgRight, dbgCenter)

    grdDepositos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    grdDepositos.AllowUpdate = True

    For i = 0 To 3
        grdDepositos.Columns(i).AllowFocus = False
    Next i
    
    grdDepositos.Columns(0).Merge = True
    grdDepositos.Columns(1).Merge = True
    grdDepositos.Columns(4).ValueItems.Presentation = dbgCheckBox

End Sub

Public Sub IniciaArray()
On Error GoTo Control
    
    ixdbDepositos.ReDim 0, -1, 0, 5
    grdDepositos.Array1 = ixdbDepositos
    grdDepositos.Rebind

    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdDepositos_AfterColUpdate(ByVal ColIndex As Integer)
Dim i%

vImporte = 0

On Error GoTo Control
   
    Select Case ColIndex
        Case 4
            grdDepositos.MoveNext
            grdDepositos.MovePrevious

            For i = 0 To ixdbDepositos.UpperBound(1)
               Select Case strMoneda
                    Case "*"
                        txtImporte.Text = "0.00"
                    Case "1"
                        If ixdbDepositos(i, 4) = "-1" Then
                            vImporte = vImporte + ixdbDepositos(i, 2)
                        End If
                        
                        txtImporte.Text = IIf(vImporte = 0, "0.00", vImporte)
                        lblADepositar.Caption = IIf(vImporte = 0, "0.00", vImporte)
                    Case "2"
                        If ixdbDepositos(i, 4) = "-1" Then
                            vImporte = vImporte + ixdbDepositos(i, 3)
                        End If
                        txtImporte.Text = vImporte
                        lblADepositar.Caption = vImporte
               End Select
            Next i
    End Select
  
   Exit Sub

Control:
    MsgBox Err.Description, vbCritical, "Error : " & Err.Number
End Sub

Private Sub txtObservaciones_GotFocus()
    txtObservaciones.BackColor = txtImporte.ColorFoco
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtObservaciones_LostFocus()
    txtObservaciones.BackColor = txtImporte.ColorDefault
End Sub

Private Sub txtImporte_Change()
On Error GoTo Control
    If Not IsNumeric(txtImporte.Text) Then Exit Sub
        dblImporte = txtImporte.Text
   Exit Sub
Control:
    MsgBox Err.Description, vbCritical, "Error : " & Err.Number
End Sub

Sub SetOpciones()
Dim varCampos As Variant
Dim varTitulo As Variant
Dim varAlinea As Variant
Dim varAncho As Variant
Dim varCampoDato As Variant

    varCampos = Array("", "")
    varAncho = Array(1500, 1000)
    varTitulo = Array("Tipo", "Importe")
    varAlinea = Array(2, 2)
    
    ctlOpciones.FormatoGrilla varCampos, varTitulo, varAncho, varAlinea
    
    
    objOpciones.Inicializa
    ctlOpciones.Array1 = objOpciones.Opcionesxdb
    ctlOpciones.AllowUpdate = True
    ctlOpciones.Rebind
 
    ctlOpciones.Columns("Importe").NumberFormat = "#0.0#"
    ctlOpciones.EditorStyle.BackColor = &H80FF&
    ctlOpciones.RowHeight = 500
    
     'COMBO DENTRO DE LA GRILLA'
    Dim odynOpciones As oraDynaset
    Dim strSQL$

'    strSQL = "SELECT '0' AS COD_OPCION,'Dinero Falso' AS DES_OPCION FROM dual UNION ALL " & _
'             "SELECT '1' AS COD_OPCION,'Robo' AS DES_OPCION FROM dual UNION ALL " & _
'             "SELECT '2' AS COD_OPCION,'Deficit de Quimico' AS DES_OPCION FROM dual"

'    strSQL = "select * from cmr.mae_producto_com where cod_producto='77299'"
'    Set odynOpciones = gclsOracle.ODataBase.CreateDynaset(strSQL, 0&)
'    MsgBox odynOpciones("COD_PRODUCTO")

'    strSQL = "SELECT COD_DIFERENCIA, DES_DIFERENCIA FROM BTLPROD.REL_DIFERENCIA where cod_diferencia='1'"
'    Set odynOpciones = gclsOracle.ODataBase.CreateDynaset(strSQL, 0&)
'    MsgBox odynOpciones("COD_DIFERENCIA")


'MODIFICADO POR MLEVANO 16/11/2012
    strSQL = "SELECT COD_DIFERENCIA AS COD_OPCION, DES_DIFERENCIA AS DES_OPCION FROM BTLPROD.REL_DIFERENCIA WHERE FLG_ESTADO='1'"
    Set odynOpciones = gclsOracle.ODataBase.CreateDynaset(strSQL, 0&)
    If Not odynOpciones.EOF Then
        lxdbOpcion.Clear
        lxdbOpcion.ReDim 0, odynOpciones.RecordCount - 1, 0, 10
        Dim i As Integer
        i = 0
        odynOpciones.MoveFirst
        While Not odynOpciones.EOF
            lxdbOpcion(i, 0) = odynOpciones(0).Value
            lxdbOpcion(i, 1) = odynOpciones(1).Value
            odynOpciones.MoveNext
            i = i + 1
        Wend
    End If
    
    Call spGrilla_CboBox(ctlOpciones, "Tipo", "COD_OPCION", odynOpciones, "DES_OPCION")
    objOpciones.Opcionesxdb.AppendRows
    ctlOpciones.Rebind
 End Sub
 
 Private Sub ctlOpciones_AfterColUpdate(ByVal ColIndex As Integer)
    ctlOpciones.Rebind
    Select Case ColIndex
        Case 0
            SendKeys "{Right}"
    End Select
End Sub


Private Sub ctlOpciones_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
 'v_Valida = True

 
 Dim fila As Integer
 fila = ctlOpciones.Bookmark
            Select Case ColIndex
                Case 0  'Tipo
                       If FN_Buscar_Opcion(Trim(ctlOpciones.Columns(0)), fila) <> "" Then
                            objOpciones.Opcionesxdb(fila, 0) = Trim(ctlOpciones.Columns(0).Value)
                       End If
                 Case 1 'Monto
                        If (Not IsNumeric(ctlOpciones.Columns(1))) Or Not (Val(ctlOpciones.Columns(1)) > 0) Then
                                'v_Valida = False
                                Cancel = True
                        Else
                                objOpciones.Opcionesxdb(fila, 1) = Trim(ctlOpciones.Columns(1))
                        End If
            End Select
            
End Sub

Private Sub ctlOpciones_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
'Dim i As Integer
'i = 0
'TotalDif = 0
'strCadMonDif = ""
'strCadCodDif = ""
    Select Case KeyCode
        Case vbKeyDelete
            If Shift = 2 Then    '<ctrl> <delete>
                If MsgBox("¿ Esta seguro de eliminar ?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
                    If ctlOpciones.ApproxCount > 1 Then
                        ctlOpciones.Delete
                        
'                        For i = 0 To ctlOpciones.ApproxCount - 1
'                                TotalDif = TotalDif + objOpciones.Opcionesxdb(i, 1)
'                                strCadCodDif = strCadCodDif & objOpciones.Opcionesxdb(i, 0) & "|"
'                                MsgBox strCadCodDif
'                                strCadMonDif = strCadMonDif & objOpciones.Opcionesxdb(i, 1) & "|"
'                                MsgBox strCadMonDif
'                        Next
                    
                    ElseIf ctlOpciones.ApproxCount = 1 Then
                        ctlOpciones.Delete
                        objOpciones.Opcionesxdb.ReDim 0, 0, 0, 3
                        ctlOpciones.Rebind
                    End If
                End If
            End If
        Case vbKeyInsert
            'Validar Columnas
            ctlOpciones.Col = 1
            If FN_Validar_Datos Then
               ctlOpciones.Update
               objOpciones.Opcionesxdb.AppendRows
               ctlOpciones.Rebind
               ctlOpciones.MoveLast
               ctlOpciones.Col = 0
            End If
        Case vbKeyReturn    'siguiente columna
            ctlOpciones.Columns(1).Text = Trim(ctlOpciones.Columns(1))
    End Select
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub


Private Function FN_Validar_Datos() As Boolean

    FN_Validar_Datos = True
  
    If ctlOpciones.Columns(0) = "" Then
            MsgBox "Falta Seleccionar la Opcion", vbCritical + vbOKOnly, App.ProductName
            FN_Validar_Datos = False
            ctlOpciones.Refresh
            Exit Function
    End If
    
     If ctlOpciones.Columns(1) = "" Or ctlOpciones.Columns(1) = "0.0" Then
            MsgBox "Falta Ingresar el Monto", vbCritical + vbOKOnly, App.ProductName
            FN_Validar_Datos = False
            ctlOpciones.Refresh
            Exit Function
    End If
    
   If (Not IsNumeric(ctlOpciones.Columns(1))) Or Not (Val(ctlOpciones.Columns(1)) > 0) Then
            MsgBox "El Monto debe ser Mayor que Cero", vbCritical + vbOKOnly, App.ProductName
            FN_Validar_Datos = False
            ctlOpciones.Refresh
            Exit Function
    End If
    
End Function

Private Function FN_Buscar_Opcion(ByVal v_opcion As String, ByVal fila As Integer) As String
     Dim intValor As Integer
     Dim m As Integer
    
   On Error GoTo handle
   
        For m = lxdbOpcion.LowerBound(1) To lxdbOpcion.UpperBound(1)
            If lxdbOpcion(m, 1) = v_opcion Then
               v_opcion = lxdbOpcion(m, 0)
               GoTo salir
            End If
        Next m
salir:
        intValor = objOpciones.Opcionesxdb.Find(0, 0, v_opcion)
        If intValor = -1 Then
            FN_Buscar_Opcion = "Encontro"
        Else
            MsgBox "La opción ya se encuentra Seleccionado Anteriormente", vbCritical + vbOKOnly, App.ProductName
            FN_Buscar_Opcion = ""
            Exit Function
        End If
    Exit Function
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Function

