VERSION 5.00
Begin VB.Form frm_VTA_CobroXResponsabilidad 
   BorderStyle     =   0  'None
   Caption         =   "Recetario Magistral"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   Icon            =   "frm_VTA_CobroXResponsabilidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlDataCombo dbcMotivo 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5685
      Picture         =   "frm_VTA_CobroXResponsabilidad.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Grabar Cobro"
      Height          =   615
      Left            =   4320
      Picture         =   "frm_VTA_CobroXResponsabilidad.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5415
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Width           =   6615
      Begin VB.Frame Frame3 
         Caption         =   "Usuarios"
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   6375
         Begin vbp_Ventas.ctlTextBox txtCriterio 
            Height          =   375
            Left            =   3960
            TabIndex        =   3
            ToolTipText     =   "Usted puede ayudarse buscando por codigo, nombre o centro de costo"
            Top             =   200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Tipo            =   2
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
         Begin VB.OptionButton Opt0 
            Caption         =   "No asignados al local"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton Opt0 
            Caption         =   "Asignados al local"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000018&
            Caption         =   "Los criterios de búsqueda pueden ser por apellido, local o código."
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   5895
         End
      End
      Begin VB.CheckBox chkProrretear 
         Caption         =   "Prorretear"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdFf 
         Caption         =   "Selecc"
         Height          =   615
         Left            =   5760
         Picture         =   "frm_VTA_CobroXResponsabilidad.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdRr 
         Caption         =   "Quita"
         Height          =   615
         Left            =   5760
         Picture         =   "frm_VTA_CobroXResponsabilidad.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   615
      End
      Begin vbp_Ventas.ctlTextBox txtImporte 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         ColorDefault    =   -2147483628
         ColorDefault    =   -2147483628
         Tipo            =   4
         Alignment       =   1
         Enabled         =   0   'False
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtImporteAFavor 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         ColorDefault    =   -2147483628
         ColorDefault    =   -2147483628
         Tipo            =   4
         Alignment       =   1
         Enabled         =   0   'False
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlTextBox txtImporteaCobrar 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         ColorDefault    =   -2147483628
         ColorDefault    =   -2147483628
         Tipo            =   4
         Alignment       =   1
         Enabled         =   0   'False
         Bloqueado       =   -1  'True
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
      Begin vbp_Ventas.ctlGrillaArray grdPersonalCobrar 
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2355
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlGrilla grdPersonalLocal 
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2143
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe a Cobrar :"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Favor :"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   555
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Neto a Cobrar :"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   915
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Datos"
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   6615
      Begin vbp_Ventas.ctlGrillaArray grdProductosCobro 
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2566
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   4560
         Picture         =   "frm_VTA_CobroXResponsabilidad.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin vbp_Ventas.ctlTextBox txtNumDocumento 
         Height          =   315
         Left            =   1920
         TabIndex        =   22
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Tipo            =   7
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
      Begin vbp_Ventas.ctlDataCombo dbcTipoDoc 
         Height          =   315
         Left            =   1920
         TabIndex        =   20
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Documento :"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Motivo del Cobro :"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   600
      Width           =   1290
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
      Index           =   5
      Left            =   6105
      TabIndex        =   16
      Top             =   6960
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
      Index           =   6
      Left            =   4320
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cobro por Responsabilidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   4
      Left            =   480
      TabIndex        =   14
      Top             =   120
      Width           =   2910
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frm_VTA_CobroXResponsabilidad.frx":1BBE
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frm_VTA_CobroXResponsabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento
Dim dblTotal As Double


Private Sub chkProrretear_Click()
Dim i As Integer

    On Error GoTo handle

                If grdPersonalCobrar.ApproxCount = 0 Then
                    grdPersonalLocal.SetFocus
                    Exit Sub
                End If
                
                If chkProrretear.Value = vbChecked Then
                
                    For i = 0 To objVenta.CobroResponsabilidad.UpperBound(1)
                        objVenta.AgregaCobroResponsabilidad grdPersonalCobrar.Columns(0), grdPersonalCobrar.Columns(1), Val(txtImporte.Text), chkProrretear.Value
                    Next i
                    grdPersonalCobrar.Columns(2).AllowFocus = False
                    
                Else
                    objVenta.AgregaCobroResponsabilidad grdPersonalCobrar.Columns(0), grdPersonalCobrar.Columns(1), grdPersonalCobrar.Columns(2).Value, chkProrretear.Value
                    grdPersonalCobrar.Columns(2).AllowFocus = True
                End If


                grdPersonalCobrar.Rebind
                CalFooter
                
    Exit Sub

handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdAceptar_Click()
Dim oTipDoc As OraParamArray
Dim oNumDoc As OraParamArray
Dim oTipDocCo As OraParamArray
Dim oNumDocCo As OraParamArray
Dim auxTipDoc As OraParamArray
Dim auxNumDoc As OraParamArray
Dim varMsgDoc As Variant
Dim varMsgDocCo As Variant
Dim objImpresion As New clsDocumento
Dim UltDocEmitido As oraDynaset
Dim strUltDocEmi As String
Dim oraMotivo As oraDynaset
Dim Importe As Double
Dim i As Integer
Dim RetPromoMensaje As String

    On Error GoTo CtrlErr


If dbcMotivo.BoundText = "" Then MsgBox "Debe seleccionar el motivo del cobro.", vbOKOnly + vbExclamation, "Error": dbcMotivo.SetFocus: Exit Sub
    
'    gclsOracle.Cerrar
'    If gclsOracle.Conexion(gvarTNSNAME, gvarUSUARIO, gvarPASSWORD) <> "" Then End
    Importe = Round(Val(txtImporte.Text), 2)
    
    If Round(dblTotal, 2) > Importe Then
        MsgBox "La suma de los importes asignados supera al monto neto por cobrar", vbInformation + vbOKOnly, App.ProductName
        Exit Sub
    End If

    objVenta.CodMotivoCobro = dbcMotivo.BoundText
    
    If objVenta.ptmModalidad = Cobro_Responsabilidad Then
    
        Set oraMotivo = objVenta.ListaMotivoCobro(objVenta.CodMotivoCobro)
        
        '** Validación Hecha el 23/08/2007 Por Crueda **'
        '** Modalidad Cobro x Responsabilidad         **'
        If objVenta.CodigoTipoVenta = Cobro_Responsabilidad And objVenta.CobroResponsabilidad.UpperBound(1) = -1 Then
            MsgBox "La modalidad de venta es Cobro por Responsabilidad , no ha seleccionado uno..", vbExclamation, App.ProductName
            Exit Sub
        End If
        '***********************************************'
        
        If Not oraMotivo.EOF Then
            If oraMotivo("FLG_DOCU_REFER").Value = "1" Then
                If frmPedido.grdPedido.ApproxCount <= 1 Then MsgBox "Ingrese el Producto con quien se realizara el cruce", vbCritical, Caption: Exit Sub
                If grdPersonalCobrar.ApproxCount <= 0 Then MsgBox "Ingrese a quien(s) se les realizara el descuento", vbCritical, Caption: Exit Sub
            End If
        Else
            MsgBox "Error al acceder a la lista Motivo de Cobro", vbCritical, App.ProductName
            Exit Sub
        End If
    
    End If
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Autor : Arturo Escate
    'Fecha : 10/11/2009
    'Proposito: Esto es para validar si necesita autorizacion previa
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim ObjValidacion As New clsAprobacion
    Dim strNumeroSolicitud As String
    Dim strAccion As String
    Dim strMensaje As String
    Dim strCodigoAutorizacion As String
    Dim srtCodigoAUTH As String
    Dim strStore As String
    srtCodigoAUTH = ""
valida:
'TxtImpTot.text
    Dim strCadCodigoProducto As String
    Dim strSubTotal As String
    Dim strCodigoUsuario As String
    Dim strCantidad As String
    Dim strCantidadfrac As String
    Dim e As Integer
    
    Dim CadenaPrecioUnitProducto As String
    Dim CadenaBaseImponible As String
    Dim CadenaImpuesto As String
    Dim CadenaExonerado As String
Dim strSubTotalUsu  As String
    
    e = 0
        strCadCodigoProducto = ""
        strCantidad = ""
        strCantidadfrac = ""
        CadenaPrecioUnitProducto = ""
        CadenaBaseImponible = ""
        CadenaImpuesto = ""
        CadenaExonerado = ""
        strSubTotal = ""
        strSubTotalUsu = ""
        strCodigoUsuario = ""
        
    While e < objVenta.Producto.Count(1)
        strCadCodigoProducto = strCadCodigoProducto & objVenta.Producto(e, 0) & "|"
        strCantidad = strCantidad & IIf(objVenta.Producto(e, 2) = "F", 0, objVenta.Producto(e, 3)) & "|"
        strCantidadfrac = strCantidadfrac & IIf(objVenta.Producto(e, 2) = "F", objVenta.Producto(e, 3), 0) & "|"
        CadenaPrecioUnitProducto = CadenaPrecioUnitProducto & "0|"
        CadenaBaseImponible = CadenaBaseImponible & "0|"
        CadenaImpuesto = CadenaImpuesto & "0|"
        CadenaExonerado = CadenaExonerado & "0|"
        strSubTotal = strSubTotal & objVenta.Producto(e, 4) & "|"
        e = e + 1
    Wend
    e = 0
    
    While e < objVenta.CobroResponsabilidad.Count(1)
        strSubTotalUsu = strSubTotalUsu & objVenta.CobroResponsabilidad(e, 2) & "|"
        strCodigoUsuario = strCodigoUsuario & objVenta.CobroResponsabilidad(e, 0) & "|"
        e = e + 1
    Wend
    
    If srtCodigoAUTH = "" Then frm_VTA_ObservaAutorizacion.Show vbModal

    strStore = ObjValidacion.Solicita("10", strAccion, strMensaje, srtCodigoAUTH, objUsuario.CodigoLocal, objUsuario.CodigoLiquidacion, _
                                      objVenta.CodigoCliente, objVenta.CodigoDocumentoVenta, objDocumento.ListaNumeroDisponible(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objVenta.CodigoDocumentoVenta), _
                                      objVenta.CodDocRef, objVenta.NumDocRef, Val(txtImporteaCobrar.Text), "", "", objUsuario.Codigo, frm_VTA_ObservaAutorizacion.OutObservacion, _
                                      strCodigoAutorizacion, "", "", "", "", "", "", "", "", "", "", strCadCodigoProducto, strCantidad, strCantidadfrac, CadenaPrecioUnitProducto, CadenaBaseImponible, CadenaImpuesto, CadenaExonerado, strSubTotal, strCodigoUsuario, strSubTotalUsu, "", dbcMotivo.BoundText, "", "", "", "", "", frm_VTA_ObservaAutorizacion.OutNumeroId)
    If Not strStore = "" Then
        MsgBox strStore, vbCritical, App.ProductName
        Exit Sub
    Else
        Select Case strAccion
            Case 0
                    MsgBox strMensaje, vbInformation, App.ProductName
            Case 1
                   MsgBox strMensaje, vbCritical, App.ProductName
                   Exit Sub
            Case 2
                   MsgBox strMensaje, vbInformation, App.ProductName
                   Exit Sub
            Case 3
                If MsgBox(strMensaje & Chr(13) & "¿Desea ingresar el codigo de autorización?", vbYesNo + vbInformation, App.ProductName) = vbYes Then
                    srtCodigoAUTH = frmAprobacion.Carga
                    If Not srtCodigoAUTH = "" Then
                        GoTo valida
                        Exit Sub
                    End If
                   Exit Sub
                Else
                    Exit Sub
                End If
            Case Else
                   MsgBox "no esta implementado", vbInformation, App.ProductName
                   Exit Sub
        End Select
    End If

If objVenta.GrabarDoc(gclsOracle.ODataBase, oTipDoc, oNumDoc, oTipDocCo, oNumDocCo, RetPromoMensaje) = False Then Exit Sub
    
        
    objVenta.OrdenaDoc gclsOracle.ODataBase, oTipDoc, oNumDoc, oTipDocCo, oNumDocCo, auxTipDoc, auxNumDoc
    
    
    
    For i = 0 To oTipDoc.ArraySize - 1
        If oTipDoc.get_Value(i) <> "" Then
        varMsgDoc = varMsgDoc & oTipDoc.get_Value(i) & " " & oNumDoc.get_Value(i) & Chr(13)
        End If
    Next i
    
    
    For i = 0 To oTipDocCo.ArraySize - 1
        If oTipDocCo.get_Value(i) <> "" Then
        varMsgDocCo = varMsgDocCo & oTipDocCo.get_Value(i) & " " & oNumDocCo.get_Value(i) & Chr(13)
        End If
    
    Next i
    
    MsgBox "Se realizo la transacción satisfactoriamente  - " & Chr(13) & varMsgDoc & _
            "Por convenio - " & varMsgDocCo, vbInformation + vbOKOnly, App.ProductName
            
    strUltDocEmi = ""
    
    For i = 0 To auxTipDoc.ArraySize - 1
        Set UltDocEmitido = objDocumento.UltDocEmitido(objUsuario.CodigoEmpresa, objUsuario.NombrePC)
        If Not UltDocEmitido.EOF Then strUltDocEmi = UltDocEmitido("COD_TIPO_DOCUMENTO").Value
        If auxTipDoc.get_Value(i) <> "" Then
            If auxTipDoc.get_Value(i) <> strUltDocEmi Then
                MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                        "en posicion de " & auxTipDoc.get_Value(i), vbInformation, App.ProductName
            End If
            objImpresion.ImprimirDocumento auxTipDoc.get_Value(i), auxNumDoc.get_Value(i)
            objDocumento.GrabaUltDocEmitido objUsuario.CodigoEmpresa, objUsuario.NombrePC, auxTipDoc.get_Value(i)
        End If
    Next i
            
    Unload Me
    ''frmPedido.psub_BeginArry
    mdiPrincipal.subNuevo
    frm_VTA_Busqueda.grdProductos.Limpiar
    frm_VTA_Busqueda.grdAlternativos.Limpiar
    frm_VTA_Busqueda.grdComplementarios.Limpiar
    frm_VTA_Busqueda.txtBuscar.selection
    frm_VTA_Modalidad.Show vbModal
      
      
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName



End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo CtrlErr
    
    
    objVenta.CargarProductoCobro dbcTipoDoc.BoundText, Replace(txtNumDocumento.Text, "-", "")
    grdProductosCobro.Array1 = objVenta.ProductoCobro
    grdProductosCobro.col = 6
    grdProductosCobro.SetFocus
    
    grdProductosCobro.Rebind
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName


End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo handle
    If Frame1.Visible = True Then
        Frame2.Visible = True
        Frame1.Visible = False
    Else
        Unload Me
        'objVenta.CancelarVenta
    End If
    
    Exit Sub
    
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
    
    
End Sub


Private Sub ctlDataCombo1_Click(Area As Integer)
    Frame1.Visible = True
End Sub

Private Sub cmdFf_Click()

    
        On Error GoTo CtrlErr
        If grdPersonalLocal.ApproxCount > 0 Then
            If Val(txtImporte.Text) > 0 Then
                objVenta.AgregaCobroResponsabilidad grdPersonalLocal.Columns(0), grdPersonalLocal.Columns(1) + " " + grdPersonalLocal.Columns(2) + ", " + grdPersonalLocal.Columns(3), txtImporte.Text, chkProrretear.Value
                grdPersonalCobrar.Rebind
                CalFooter
            Else
                MsgBox "El Importe Neto a Cobrar no es mayor a Cero", vbInformation, App.ProductName
            End If
        End If
        Exit Sub
        
CtrlErr:
        MsgBox Err.Description, vbCritical, App.ProductName


End Sub

Private Sub cmdRr_Click()
Dim intOriginal As Integer
                    
            On Error GoTo CtrlErr
                    
                If grdPersonalCobrar.ApproxCount = 0 Then
                    grdPersonalLocal.SetFocus
                    Exit Sub
                End If
                grdPersonalCobrar.Delete
                grdPersonalCobrar.Rebind
                CalFooter
                
                If grdPersonalCobrar.ApproxCount = 0 Then
                    grdPersonalLocal.SetFocus
                Else
                    If chkProrretear.Value = 1 Then
                        If Val(txtImporte.Text) > 0 Then objVenta.AgregaCobroResponsabilidad grdPersonalCobrar.Columns(0), grdPersonalCobrar.Columns(1), Val(txtImporte.Text), chkProrretear.Value
                    Else
                        If Val(txtImporte.Text) > 0 Then objVenta.AgregaCobroResponsabilidad grdPersonalCobrar.Columns(0), grdPersonalCobrar.Columns(1), grdPersonalCobrar.Columns(2).Value, chkProrretear.Value
                    End If
                End If
                'lblTotalCobrar.Caption = objVenta.TotalCobroResp
                intOriginal = chkProrretear.Value
                If chkProrretear.Value = 0 Then chkProrretear.Value = 1
                Call chkProrretear_Click
                chkProrretear.Value = intOriginal
                    
         Exit Sub
        
CtrlErr:
        MsgBox Err.Description, vbCritical, App.ProductName
                   

End Sub


Private Sub ctlTextBox1_KeyPress(KeyAscii As Integer)
'''

End Sub

Private Sub dbcMotivo_Click(Area As Integer)
Dim oraMotivo As oraDynaset
    On Error GoTo CtrlErr
    Select Case Area
        Case dbcAreaButton
        Case dbcAreaList
            Set oraMotivo = objVenta.ListaMotivoCobro(dbcMotivo.BoundText)
            If Not oraMotivo.EOF Then
'                If oraMotivo("FLG_DOCU_REFER").Value = "1" Then
'                    Frame1.Visible = True
'                    Frame2.Visible = False
'                Else
'                    Frame1.Visible = False
'                    Frame2.Visible = True
'                    txtImporte.Text = objVenta.Totales(0)
'                End If
                
                If oraMotivo("FLG_DOCU_REFER").Value = "1" Then frm_VTA_CobroRespoProd.Datos "Cobro por Cruce en contra"
                
                txtImporte.Text = Format(Val(txtImporteaCobrar.Text) - Val(txtImporteAFavor.Text), "0.00")
                
            End If
    End Select
    Exit Sub
    
CtrlErr:

    MsgBox Err.Description, vbOKOnly + vbQuestion, App.ProductName



End Sub

Private Sub Form_Activate()

    txtImporteaCobrar.Text = Format(frmPedido.LblTotal, "0.00")
    
    txtImporteaCobrar.Text = CalculaImporteACobrar
    
    txtImporte.Text = Format(Val(txtImporteaCobrar.Text) - Val(txtImporteAFavor.Text), "0.00")
    

End Sub

Private Sub Form_Load()
    setteaFormulario Me
    Frame1.Visible = False: Frame2.Visible = True
    txtCriterio.Bloqueado = True
    Set dbcMotivo.RowSource = objVenta.ListaMotivoCobro
    dbcMotivo.ListField = "DES_MOTIVO_COBRO"
    dbcMotivo.BoundColumn = "COD_MOTIVO_COBRO"
         
    SetteaGrd
    
    Opt0(0).Value = True
    Set grdPersonalLocal.DataSource = objUsuario.Lista("", objUsuario.CodigoLocal, "ACT")
    grdPersonalCobrar.Array1 = objVenta.CobroResponsabilidad
End Sub

Private Sub SetteaGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
Dim arrFoco As Variant

arrCampos = Array("COD_USUARIO", "APE_PAT_USUARIO", "APE_MAT_USUARIO", "DES_NOMBRE")
arrCaption = Array("Codigo", "Ape. Paterno", "Ape. Materno", "Nombre")
arrAncho = Array(800, 1000, 1000, 2000)
arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgLeft)
grdPersonalLocal.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

arrCampos = Array("", "", "")
arrCaption = Array("Código", "Nombre", "Importe")
arrAncho = Array(800, 3100, 900)
arrAlineacion = Array(dbgCenter, dbgLeft, dbgRight)
arrFoco = Array(False, False, True)
grdPersonalCobrar.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
grdPersonalCobrar.AllowUpdate = True
grdPersonalCobrar.Columns(0).AllowFocus = False
grdPersonalCobrar.Columns(1).AllowFocus = False
''grdPersonalCobrar.Columns(2).EditMask = ""
grdPersonalCobrar.Columns(2).NumberFormat = "Fixed"

grdPersonalCobrar.ColumnFooter = True
grdPersonalCobrar.Columns(1).FooterAlignment = dbgRight
grdPersonalCobrar.Columns(1).FooterText = "Total"
grdPersonalCobrar.Columns(2).FooterAlignment = dbgLeft




arrCampos = Array("", "", "", "", "", "", "", "")
arrCaption = Array("Código", "Descripción", "Fracc", "Cant", "Importe", "Cant.", "Unid.", "Importe a Favor")
arrAncho = Array(800, 2500, 200, 500, 700, 800, 500, 1200)
arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgCenter, dbgRight, dbgRight)
arrFoco = Array(False, False, False, False, False, True, True, True)
grdProductosCobro.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
'grdProductosCobro.Columns(2).Visible = False
'grdProductosCobro.Columns(3).Visible = False
'grdProductosCobro.Columns(4).Visible = False




End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyEscape
            cmdCancelar_Click
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub




Private Function CalculaImporteACobrar() As Double
Dim i As Integer
Dim dblImporteACobrar As Double

        dblImporteACobrar = 0

        For i = 0 To objVenta.Producto.UpperBound(1)
                If objVenta.Producto(i, 4) > 0 Then
                    dblImporteACobrar = dblImporteACobrar + objVenta.Producto(i, 4)
                End If
        Next i


        CalculaImporteACobrar = dblImporteACobrar

End Function

Private Sub grdPersonalCobrar_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Select Case ColIndex
        Case 2
            If Not IsNumeric(grdPersonalCobrar.Columns(ColIndex).Value) Then
                grdPersonalCobrar.Columns(ColIndex).Value = OldValue
                Cancel = 1
                
            End If
            CalFooter
    End Select

End Sub


Private Sub grdPersonalCobrar_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
            grdPersonalCobrar.Update
            grdPersonalCobrar.Rebind
    End If
    CalFooter
End Sub




Private Sub CalFooter()
    Dim k%
    Dim intCant As Integer
    
    intCant = 0: dblTotal = 0
    
    For k = 0 To objVenta.CobroResponsabilidad.UpperBound(1)
        If objVenta.CobroResponsabilidad(k, 2) > 0 Or objVenta.CobroResponsabilidad(k, 2) <> "" Then
            dblTotal = dblTotal + Val(objVenta.CobroResponsabilidad(k, 2))
        End If
    Next k
    grdPersonalCobrar.Columns(2).FooterText = Format(dblTotal, "#,###,##0.00")
    
    
End Sub

Private Sub Opt0_Click(Index As Integer)
    If Opt0(0).Value = True Then
        txtCriterio.Bloqueado = True
        Set grdPersonalLocal.DataSource = objUsuario.Lista("", objUsuario.CodigoLocal, "EMI")
        grdPersonalCobrar.Array1 = objVenta.CobroResponsabilidad
      Else
      txtCriterio.Bloqueado = False
        Set grdPersonalLocal.DataSource = objUsuario.Lista_No_Asign(objUsuario.CodigoLocal)
        grdPersonalCobrar.Array1 = objVenta.CobroResponsabilidad
    End If
    
End Sub

Private Sub txtCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Set grdPersonalLocal.DataSource = objUsuario.Lista_No_Asign(objUsuario.CodigoLocal, Trim(txtCriterio.Text))
    grdPersonalLocal.Refresh
End If

End Sub





