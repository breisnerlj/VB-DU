VERSION 5.00
Begin VB.Form frm_VTA_GuiaRemision 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      ForeColor       =   &H80000007&
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6615
      Begin vbp_Ventas.ctlDataCombo cboTipoDev 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboMotivoDev 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboOrigen 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   705
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5580
      Picture         =   "frm_VTA_GuiaRemision.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4260
      Picture         =   "frm_VTA_GuiaRemision.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ajuste"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   6615
      Begin vbp_Ventas.ctlDataCombo cboMotivo 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboTipo 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   760
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Emision de Guia"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   6615
      Begin vbp_Ventas.ctlDataCombo cboOrigen2 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblDestino 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   540
      End
   End
   Begin vbp_Ventas.ctlTextBox txtNota 
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   1508
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Width           =   1065
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
      Left            =   5910
      TabIndex        =   11
      Top             =   6660
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
      Left            =   4200
      TabIndex        =   10
      Top             =   6660
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_GuiaRemision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLocal As New clsLocal
Dim objAjuste As New clsAjuste
Dim objDocumento As New clsDocumento
Public FlgErr As Boolean

Private Sub cboMotivoDev_Change()
    Dim objLista As New clsGuia
    Dim objResult As oraDynaset
    
    On Error GoTo Control
    
    With cboOrigen
        Set objResult = objLista.ListaDestinos(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, cboTipoDev.BoundText, cboMotivoDev.BoundText)
        Set .RowSource = objResult
        .ListField = "DES"
        .BoundColumn = "COD"
        
        If objResult.RecordCount = 2 Then
            objResult.MoveFirst
            objResult.MoveNext
            .BoundText = objResult("COD").Value
        Else
            .BoundText = "*"
        End If
        
    End With
            
    Set objLista = Nothing
    Set objResult = Nothing
    

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cboOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Check1.Visible = True Then Check1.SetFocus
    End If
End Sub

Private Sub cboTipo_Change()
    cboMotivo.BoundText = ""
    Set cboMotivo.RowSource = objAjuste.ListaMotivoTipoAju(cboTipo.BoundText, objUsuario.CodigoLocal)
    cboMotivo.ListField = "DES_MOTIVO_AJUSTE"
    cboMotivo.BoundColumn = "COD_MOTIVO_AJUSTE"
End Sub

Public Sub cboTipoDev_Change()
    Dim objMotivoDev As New clsMotivoDev
    Dim objResult As oraDynaset
    
    On Error GoTo Control
    
    With cboMotivoDev
        
        
        
        Set objResult = objMotivoDev.Lista(cboTipoDev.BoundText, objUsuario.CodigoLocal, IIf(objVenta.CodDocRef = "PRO", "", "1"))
        Set .RowSource = objResult
        .ListField = "DES"
        .BoundColumn = "COD"
        
        If objResult.RecordCount = 2 Then
            objResult.MoveFirst
            objResult.MoveNext
            .BoundText = objResult("COD").Value
        Else
            .BoundText = "*"
        End If
        
    End With
            
    Set objMotivoDev = Nothing
    Set objResult = Nothing
    

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        cboMotivo.Enabled = False
        cboMotivo.BoundText = ""
        cboTipo.Enabled = False
        cboTipo.BoundText = ""
    Else
        cboMotivo.Enabled = True
        cboTipo.Enabled = True
    End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboMotivo.Enabled = False Then txtNota.SetFocus Else cboMotivo.SetFocus
    End If
End Sub

Public Sub cmdAceptar_Click()
    Dim strCadCodProducto As String
    Dim strCadCtdProducto As String
    Dim strCadCtdProductoFrac As String
    Dim strCadNumLote As String
    Dim strCadFchVenc As String
    
    Dim strCadImpProducto As String
    Dim strCadProdFormulario As String
    Dim strCadNumFormulario As String
    
    Dim GrabaGuia As String
    Dim arrValor As Variant
    Dim strNumDocumento As String
    Dim strCodFacSap As String
    Dim strImpTotal As String
    Dim i As Integer
    Dim j As Integer
    Dim indice As Integer
    Dim arrDocs As Variant
    Dim strNumTmp As String
    Dim FlgDevFraccion As String
On Error GoTo CtrlErr

'    If cboOrigen.BoundText = "CJE" Then
'        MsgBox "No se pueden generar Guías de Transferencia a Canje sin " & vbCrLf & _
'               "una orden de devolución, por favor utilize la opción Transferencia " & vbCrLf & _
'               "a Canje desde el Módulo Administrador.", _
'               vbCritical + vbOKOnly, "Error"
'        Exit Sub
'    End If
    
    If objUsuario.CodigoLocal = cboOrigen.BoundText And Check1.Value = 0 Then
        MsgBox "El local destino no puede ser el local de emisión", vbCritical + vbOKOnly, App.ProductName
        Exit Sub
    End If

    If Check1.Value = 1 And cboTipo.BoundText = "" Then
        MsgBox "Debe escoger el tipo.", vbCritical + vbOKOnly, App.ProductName
        Exit Sub
    End If
    
    If Check1.Value = 1 And cboMotivo.BoundText = "" Then
            MsgBox "Debe escoger el motivo.", vbCritical + vbOKOnly, App.ProductName
            Exit Sub
    End If
    FlgDevFraccion = "0"
    FlgDevFraccion = objAjuste.DevPermiteFraccion(cboTipoDev.BoundText, cboMotivoDev.BoundText)
    i = 0
    j = 0
    
    '===================================================================================================================================
    '===================================================================================================================================
    '===================================================================================================================================
    'MODIFICADO POR 10/10/2012 MLEVANO
    'While i <= objVenta.Producto.UpperBound(1)
    
    If objVenta.Producto.Count(1) = 0 Then
        MsgBox "No ha Seleccionado Productos", vbCritical, App.ProductName
         Exit Sub
    End If
    
     
     If objVenta.ptmModalidad = Guias_Remision Then
        While j <= objVenta.Producto.UpperBound(1)
           If objVenta.ProductoLote.Count(1) > 0 Then
               indice = objVenta.ProductoLote.Find(0, 0, objVenta.Producto(j, 0), XORDER_ASCEND)
               If indice < 0 Then
                  MsgBox "El producto " & objVenta.Producto(j, 0) & " esta sin datos lote o fecha de vencimiento", vbCritical, App.ProductName
                  Exit Sub
               End If
               
           Else
                 MsgBox "Los Productos estan sin datos de lote o fecha de vencimiento", vbCritical, App.ProductName
                 Exit Sub
           End If
           j = j + 1
        Wend
     End If
     
   If objVenta.ptmModalidad = Guias_Remision Then
        While i <= objVenta.ProductoLote.UpperBound(1)
            Dim objproducto As New clsProducto
            If objVenta.ProductoLote(i, 0) <> "" Then
                'If objProducto.EsPsicotropico(objVenta.Producto(i, 0)) = "1" And cboOrigen.BoundText <> "PRV" Then
                    'MsgBox "El producto " & objVenta.Producto(i, 1) & " es un Psicotropico y esta prohibido su trasnferencia entre locales", vbCritical, App.ProductName
                If objproducto.EsPsicotropico(objVenta.ProductoLote(i, 0)) = "1" And cboOrigen.BoundText <> "PRV" Then
                    MsgBox "El producto " & objVenta.ProductoLote(i, 1) & " es un Psicotropico y esta prohibido su trasnferencia entre locales", vbCritical, App.ProductName
                    Exit Sub
                End If
                'If FlgDevFraccion = "0" And objVenta.Producto(i, 6) = "1" Then
                    'MsgBox "El motivo " & cboMotivoDev.Text & " no permite devolver fracciones, verifique el producto" & objVenta.Producto(i, 1) & " ", vbCritical, App.ProductName
                 If FlgDevFraccion = "0" And objVenta.ProductoLote(i, 6) = "1" Then
                    MsgBox "El motivo " & cboMotivoDev.Text & " no permite devolver fracciones, verifique el producto" & objVenta.ProductoLote(i, 1) & " ", vbCritical, App.ProductName
                    Exit Sub
                End If
                'strCadCodProducto = strCadCodProducto & objVenta.Producto(i, 0) & "|"
                'If objVenta.Producto(i, 6) = "0" Then
                    'strCadCtdProducto = strCadCtdProducto & objVenta.Producto(i, 3) & "|"
                strCadCodProducto = strCadCodProducto & objVenta.ProductoLote(i, 0) & "|"
                If objVenta.ProductoLote(i, 6) = "0" Then
                    strCadCtdProducto = strCadCtdProducto & objVenta.ProductoLote(i, 3) & "|"
                    strCadCtdProductoFrac = strCadCtdProductoFrac & "0|"
                Else
                    strCadCtdProducto = strCadCtdProducto & "0|"
                    'strCadCtdProductoFrac = strCadCtdProductoFrac & objVenta.Producto(i, 3) & "|"
                    strCadCtdProductoFrac = strCadCtdProductoFrac & objVenta.ProductoLote(i, 3) & "|"
                End If
                
                'strCadImpProducto = strCadImpProducto & objVenta.Producto(i, 4) & "|"
                'strCadNumLote = strCadNumLote & UCase(objVenta.Producto(i, 22)) & "|"
                'strCadFchVenc = strCadFchVenc & Replace(objVenta.Producto(i, 23), "__/__/____", "") & "|"
                strCadImpProducto = strCadImpProducto & objVenta.ProductoLote(i, 4) & "|"
                strCadNumLote = strCadNumLote & UCase(objVenta.ProductoLote(i, 22)) & "|"
                strCadFchVenc = strCadFchVenc & Replace(objVenta.ProductoLote(i, 23), "__/__/____", "") & "|"
            End If
            i = i + 1
        Wend
    Else
        While i <= objVenta.Producto.UpperBound(1)
            'Dim objproducto As New clsProducto
            If objproducto.EsPsicotropico(objVenta.Producto(i, 0)) = "1" And cboOrigen.BoundText <> "PRV" Then
                MsgBox "El producto " & objVenta.Producto(i, 1) & " es un Psicotropico y esta prohibido su trasnferencia entre locales", vbCritical, App.ProductName
            'If objproducto.EsPsicotropico(objVenta.ProductoLote(i, 0)) = "1" And cboOrigen.BoundText <> "PRV" Then
                'MsgBox "El producto " & objVenta.ProductoLote(i, 1) & " es un Psicotropico y esta prohibido su trasnferencia entre locales", vbCritical, App.ProductName
                Exit Sub
            End If
            If FlgDevFraccion = "0" And objVenta.Producto(i, 6) = "1" Then
                MsgBox "El motivo " & cboMotivoDev.Text & " no permite devolver fracciones, verifique el producto" & objVenta.Producto(i, 1) & " ", vbCritical, App.ProductName
             'If FlgDevFraccion = "0" And objVenta.ProductoLote(i, 6) = "1" Then
                'MsgBox "El motivo " & cboMotivoDev.Text & " no permite devolver fracciones, verifique el producto" & objVenta.ProductoLote(i, 1) & " ", vbCritical, App.ProductName
                Exit Sub
            End If
            strCadCodProducto = strCadCodProducto & objVenta.Producto(i, 0) & "|"
            If objVenta.Producto(i, 6) = "0" Then
                strCadCtdProducto = strCadCtdProducto & objVenta.Producto(i, 3) & "|"
            'strCadCodProducto = strCadCodProducto & objVenta.ProductoLote(i, 0) & "|"
            'If objVenta.ProductoLote(i, 6) = "0" Then
                'strCadCtdProducto = strCadCtdProducto & objVenta.ProductoLote(i, 3) & "|"
                strCadCtdProductoFrac = strCadCtdProductoFrac & "0|"
            Else
                strCadCtdProducto = strCadCtdProducto & "0|"
                strCadCtdProductoFrac = strCadCtdProductoFrac & objVenta.Producto(i, 3) & "|"
                'strCadCtdProductoFrac = strCadCtdProductoFrac & objVenta.ProductoLote(i, 3) & "|"
            End If
            
            strCadImpProducto = strCadImpProducto & objVenta.Producto(i, 4) & "|"
            strCadNumLote = strCadNumLote & UCase(objVenta.Producto(i, 22)) & "|"
            strCadFchVenc = strCadFchVenc & Replace(objVenta.Producto(i, 23), "__/__/____", "") & "|"
            'strCadImpProducto = strCadImpProducto & objVenta.ProductoLote(i, 4) & "|"
            'strCadNumLote = strCadNumLote & UCase(objVenta.ProductoLote(i, 22)) & "|"
            'strCadFchVenc = strCadFchVenc & Replace(objVenta.ProductoLote(i, 23), "__/__/____", "") & "|"
            
            i = i + 1
        Wend
        
    End If
    
    i = 0
    While i <= objVenta.EspeciesValoradas.UpperBound(1)
        strCadProdFormulario = strCadProdFormulario & objVenta.EspeciesValoradas(i, 0) & "|"
        strCadNumFormulario = strCadNumFormulario & objVenta.EspeciesValoradas(i, 1) & "|"
        i = i + 1
    Wend
    
    Dim gvarError As String
    
    'GrabaGuia = objDocumento.GrabaGuia(objUsuario.CodigoEmpresa, _
                           objUsuario.CodigoLocal, _
                           objUsuario.TipoDocGuia, _
                           objUsuario.NombrePC, _
                           "", _
                           "", _
                           "1", _
                           cboOrigen.BoundText, _
                           objUsuario.Codigo, _
                           "001", _
                           "", _
                           "TRL", _
                           Trim(strCodProducto), _
                           Trim(strCtdProductoFrac), _
                           Trim(strCtdProducto), _
                           Trim(strImpProducto), _
                           strNumDocumento, _
                           strImpTotal, _
                           "S", _
                           arrValor)
    strNumDocumento = ""
        
    If Check1.Value = 0 Then
            
        If objVenta.ptmModalidad <> Guias_Remision Then
            If objVenta.ValidaVentaLotes = False Then
                    Exit Sub
            End If
        End If
    
        'If objVenta.ValidaVentaLotes = False Then
            'Exit Sub
         ' Else
            i = 0
            strCadNumLote = "": strCadFchVenc = ""
            '===================================================================================================================================
            '===================================================================================================================================
            '===================================================================================================================================
            'MODIFICADO POR 10/10/2012 MLEVANO
         If objVenta.ptmModalidad = Guias_Remision Then
            While i <= objVenta.ProductoLote.UpperBound(1)
                strCadNumLote = strCadNumLote & UCase(objVenta.ProductoLote(i, 22)) & "|"
                strCadFchVenc = strCadFchVenc & Replace(objVenta.ProductoLote(i, 23), "__/__/____", "") & "|"
'            While i <= objVenta.Producto.UpperBound(1)
'                strCadNumLote = strCadNumLote & UCase(objVenta.Producto(i, 22)) & "|"
'                strCadFchVenc = strCadFchVenc & Replace(objVenta.Producto(i, 23), "__/__/____", "") & "|"
                i = i + 1
            Wend
         End If
        'End If
    
        GrabaGuia = objDocumento.GeneraGuiaLocal(strNumDocumento, objUsuario.CodigoEmpresa, _
                                    objUsuario.CodigoLocal, _
                                    cboOrigen.BoundText, _
                                    objUsuario.NombrePC, _
                                    objUsuario.TipoDocGuia, _
                                    objUsuario.MotivoGeneraGuiaLocal, _
                                    objUsuario.Codigo, _
                                    txtNota.Text, _
                                    strCadCodProducto, _
                                    strCadCtdProducto, _
                                    strCadCtdProductoFrac, _
                                    strCadProdFormulario, _
                                    strCadNumFormulario, _
                                    strCadNumLote, strCadFchVenc, _
                                    objVenta.CodDocRef, objVenta.NumDocRef, _
                                    cboTipoDev.BoundText, cboMotivoDev.BoundText, strCodFacSap)
    Else
        GrabaGuia = objDocumento.GeneraGuiaAjuste(strNumDocumento, objUsuario.CodigoEmpresa, _
                                    objUsuario.CodigoLocal, _
                                    cboTipo.BoundText, _
                                    objUsuario.Codigo, _
                                    txtNota.Text, _
                                    strCadCodProducto, _
                                    strCadCtdProducto, _
                                    strCadCtdProductoFrac, _
                                    strCadImpProducto, _
                                    cboMotivo.BoundText, _
                                    gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FLG_AJUSTES_HABILITADOS"))
    
    End If
                           
    'Dim Valor As String
    'Valor = arrValor(0)
    If GrabaGuia = "" Then
                    
        If InStr(strNumDocumento, "|") = 0 Then strNumDocumento = strNumDocumento & "|"
       
        arrDocs = Split(strNumDocumento, "|")

        FlgErr = False
        MsgBox "Se grabó las guias Nº " & vbCrLf & Join(arrDocs, vbCrLf), vbInformation, App.ProductName
        MsgBox "Sirvase verificar el formato de guia en la impresora", vbInformation, App.ProductName
        If Check1.Value = 0 Then
                    
            Dim objImpre As New clsImpresion
            If objImpre.PuedeImprimir(objUsuario.CodigoEmpresa, objUsuario.NombrePC, "GRL") = True Then
                
                For i = LBound(arrDocs) To UBound(arrDocs)
                    If Trim(arrDocs(i)) = vbNullString Then Exit For
                    strNumTmp = arrDocs(i)
                    If objImpre.ListaPendiente(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.NombrePC, strNumTmp) > 0 Then frmColaImpresion.Show vbModal: Exit Sub
                    objDocumento.ImprimirGuiaTransferencia strNumTmp
                Next i
            End If
            Set objImpre = Nothing
        Else
            For i = LBound(arrDocs) To UBound(arrDocs)
                If Trim(arrDocs(i)) = vbNullString Then Exit For
                strNumTmp = arrDocs(i)
                objDocumento.Imprime_Ajuste_Cje strNumTmp
            Next i
        End If
        Unload Me
        Call Salida
    Else
        FlgErr = True
        MsgBox GrabaGuia, vbCritical, App.ProductName
    End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, Err.Number
    'Unload Me
    'Call Salida
End Sub

Private Sub Form_Load()
   On Error GoTo CtrlErr

    setteaFormulario Me
    Set cboOrigen.RowSource = objLocal.ListaXUbigeo(objUsuario.CodigoEmpresa, "", "A", "002,003", "")
    cboOrigen.BoundColumn = "COD_LOCAL"
    cboOrigen.ListField = "local_dex"
    
    
    CargaListaTipos
    
    
    If objVenta.ParametroValor("ACTAJUSTE") = 1 Then
        Check1.Visible = False
        Frame2.Visible = False
    Else
        Set cboTipo.RowSource = objAjuste.ListaTipoAju
            cboTipo.ListField = "DES_TIP_AJUSTE"
            cboTipo.BoundColumn = "COD_TIP_AJUSTE"
            
        Check1.Visible = True
        Frame2.Visible = True
    End If

   
   Exit Sub

CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            'If Shift = 1 Then cmdAceptar_Click
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    'objVenta.CancelarVenta
End Sub

Sub subNuevo()
    'ctlCliente1.Limpiar
    frmPedido.lblTotal.Caption = "0.00": frmPedido.lblTotalPagar.Caption = "0.00":
    frmPedido.lblRedondeo.Caption = "0.00": frmPedido.lblPagado.Caption = "0.00":
    frmPedido.lblVuelto.Caption = "0.00"
    '-- Limpia Variables Publicas --'
    frmPedido.pstrCant = "": frmPedido.pstrPrecio = "": frmPedido.pstrProd = "": frmPedido.pstrSubTot = ""
    frmPedido.pstrCtdFracc = "": frmPedido.pstrCtdFracc = ""
    frmPedido.pstrPctUnit = "": frmPedido.pstrPrcUniKairo = ""
    frmPedido.pstrImpuesto = "": frmPedido.pstrComision = ""
    frmPedido.pstrPromocion = ""
    frm_VTA_FacServPrestados.pstrIDKeito = ""
    frm_VTA_RecetarioM.pstrFlgRM = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objLocal = Nothing
    Set objDocumento = Nothing
End Sub

Private Sub Salida()
    ''frmPedido.psub_BeginArry
    mdiPrincipal.subNuevo
    frm_VTA_Busqueda.grdProductos.Limpiar
    frm_VTA_Busqueda.grdAlternativos.Limpiar
    frm_VTA_Busqueda.grdComplementarios.Limpiar
    frm_VTA_Busqueda.txtBuscar.selection
    frm_VTA_Modalidad.Show vbModal
End Sub

Private Sub CargaListaTipos()
Dim objTipoDev As New clsTipoDev
Dim objResult As oraDynaset

    With cboTipoDev
        Set objResult = objTipoDev.Lista(objUsuario.CodigoLocal, "1")
        Set .RowSource = objResult
        .ListField = "DES"
        .BoundColumn = "COD"

        If objResult.RecordCount = 2 Then
            objResult.MoveFirst
            objResult.MoveNext
            .BoundText = objResult("COD").Value
        Else
            .BoundText = "*"
        End If

        
    End With
    
    Set objResult = Nothing
    Set objTipoDev = Nothing
        
End Sub

