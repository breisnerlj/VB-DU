VERSION 5.00
Begin VB.Form frm_ADM_ProdDevGuia 
   Caption         =   "Datos de Transporte"
   ClientHeight    =   5895
   ClientLeft      =   4950
   ClientTop       =   4935
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6615
   Begin VB.CommandButton cmdEsc 
      Caption         =   "[Esc] Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdF11 
      Caption         =   "[F11] Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   6135
      Begin vbp_Ventas.ctlTextBox txtRucTransp 
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         Tipo            =   3
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
      Begin vbp_Ventas.ctlTextBox txtTransp 
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
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
      Begin vbp_Ventas.ctlTextBox txtDireccionTransp 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
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
      Begin vbp_Ventas.ctlTextBox txtPlacaTransp 
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Placa :   "
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Direccion :   "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Transportista :   "
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "RUC Transp. :   "
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   6135
      Begin vbp_Ventas.ctlTextBox txtProveedor 
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
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
      Begin vbp_Ventas.ctlTextBox txtRucProveedor 
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
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
      Begin vbp_Ventas.ctlTextBox txtDireccion 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Sres. :   "
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "RUC :   "
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Direccion :   "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   6135
      Begin vbp_Ventas.ctlDataCombo cboTipoDev 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboMotivoDev 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboOrigen 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Destino :   "
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo :   "
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo :   "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_ADM_ProdDevGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLocal As New clsLocal
Dim objEntrega As New clsEntrega
Public strTipo As String
Public strCadCodProducto, strCadCtdProducto, strCadNumLote, strCadFchVenc, strCadCtdFProducto, strCadCodFacSap As String
Public strIdEntrega As String

Private Sub ctlTextBox3_KeyPress(KeyAscii As Integer)

End Sub

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
            .BoundText = objVenta.ParametroValor("PROVDEVAL") '"PRV"
        End If
        
        Me.cboOrigen.Enabled = False
        
    End With
            
    Set objLista = Nothing
    Set objResult = Nothing
    

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cboTipoDev_Change()
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
            If strTipo = "DET" Then
            .BoundText = "DP7"
            ElseIf strTipo = "VCMTO" Then
            .BoundText = "DP2"
            Else
            .BoundText = "*"
            End If
            Me.cboMotivoDev.Enabled = False
        End If
        
    End With
            
    Set objMotivoDev = Nothing
    Set objResult = Nothing
    

    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
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
            .BoundText = "DPR"
        End If
        Me.cboTipoDev.Enabled = False
        
    End With
    
    Set objResult = Nothing
    Set objTipoDev = Nothing
        
End Sub

Private Sub cmdEsc_Click()
Unload Me
End Sub

Private Sub cmdF11_Click()
On Error GoTo Control
If validaForm Then
    
    Dim objDocumento As New clsDocumento
    Dim objEntrega As New clsEntrega
    Dim arr As Variant
    Dim GrabaGuia As String
    Dim strNumDocumento As String
    strNumDocumento = ""
    GrabaGuia = objDocumento.GeneraGuiaLocal(strNumDocumento, objUsuario.CodigoEmpresa, _
                                    objUsuario.CodigoLocal, _
                                    cboOrigen.BoundText, _
                                    objUsuario.NombrePC, _
                                    objUsuario.TipoDocGuia, _
                                    objUsuario.MotivoGeneraGuiaLocal, _
                                    objUsuario.Codigo, _
                                    "", _
                                    strCadCodProducto, _
                                    strCadCtdProducto, _
                                    strCadCtdFProducto, _
                                    "", _
                                    "", _
                                    strCadNumLote, strCadFchVenc, _
                                    "", "", _
                                    cboTipoDev.BoundText, cboMotivoDev.BoundText, strCadCodFacSap)
   
   objEntrega.UpdDevueltosRecep strIdEntrega, strCadCodProducto, strCadCtdProducto
   
   Dim cadenaGuias As String
   cadenaGuias = strNumDocumento
   
   'MsgBox "se ha generado la guia Nº " & strNumDocumento, vbCritical + vbInformation, "Guia Generada"
   arr = Split(cadenaGuias, "|")
   If UBound(arr) = 1 Then
    MsgBox "Se ha generado la(s) Guía(s) Nº " + Replace(cadenaGuias, "|", " ") + "." + Chr(13) + "Verifique su Impresora", vbInformation, "Guías de Devolución"
   Else
    MsgBox "Se ha generado la(s) Guía(s) Nº " + Replace(cadenaGuias, "|", " - ") + "." + Chr(13) + "Verifique su Impresora", vbInformation, "Guías de Devolución"
   End If
   Dim j As Integer
   For j = 0 To UBound(arr)
        objEntrega.UpdTransportistaGuia arr(j), Me.txtRucTransp.Text, Me.txtDireccionTransp.Text, Me.txtPlacaTransp.Text, objUsuario.CodigoLocal
        objDocumento.ImprimirGuiaTransferencia arr(j)
   Next
   Dim Tipo As String
   Tipo = ""
   If strTipo = "DET" Then
    Tipo = "D"
   ElseIf strTipo = "VCMTO" Then
    Tipo = "V"
   End If
   objEntrega.UpdFlgGeneraGuia strIdEntrega, Tipo
   'Unload Me
   frm_ADM_Sobrantes.recibe
   
End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    cmdEsc_Click
End If
If KeyCode = vbKeyF11 Then
    cmdF11_Click
End If
End Sub

Private Sub Form_Load()
CargaListaTipos
Dim odyn As oraDynaset
Set odyn = objEntrega.ListaDatosProv(strIdEntrega)
odyn.MoveFirst

Me.txtRucProveedor.Text = "" & odyn("RUC_PROVEEDOR")
Me.txtDireccion.Text = "" & odyn("DIR_PROVEEDOR")
Me.txtProveedor.Text = "" & odyn("NOM_PROVEEDOR")

Me.txtRucProveedor.Enabled = False
Me.txtDireccion.Enabled = False
Me.txtProveedor.Enabled = False
End Sub

Function validaForm()
    validaForm = True
    If Me.txtRucTransp.Text = "" Then
        MsgBox "Ingrese RUC de Transportista", vbCritical + vbInformation, "Aviso"
        validaForm = False
        Exit Function
    End If
    If Me.txtTransp.Text = "" Then
        MsgBox "Ingrese Transportista", vbCritical + vbInformation, "Aviso"
        validaForm = False
        Exit Function
    End If
    If Me.txtDireccionTransp.Text = "" Then
        MsgBox "Ingrese Direccion de Transportista", vbCritical + vbInformation, "Aviso"
        validaForm = False
        Exit Function
    End If
    If Me.txtPlacaTransp.Text = "" Then
        MsgBox "Ingrese Placa de Transportista", vbCritical + vbInformation, "Aviso"
        validaForm = False
        Exit Function
    End If
End Function

Private Sub txtRucTransp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim odyn As oraDynaset
    Set odyn = objEntrega.BuscaTransportista(Me.txtRucTransp.Text)
    If odyn.RecordCount > 0 Then
        Me.txtRucTransp.Text = "" & odyn(1).Value
        Me.txtTransp.Text = "" & odyn(2).Value
    Else
        MsgBox "No se encontró la empresa de Transportes", vbCritical + vbInformation, "Validacion"
    End If
End If
End Sub
