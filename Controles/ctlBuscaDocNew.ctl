VERSION 5.00
Begin VB.UserControl ctlBuscaDocNew 
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   1185
   ScaleWidth      =   2940
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      Begin vbp_Ventas.ctlDataCombo dbcTipo 
         Height          =   315
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
      End
      Begin vbp_Ventas.ctlTextBox TxtBusca 
         Height          =   330
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
      Begin VB.Label Label2 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nª"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   225
      End
   End
End
Attribute VB_Name = "ctlBuscaDocNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento
Event RetornaDoc(odynCab As oraDynaset, odynDet As oraDynaset) 'Declarando un Evento para la busqueda'


Private Sub TxtBusca_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
'On Error GoTo CntrlErr
 Dim odynC As oraDynaset
 Dim odynD As oraDynaset
 '-----------------------'
 Dim odynCdocumentos As oraDynaset
 Dim odynDdocumentos As oraDynaset
 
 Dim strEsConvenio As String
 Dim strCodCnvEnBTLPROD As String
 
    If KeyAscii = 13 Then
        TxtBusca.Text = Trim(TxtBusca.Text)
        'Cabecera'
        Set odynC = objDocumento.ListaCabecera(objUsuario.CodigoEmpresa, Trim(TxtBusca.Text), _
                                               dbcTipo.BoundText, objUsuario.CodigoLocal, "0")
        
        If odynC.RecordCount <= 0 Then
            'Paso 1 : Saber si el documento es convenio y de eso sacar el codigo de convenio antiguo'
             strEsConvenio = objDocumento.DocumentoEsConvenio(objUsuario.CodigoLocal, _
                                                              Trim(TxtBusca.Text), _
                                                              "1")
                                                                          
            If strEsConvenio <> "" Then
                'Paso 2: Del codigo antiguo conseguir su codigo en el nuevo sistema
                strCodCnvEnBTLPROD = objDocumento.FindConvenioEnBtlprod(objUsuario.CodigoEmpresa, _
                                                                        strEsConvenio)
            End If
            
            
            'Detalle de c_documentos de BTLCADENA'
            Set odynCdocumentos = objDocumento.Lista_CDocumentos(Trim(TxtBusca.Text), dbcTipo.BoundText, objUsuario.CodigoLocal)
        End If
        
        'Detalle'
        Set odynD = objDocumento.ListaDetalle(objUsuario.CodigoEmpresa, Trim(TxtBusca.Text), _
                                               dbcTipo.BoundText, objUsuario.CodigoLocal)
                                                       
        If odynD.RecordCount <= 0 Then
            'Detalle de d_documentos de BTLCADENA'
            Set odynDdocumentos = objDocumento.Lista_DDocumentos(Trim(TxtBusca.Text), dbcTipo.BoundText)
        End If
                                        
        If odynC.RecordCount > 0 Then
            RaiseEvent RetornaDoc(odynC, odynD) 'Ejecutando el evento'
          Else
            RaiseEvent RetornaDoc(odynCdocumentos, odynDdocumentos) 'Ejecutando el evento'
        End If
    End If
    Exit Sub
Handle:
     MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Public Sub CargarTipo()
    Set dbcTipo.RowSource = objDocumento.ListaTipoDocumento("", "1")
    dbcTipo.ListField = "DES_TIPODOC"
    dbcTipo.BoundColumn = "COD_TIPODOC"
End Sub
