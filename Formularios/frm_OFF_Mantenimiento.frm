VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_OFF_Mantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento - Contingencia.ini"
   ClientHeight    =   7815
   ClientLeft      =   465
   ClientTop       =   540
   ClientWidth     =   8685
   Icon            =   "frm_OFF_Mantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   2310
      Picture         =   "frm_OFF_Mantenimiento.frx":1C9A2
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   615
      Left            =   1215
      Picture         =   "frm_OFF_Mantenimiento.frx":1CF2C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAñadir 
      Caption         =   "&Añadir"
      Height          =   615
      Left            =   120
      Picture         =   "frm_OFF_Mantenimiento.frx":1D4B6
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6840
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "frm_OFF_Mantenimiento.frx":1DA40
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "G&losa"
      TabPicture(1)   =   "frm_OFF_Mantenimiento.frx":1DA5C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Moneda"
      TabPicture(2)   =   "frm_OFF_Mantenimiento.frx":1DA78
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Impresión"
      TabPicture(3)   =   "frm_OFF_Mantenimiento.frx":1DA94
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Documento"
      TabPicture(4)   =   "frm_OFF_Mantenimiento.frx":1DAB0
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Forma de Pago"
      TabPicture(5)   =   "frm_OFF_Mantenimiento.frx":1DACC
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "[F7] - Forma de Pago"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   32
         Top             =   780
         Width           =   8295
         Begin vbp_Ventas.ctlGrillaArray grdFormaPago 
            Height          =   2415
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   4260
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[F6] - Documento"
         Height          =   2775
         Left            =   120
         TabIndex        =   30
         Top             =   780
         Width           =   8295
         Begin vbp_Ventas.ctlGrillaArray grdDocumento 
            Height          =   2415
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   4260
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "[F2] - Local"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   24
         Top             =   3000
         Width           =   8175
         Begin vbp_Ventas.ctlTextBox txtSecuencia 
            Height          =   315
            Left            =   1080
            TabIndex        =   7
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
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
         Begin vbp_Ventas.ctlTextBox txtDireccionLocal 
            Height          =   315
            Left            =   1080
            TabIndex        =   5
            Top             =   720
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   556
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
         Begin vbp_Ventas.ctlTextBox txtLocal 
            Height          =   315
            Left            =   1080
            TabIndex        =   4
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            MaxLength       =   3
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
         Begin vbp_Ventas.ctlTextBox txtSerieTicket 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            MaxLength       =   12
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
            AutoSize        =   -1  'True
            Caption         =   "Secuencia:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1500
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Serie Tick.:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   780
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "[F5] - Lista Formatos de Impresión"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   23
         Top             =   780
         Width           =   8295
         Begin vbp_Ventas.ctlGrillaArray grdImpresion 
            Height          =   2415
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   4260
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "[F4] - Lista Monedas"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   22
         Top             =   780
         Width           =   8295
         Begin vbp_Ventas.ctlGrillaArray grdMoneda 
            Height          =   2415
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   4260
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "[F3] - Línea de Impresión"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   21
         Top             =   780
         Width           =   8295
         Begin vbp_Ventas.ctlTextBox txtGlosa 
            Height          =   315
            Left            =   840
            TabIndex        =   8
            Top             =   480
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Texto:"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   540
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "[F1] - Compañia"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   16
         Top             =   780
         Width           =   8175
         Begin vbp_Ventas.ctlTextBox txtDireccionCia 
            Height          =   315
            Left            =   1080
            TabIndex        =   3
            Top             =   1320
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   556
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
         Begin vbp_Ventas.ctlTextBox txtRUC 
            Height          =   315
            Left            =   1080
            TabIndex        =   2
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            MaxLength       =   11
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
         Begin vbp_Ventas.ctlTextBox txtEmpresa 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
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
         Begin vbp_Ventas.ctlTextBox txtCia 
            Height          =   315
            Left            =   1080
            TabIndex        =   0
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            MaxLength       =   2
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
            AutoSize        =   -1  'True
            Caption         =   "Cia:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "RUC:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   1020
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   1380
            Width           =   720
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   7365
      Picture         =   "frm_OFF_Mantenimiento.frx":1DAE8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   6060
      Picture         =   "frm_OFF_Mantenimiento.frx":1E072
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
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
      Left            =   7710
      TabIndex        =   15
      Top             =   7500
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
      Left            =   6000
      TabIndex        =   14
      Top             =   7500
      Width           =   1215
   End
End
Attribute VB_Name = "frm_OFF_Mantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCia As String
Dim strEmpresa As String
Dim strRuc As String
Dim strDireccion As String
Dim strLocal As String
Dim strDireccionLocal As String
Dim strSerieTicket As String
Dim strSecuencia As String
Dim strGlosa As String
Dim strCodMoneda As String
Dim strDesMoneda As String
Dim strSmbMoneda As String
'-------------------------------
Dim strCodFormaPago As String
Dim strDesFormaPago As String
'-------------------------------
Dim strCodDocumento As String
Dim strDesDocumento As String
Dim strNumDocumento As String
Dim strNumLinDocumento As String
Dim strAnchoDocumento As String
'-------------------------------

Dim strCodFormato As String
Dim strDesFormato As String
Dim strCtdAncho As String
Dim strCtdAlto As String

Dim objIni As New cls_ArchivoIni
Dim objDocumento As New cls_OFF_Documento
Dim objFormaPago As New cls_OFF_FormaPago



Private Sub cmdAceptar_Click()

On Error GoTo CtrlErr

    If txtCia.Text = "" Then
        SSTab1.Tab = 0
        MsgBox "Campo Cia no puede ser vacío", vbCritical, App.ProductName
        txtCia.SetFocus
        Exit Sub
    Else
        strCia = Trim(txtCia.Text)
    End If

    If txtEmpresa.Text = "" Then
        SSTab1.Tab = 0
        MsgBox "Campo Empresa no puede ser vacío", vbCritical, App.ProductName
        txtEmpresa.SetFocus
        Exit Sub
    Else
        strEmpresa = Trim(txtEmpresa.Text)
    End If

    If txtRUC.Text = "" Or Len(txtRUC.Text) < 11 Then
        SSTab1.Tab = 0
        MsgBox "Campo RUC no puede ser vacío o la longitud es menor de 11", vbCritical, App.ProductName
        txtRUC.SetFocus
        Exit Sub
    Else
        strRuc = Trim(txtRUC.Text)
    End If


    If txtDireccionCia.Text = "" Then
        SSTab1.Tab = 0
        MsgBox "Campo Direccion no puede ser vacío", vbCritical, App.ProductName
        txtDireccionCia.SetFocus
        Exit Sub
    Else
        strDireccion = Trim(txtDireccionCia.Text)
    End If


    If txtLocal.Text = "" Then
        SSTab1.Tab = 0
        MsgBox "Campo Local no puede ser vacio", vbCritical, App.ProductName
        txtLocal.SetFocus
        Exit Sub
    Else
        strLocal = Trim(txtLocal.Text)
    End If


    If txtDireccionLocal.Text = "" Then
        SSTab1.Tab = 0
        MsgBox "Campo Dirección del Local no puede ser vacio", vbCritical, App.ProductName
        txtDireccionLocal.SetFocus
        Exit Sub
    Else
        strDireccionLocal = Trim(txtDireccionLocal.Text)
    End If


    If txtSerieTicket.Text = "" Then
        SSTab1.Tab = 0
        MsgBox "Campo Serie de Etiquetera no puede ser vacio", vbCritical, App.ProductName
        txtSerieTicket.SetFocus
        Exit Sub
    Else
        strSerieTicket = Trim(txtSerieTicket.Text)
    End If

    If txtSecuencia.Text = "" Then
        SSTab1.Tab = 0
        MsgBox "Campo Secuencia no puede ser vacio", vbCritical, App.ProductName
        txtSecuencia.SetFocus
        Exit Sub
    Else
        strSecuencia = Trim(txtSecuencia.Text)
    End If
    
    
    

    If txtGlosa.Text = "" Then
        SSTab1.Tab = 1
        MsgBox "Campo Glosa no puede ser vacio", vbCritical, App.ProductName
        txtGlosa.SetFocus
        Exit Sub
    Else
        strGlosa = Trim(txtGlosa.Text)
    End If


    If grdMoneda.ApproxCount = 0 Then
        SSTab1.Tab = 2
        MsgBox "Detalle de Moneda no puede ser vacio", vbCritical, App.ProductName
    End If


    If grdImpresion.ApproxCount = 0 Then
        SSTab1.Tab = 3
        MsgBox "Detalle de Impresion no puede ser vacio", vbCritical, App.ProductName
    End If




    If grdDocumento.ApproxCount = 0 Then
        SSTab1.Tab = 4
        MsgBox "Detalle de Documento no puede ser vacio", vbCritical, App.ProductName
    End If


    If grdFormaPago.ApproxCount = 0 Then
        SSTab1.Tab = 5
        MsgBox "Detalle de Forma de Pago no puede ser vacio", vbCritical, App.ProductName
    End If


    Guardar
    
    MsgBox "Se completo el proceso satisfactóriamente", vbInformation, App.ProductName
    
    Unload Me

Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName


    
End Sub

Private Sub cmdAñadir_Click()

On Error GoTo CtrlErr
    Select Case SSTab1.Tab
        Case 2
            frm_OFF_MantMoneda.Datos "", "", "", "Añadir"
            frm_OFF_MantMoneda.Show vbModal
            If Not frm_OFF_MantMoneda.bolCancelar Then
                objOFFUsuario.AgregaMoneda frm_OFF_MantMoneda.strCodMoneda, _
                            frm_OFF_MantMoneda.strDesMoneda, _
                            frm_OFF_MantMoneda.strSmbMoneda
            
                grdMoneda.Rebind
            End If
        Case 3
            frm_OFF_MantFormato.Datos "", "", "", "", "Añadir"
            frm_OFF_MantFormato.Show vbModal
            If Not frm_OFF_MantFormato.bolCancelar Then
                objDocumento.AgregaFormato frm_OFF_MantFormato.strCodFormato, _
                        frm_OFF_MantFormato.strDesFormato, _
                        frm_OFF_MantFormato.strCtdAncho, _
                        frm_OFF_MantFormato.strCtdAlto
                grdImpresion.Rebind
            End If
        Case 4
            frm_OFF_MantDocumento.Datos "", "", "", "", "", "Añadir"
            frm_OFF_MantDocumento.Show vbModal
            If Not frm_OFF_MantDocumento.bolCancelar Then
                objDocumento.AgregaTipoDocumento frm_OFF_MantDocumento.strCodDocumento, _
                    frm_OFF_MantDocumento.strDesDocumento, _
                    frm_OFF_MantDocumento.strNumDocumento, _
                    Val(frm_OFF_MantDocumento.strNumLinDocumento), _
                    Val(frm_OFF_MantDocumento.strAnchoDocumento)
                grdDocumento.Rebind
            End If
        Case 5
            frm_OFF_MantFPago.Datos "", "", "Añadir"
            frm_OFF_MantFPago.Show vbModal
            If Not frm_OFF_MantFPago.bolCancelar Then
                objFormaPago.AgregaFormaPago frm_OFF_MantFPago.strCodFormaPago, _
                                        frm_OFF_MantFPago.strDesFormaPago
                grdFormaPago.Rebind
            End If
        
    End Select
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEditar_Click()
On Error GoTo CtrlErr
    Select Case SSTab1.Tab
        Case 2
            frm_OFF_MantMoneda.Datos grdMoneda.Columns(0).Value, grdMoneda.Columns(1).Value, grdMoneda.Columns(2).Value, "Modificar"
            frm_OFF_MantMoneda.Show vbModal
            If Not frm_OFF_MantMoneda.bolCancelar Then
                objOFFUsuario.AgregaMoneda frm_OFF_MantMoneda.strCodMoneda, _
                            frm_OFF_MantMoneda.strDesMoneda, _
                            frm_OFF_MantMoneda.strSmbMoneda
            
                grdMoneda.Rebind
            End If
        Case 3
            frm_OFF_MantFormato.Datos grdImpresion.Columns(0).Value, grdImpresion.Columns(1).Value, Val(grdImpresion.Columns(2).Value), Val(grdImpresion.Columns(3).Value), "Modificar"
            frm_OFF_MantFormato.Show vbModal
            If Not frm_OFF_MantFormato.bolCancelar Then
                objDocumento.AgregaFormato frm_OFF_MantFormato.strCodFormato, _
                        frm_OFF_MantFormato.strDesFormato, _
                        frm_OFF_MantFormato.strCtdAncho, _
                        frm_OFF_MantFormato.strCtdAlto
                grdImpresion.Rebind
            End If
        Case 4
            frm_OFF_MantDocumento.Datos grdDocumento.Columns(0).Value, grdDocumento.Columns(1).Value, grdDocumento.Columns(2).Value, grdDocumento.Columns(3).Value, grdDocumento.Columns(4).Value, "Añadir"
            frm_OFF_MantDocumento.Show vbModal
            If Not frm_OFF_MantDocumento.bolCancelar Then
                objDocumento.AgregaTipoDocumento frm_OFF_MantDocumento.strCodDocumento, _
                    frm_OFF_MantDocumento.strDesDocumento, _
                    frm_OFF_MantDocumento.strNumDocumento, _
                    Val(frm_OFF_MantDocumento.strNumLinDocumento), _
                    Val(frm_OFF_MantDocumento.strAnchoDocumento)
                grdDocumento.Rebind
            End If
        Case 5
            frm_OFF_MantFPago.Datos grdFormaPago.Columns(0).Value, grdFormaPago.Columns(1).Value, "Añadir"
            frm_OFF_MantFPago.Show vbModal
            If Not frm_OFF_MantFPago.bolCancelar Then
                objFormaPago.AgregaFormaPago frm_OFF_MantFPago.strCodFormaPago, _
                                    frm_OFF_MantFPago.strDesFormaPago
                grdFormaPago.Rebind
            End If
            
            
    End Select
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdEliminar_Click()
On Error GoTo CtrlErr
    Select Case SSTab1.Tab
        Case 2
                If grdMoneda.ApproxCount = 0 Then Exit Sub
                grdMoneda.Delete
                grdMoneda.Rebind
        Case 3
                If grdImpresion.ApproxCount = 0 Then Exit Sub
                grdImpresion.Delete
                grdImpresion.Rebind
                
        Case 4
                If grdDocumento.ApproxCount = 0 Then Exit Sub
                grdDocumento.Delete
                grdDocumento.Rebind
        Case 5
                If grdFormaPago.ApproxCount = 0 Then Exit Sub
                grdFormaPago.Delete
                grdFormaPago.Rebind
                
                
    End Select
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
    
    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0
    
    Select Case KeyCode
        Case vbKeyF1
            SSTab1.Tab = 0
            txtCia.SetFocus
        Case vbKeyF2
            SSTab1.Tab = 0
            txtLocal.SetFocus
        Case vbKeyF3
            SSTab1.Tab = 1
            txtGlosa.SetFocus
        Case vbKeyF4
            SSTab1.Tab = 2
            grdMoneda.SetFocus
        Case vbKeyF5
            SSTab1.Tab = 3
            grdImpresion.SetFocus
        Case vbKeyF6
            SSTab1.Tab = 4
            grdDocumento.SetFocus
        Case vbKeyF7
            SSTab1.Tab = 5
            grdFormaPago.SetFocus
        Case vbKeyReturn
            If Shift = 1 Then Call cmdAceptar_Click
    End Select

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName



End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    SetGrid
    
    LoadValues

End Sub


Private Sub LoadValues()

On Error GoTo CtrlErr


    txtCia.Text = objIni.LeerIni(gstrIni, "general", "CIA", "")
    txtEmpresa.Text = objIni.LeerIni(gstrIni, "general", "EMPRESA", "")
    txtRUC.Text = objIni.LeerIni(gstrIni, "general", "RUC", "")
    txtDireccionCia.Text = objIni.LeerIni(gstrIni, "general", "DIRECCION", "")
    txtLocal.Text = objIni.LeerIni(gstrIni, "general", "LOCAL", "")
    txtDireccionLocal.Text = objIni.LeerIni(gstrIni, "general", "DIRECCION_LOCAL", "")
    txtSerieTicket.Text = objIni.LeerIni(gstrIni, "general", "COD_SERIE_ETIQ", "")
    txtSecuencia.Text = objIni.LeerIni(gstrIni, "general", "SEC_VENTA", "")
    txtGlosa.Text = objIni.LeerIni(gstrIni, "GLOSA", "LINE1", "")
    
    grdDocumento.Array1 = objDocumento.TipoDocumento
    grdImpresion.Array1 = objDocumento.Formato
    
    
    grdFormaPago.Array1 = objFormaPago.FormaPago
    


    grdMoneda.Array1 = objOFFUsuario.Moneda



Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub


Private Sub SetGrid()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
  Dim columna As TrueDBGrid70.Column
  
    
    arrCampos = Array("", "", "", "", "")
    arrCaption = Array("Código", "Descripción", "Número", "#Linea", "Ancho")
    arrAncho = Array(900, 3500, 1200, 800, 800)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgRight, dbgRight)
    arrFoco = Array(True, False, False, False, False)

    grdDocumento.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    For Each columna In grdDocumento.Columns
        columna.AllowSizing = False
    Next
    
    
    
    
    
    arrCampos = Array("", "")
    arrCaption = Array("Código", "Descripción")
    arrAncho = Array(900, 5500)
    arrAlineacion = Array(dbgCenter, dbgLeft)
    arrFoco = Array(False, True)
    
    grdFormaPago.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    
    For Each columna In grdFormaPago.Columns
        columna.AllowSizing = False
    Next
    
    
    
    arrCampos = Array("", "", "")
    arrCaption = Array("Código", "Descripción", "Simbolo")
    arrAncho = Array(900, 5500, 1000)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft)
    
    grdMoneda.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    For Each columna In grdMoneda.Columns
        columna.AllowSizing = False
    Next
    
    
    arrCampos = Array("", "", "", "")
    arrCaption = Array("Código", "Descripción", "Ancho", "Alto")
    arrAncho = Array(900, 4000, 1000, 1000)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgLeft)
    
    grdImpresion.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    For Each columna In grdImpresion.Columns
        columna.AllowSizing = False
    Next
    
    
End Sub








Private Sub Form_Unload(Cancel As Integer)

    If MsgBox("Desea salir de la pantalla de mantenimiento ?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Cancel = True: Exit Sub

    Set objIni = Nothing
    Set objDocumento = Nothing
    Set objFormaPago = Nothing



End Sub



Private Sub grdDocumento_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub grdFormaPago_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub grdImpresion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub grdMoneda_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



Private Sub Guardar()
Dim i As Integer
Dim xTemp As New XArrayDB

On Error GoTo CtrlErr


    
    objIni.GuardarIni gstrIni, "general", "CIA", strCia
    objIni.GuardarIni gstrIni, "general", "EMPRESA", strEmpresa
    objIni.GuardarIni gstrIni, "general", "RUC", strRuc
    objIni.GuardarIni gstrIni, "general", "DIRECCION", strDireccion
    objIni.GuardarIni gstrIni, "general", "LOCAL", strLocal
    objIni.GuardarIni gstrIni, "general", "DIRECCION_LOCAL", strDireccionLocal
    objIni.GuardarIni gstrIni, "general", "COD_SERIE_ETIQ", strSerieTicket
    objIni.GuardarIni gstrIni, "general", "SEC_VENTA", strSecuencia
    objIni.GuardarIni gstrIni, "GLOSA", "LINE1", strGlosa
    
    
   Set xTemp = objOFFUsuario.Moneda
    
    For i = 0 To xTemp.UpperBound(1)
        strCodMoneda = strCodMoneda & xTemp(i, 0) & ","
        strDesMoneda = strDesMoneda & xTemp(i, 1) & ","
        strSmbMoneda = strSmbMoneda & xTemp(i, 2) & ","
    Next i
    
    strCodMoneda = left(strCodMoneda, Len(strCodMoneda) - 1)
    strDesMoneda = left(strDesMoneda, Len(strDesMoneda) - 1)
    strSmbMoneda = left(strSmbMoneda, Len(strSmbMoneda) - 1)
    
    objIni.GuardarIni gstrIni, "MONEDA", "COD_MONEDA", strCodMoneda
    objIni.GuardarIni gstrIni, "MONEDA", "DES_MONEDA", strDesMoneda
    objIni.GuardarIni gstrIni, "MONEDA", "SMB_MONEDA", strSmbMoneda
    
    xTemp.Clear
    
    Set xTemp = Nothing
    
    Set xTemp = objFormaPago.FormaPago
    
    For i = 0 To xTemp.UpperBound(1)
        strCodFormaPago = strCodFormaPago & xTemp(i, 0) & ","
        strDesFormaPago = strDesFormaPago & xTemp(i, 1) & ","
    Next i
    strCodFormaPago = left(strCodFormaPago, Len(strCodFormaPago) - 1)
    strDesFormaPago = left(strDesFormaPago, Len(strDesFormaPago) - 1)
    
    objIni.GuardarIni gstrIni, "general", "COD_FORMAPAGO", strCodFormaPago
    objIni.GuardarIni gstrIni, "general", "DES_FORMAPAGO", strDesFormaPago
    
    xTemp.Clear
    
    Set xTemp = Nothing
    
    
    
    Set xTemp = objDocumento.TipoDocumento
    
    For i = 0 To xTemp.UpperBound(1)
        strCodDocumento = strCodDocumento & xTemp(i, 0) & ","
        strDesDocumento = strDesDocumento & xTemp(i, 1) & ","
        strNumDocumento = strNumDocumento & Replace(xTemp(i, 2), "-", "") & ","
        strNumLinDocumento = strNumLinDocumento & xTemp(i, 3) & ","
        strAnchoDocumento = strAnchoDocumento & xTemp(i, 4) & ","
    Next i
    strCodDocumento = left(strCodDocumento, Len(strCodDocumento) - 1)
    strDesDocumento = left(strDesDocumento, Len(strDesDocumento) - 1)
    strNumDocumento = left(strNumDocumento, Len(strNumDocumento) - 1)
    strNumLinDocumento = left(strNumLinDocumento, Len(strNumLinDocumento) - 1)
    strAnchoDocumento = left(strAnchoDocumento, Len(strAnchoDocumento) - 1)
    
    objIni.GuardarIni gstrIni, "general", "COD_DOCUMENTO", strCodDocumento
    objIni.GuardarIni gstrIni, "general", "DES_DOCUMENTO", strDesDocumento
    objIni.GuardarIni gstrIni, "general", "NUM_DOCUMENTO", strNumDocumento
    objIni.GuardarIni gstrIni, "general", "NUM_LIN_DOCUMENTO", strNumLinDocumento
    objIni.GuardarIni gstrIni, "general", "ANCHO_DOCUMENTO", strAnchoDocumento
    
    xTemp.Clear
    
    Set xTemp = Nothing
    
    
    Set xTemp = objDocumento.Formato
    For i = 0 To xTemp.UpperBound(1)
        strCodFormato = strCodFormato & xTemp(i, 0) & ","
        strDesFormato = strDesFormato & xTemp(i, 1) & ","
        strCtdAncho = strCtdAncho & xTemp(i, 2) & ","
        strCtdAlto = strCtdAlto & xTemp(i, 3) & ","
    Next i
    
    strCodFormato = left(strCodFormato, Len(strCodFormato) - 1)
    strDesFormato = left(strDesFormato, Len(strDesFormato) - 1)
    strCtdAncho = left(strCtdAncho, Len(strCtdAncho) - 1)
    strCtdAlto = left(strCtdAlto, Len(strCtdAlto) - 1)
    
    objIni.GuardarIni gstrIni, "IMPRESION", "COD_FORMATO", strCodFormato
    objIni.GuardarIni gstrIni, "IMPRESION", "DES_FORMATO", strDesFormato
    objIni.GuardarIni gstrIni, "IMPRESION", "CTD_ANCHO", strCtdAncho
    objIni.GuardarIni gstrIni, "IMPRESION", "CTD_ALTO", strCtdAlto
    
    
    xTemp.Clear
    
    Set xTemp = Nothing
    
Exit Sub

CtrlErr:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub
