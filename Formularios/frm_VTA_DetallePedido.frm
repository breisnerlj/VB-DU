VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_VTA_DetallePedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Pedido"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12735
   Icon            =   "frm_VTA_DetallePedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   375
      Left            =   11760
      TabIndex        =   59
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton CmdEntregaTercero 
      Caption         =   "&Datos Entrega a Tercero"
      Height          =   375
      Left            =   2340
      TabIndex        =   55
      Top             =   8840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton CmdObs 
      Caption         =   "&Observaciones"
      Height          =   375
      Left            =   120
      TabIndex        =   53
      Top             =   8840
      Width           =   2055
   End
   Begin VB.TextBox lblPedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "1324567890"
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdCliente 
      Caption         =   "&Cliente"
      Height          =   375
      Left            =   10800
      TabIndex        =   36
      Top             =   9360
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   9840
      TabIndex        =   35
      Top             =   9360
      Width           =   855
   End
   Begin VB.TextBox txtObservacion 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   9240
      Width           =   9615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Detalles de Transferencia"
      Height          =   1575
      Left            =   60
      TabIndex        =   6
      Top             =   7200
      Width           =   12615
      Begin vbp_Ventas.ctlGrilla grdTransferencias 
         Height          =   1215
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   2143
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Detalle del Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   60
      TabIndex        =   5
      Top             =   5280
      Width           =   12615
      Begin vbp_Ventas.ctlGrilla grdDetalle 
         Height          =   1455
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   2566
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Formas de Pago"
      Height          =   1455
      Left            =   60
      TabIndex        =   4
      Top             =   3720
      Width           =   12615
      Begin vbp_Ventas.ctlGrilla grdFormaPago 
         Height          =   1095
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   1931
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Pedido"
      Height          =   2715
      Left            =   6420
      TabIndex        =   3
      Top             =   0
      Width           =   6255
      Begin VB.Label lblBTLAsignada 
         Caption         =   "BTL Asignada"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   58
         Top             =   1164
         Width           =   1335
      End
      Begin VB.Label lblConvenio 
         Caption         =   "M?todo de entrega :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblFechaPactada 
         Height          =   255
         Left            =   2040
         TabIndex        =   56
         Top             =   2400
         Width           =   4035
      End
      Begin VB.Label lblOperadora 
         Caption         =   "Teleoperadora"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   52
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblRuteador 
         Caption         =   "Ruteador"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   51
         Top             =   548
         Width           =   3735
      End
      Begin VB.Label lblMotorizado 
         Caption         =   "Motorizado"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   50
         Top             =   856
         Width           =   3735
      End
      Begin VB.Label lblBTLAsignada 
         Caption         =   "BTL Asignada"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   49
         Top             =   1164
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblDocumentoLocal 
         Caption         =   "Documento Local"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   48
         Top             =   1472
         Width           =   1455
      End
      Begin VB.Label LblCovenio 
         Caption         =   "Convenio"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   47
         Top             =   1785
         Width           =   3135
      End
      Begin VB.Label LblCoPago 
         Caption         =   "CoPago"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   46
         Top             =   2088
         Width           =   1335
      End
      Begin VB.Label lblConvenio 
         Caption         =   "Co Pago :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   2088
         Width           =   1335
      End
      Begin VB.Label lblConvenio 
         Caption         =   "Convenio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   1780
         Width           =   1335
      End
      Begin VB.Label LblModalidad 
         Caption         =   "Modalidad"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   39
         Top             =   1164
         Width           =   2175
      End
      Begin VB.Label lblModalid 
         Caption         =   "Modalidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   38
         Top             =   1164
         Width           =   975
      End
      Begin VB.Label lblNumeroItems 
         Caption         =   "000"
         Height          =   255
         Index           =   1
         Left            =   5820
         TabIndex        =   31
         Top             =   1785
         Width           =   375
      End
      Begin VB.Label LblIngresoFecha 
         Caption         =   "F. Ingreso"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   30
         Top             =   1472
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblTransferencia 
         Caption         =   "Transferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   29
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblNumeroItems 
         Caption         =   "N? de Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   19
         Top             =   1785
         Width           =   1095
      End
      Begin VB.Label lblDocumentoLocal 
         Caption         =   "Documento :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1472
         Width           =   1335
      End
      Begin VB.Label LblIngresoFecha 
         Caption         =   "F. Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   17
         Top             =   1472
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblMotorizado 
         Caption         =   "Motorizado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label lblBTLAsignada 
         Caption         =   "Local Asignado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1170
         Width           =   1455
      End
      Begin VB.Label lblRuteador 
         Caption         =   "Ruteador :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   548
         Width           =   1335
      End
      Begin VB.Label lblOperadora 
         Caption         =   "Teleoperadora :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiempos"
      Height          =   855
      Left            =   60
      TabIndex        =   2
      Top             =   2820
      Width           =   12615
      Begin vbp_Ventas.ctlGrilla grdTiempos 
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   873
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame frmCliente 
      Caption         =   "Datos del Cliente"
      Height          =   2355
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   6255
      Begin VB.Label lblUrba 
         Caption         =   "Urbanizaci?n"
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   44
         Top             =   1140
         Width           =   4335
      End
      Begin VB.Label lblDist 
         Caption         =   "Distrito"
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   43
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label lblUrbanizacion 
         BackStyle       =   0  'Transparent
         Caption         =   "Urbanizaci?n :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   42
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label lblDistrito 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   75
         TabIndex        =   41
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label lblRUC 
         Caption         =   "RUC/R.Soc."
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   28
         Top             =   2040
         Width           =   4635
      End
      Begin VB.Label lblReferencia 
         Caption         =   "Referencia"
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   27
         Top             =   1740
         Width           =   4635
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   25
         Top             =   540
         Width           =   4605
      End
      Begin VB.Label lblFechaIngreso 
         Caption         =   "f. Ingreso"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   24
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label lblTelefono 
         Caption         =   "lblTelefono"
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
         Index           =   1
         Left            =   1380
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFechaIngreso 
         Caption         =   "F. Ingreso"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblRUC 
         BackStyle       =   0  'Transparent
         Caption         =   "RUC/R.Soc. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   11
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label lblReferencia 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   10
         Top             =   1740
         Width           =   1305
      End
      Begin VB.Label lblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   8
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label lblTelefono 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel?fono :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblDireccion 
         Caption         =   "Direcci?n"
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   26
         Top             =   840
         Width           =   4470
      End
      Begin VB.Label lblDireccion 
         BackStyle       =   0  'Transparent
         Caption         =   "Direcci?n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   9
         Top             =   840
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   9480
      TabIndex        =   33
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label LblConexSalud 
      Alignment       =   2  'Center
      Caption         =   "Conexi?n Salud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3120
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido N?"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   98
      Width           =   840
   End
End
Attribute VB_Name = "frm_VTA_DetallePedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NumeroPedido As String
Public CodigoLocal As String
Private CodigoEmpresa As String
Private CodigoCliente As String
Private strCodDireccionCli As String
Dim objPedido As New clsProforma
Dim rsDetalle As oraDynaset
Dim rsCabecera As oraDynaset

Private Sub cmdCliente_Click()
    frm_VTA_Cliente.ctlCliente1.Cargar
    frm_VTA_Cliente.ctlCliente1.CodDireccionCli = strCodDireccionCli
    frm_VTA_Cliente.ctlCliente1.ConsultaCliente CodigoCliente
    frm_VTA_Cliente.CargarValores
    frm_VTA_Cliente.Show vbModal
End Sub

Private Sub CmdEntregaTercero_Click()
    frm_DLV_EntregaTercero.Show vbModal
End Sub

Private Sub cmdExcel_Click()
On Error GoTo Control
    Dim FileNm As Variant
    Dim path_link As Variant
    With CommonDialog1 'Lets you call common dialog pretty much
        '.InitDir = App.Path 'Where you want the program to start to show to save
        '.FileName = "" 'If you only want to save as a specific file, put something in between quotes
        .Filter = "Libro de Excel 97-2003 (*.xls)|*.xls|Archivo de Excel (*.xlsx)|*.xlsx|CSV (Delimitado por comas) (*.csv)|*.csv"
        .DialogTitle = "Ingrese el nombre del archivo" 'Just what you want the caption to say
        .CancelError = True
        .ShowSave 'Makes it look like it saved
        
        FileNm = .FileName
        If Trim(FileNm) = "" Then MsgBox "No se encontro el archivo.", vbOKOnly + vbExclamation, "Error": Exit Sub
        If Dir(FileNm) <> "" Then
            If MsgBox("Ya existe un archivo con el nombre especificado." & Chr(13) & "Desea continuar ?", vbYesNo + vbQuestion, "Confirme") = vbNo Then Exit Sub
        End If
'        If Err.Number = 32755 Then
'            MsgBox "Archivo no se guardara, se cancelo o cerro la venta.", vbExclamation, App.ProductName
'            Exit Sub
'        End If

    End With
    Dim rsCabExcel As oraDynaset
    Dim rsDetExcel As oraDynaset
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim countDet As Integer
'    For VerticalAlignment:
'    Top:    -4160
'    Center: -4108
'    Bottom: -4107
'    And HorizontalAlignment:
'
'    Left:    -4131
'    Center:  -4108
'    Right:   -4152

    Set rsCabExcel = objPedido.ListaCabecera(CodigoEmpresa, CodigoLocal, NumeroPedido)
    Set rsDetExcel = objPedido.ListaDetalle(CodigoEmpresa, CodigoLocal, NumeroPedido)
    
    'Debug.Print rsCabExcel.RecordCount
    countDet = rsDetExcel.RecordCount
    
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    oSheet.name = right(FileNm, Len(FileNm) - InStrRev(FileNm, "\"))
    oSheet.Activate
    'oSheet.Range("A1", "Z1000").Font.name = "Arial"
    
    'CABECERA
    With oSheet.Cells(4, 2)
        .Value = "CLIENTE:"
        .Borders.LineStyle = 1
    End With
    With oSheet.Cells(5, 2)
        .Value = "FECHA:"
        .Borders.LineStyle = 1
    End With
    With oSheet.Cells(6, 2)
        .Value = "PEDIDO:"
        .Borders.LineStyle = 1
    End With
    
    With oSheet.Cells(4, 3)
        .Value = rsCabExcel("DES_CLIENTE").Value
        .Columns.ColumnWidth = 38
        .HorizontalAlignment = -4131
        .Borders.LineStyle = 1
    End With
    
    With oSheet.Cells(5, 3)
        .NumberFormat = "@"
        .Value = rsCabExcel("FCH_REGISTRA").Value
        .HorizontalAlignment = -4152
        .Borders.LineStyle = 1
    End With
    
    With oSheet.Cells(6, 3)
        .NumberFormat = "@"
        .Value = rsCabExcel("NUM_PROFORMA").Value
        .HorizontalAlignment = -4152
        .Borders.LineStyle = 1
    End With
    
    'DETALLE
    With oSheet.Range("A9:F9")
        .Value = Array("C?digo", "Items", "Descripci?n", "Unidad", "Unitario", "Total")
        .HorizontalAlignment = -4131
        .Borders.LineStyle = 1
    End With
    
    ReDim dataArray(1 To countDet, 1 To 6) As Variant
    Dim r As Integer
    'Dim totalRow As String
    'Cell1 = "A10"
    rsCabExcel.MoveFirst
    r = 1
    While Not rsDetExcel.EOF
        dataArray(r, 1) = rsDetExcel("COD_PRODUCTO").Value
        dataArray(r, 2) = rsDetExcel("CANT").Value
        dataArray(r, 3) = rsDetExcel("DES_PRODUCTO").Value
        dataArray(r, 4) = " " 'PKG NO DEVUELVE ESTE DATO [UNIDAD]
        dataArray(r, 5) = rsDetExcel("PRC_UNIT_VTA").Value
        dataArray(r, 6) = rsDetExcel("MTO_SUBTOTAL").Value
        r = r + 1
        rsDetExcel.MoveNext
    Wend
    r = 1
'    For r = 1 To countDet
'        dataArray(r, 1) = rsDetExcel("COD_PRODUCTO").Value
'        dataArray(r, 2) = rsDetExcel("CANT").Value
'        dataArray(r, 3) = rsDetExcel("DES_PRODUCTO").Value
'        dataArray(r, 4) = " " 'PKG NO DEVUELVE ESTE DATO [UNIDAD]
'        dataArray(r, 5) = rsDetExcel("PRC_UNIT_VTA").Value
'        dataArray(r, 6) = rsDetExcel("MTO_SUBTOTAL").Value
'    Next
    
    With oSheet.Range("A10").Resize(countDet, 6)
        .Value = dataArray
        .Borders.LineStyle = 1
    End With
    
    Dim rowTotal As Integer
    rowTotal = 10 + countDet
    
    With oSheet.Cells(rowTotal, 5)
        .Value = "TOTAL"
        .Borders.LineStyle = 1
        .Interior.Color = RGB(255, 255, 0)
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
        .HorizontalAlignment = -4108
    End With
    
    With oSheet.Cells(rowTotal, 6)
        .Value = rsCabExcel("MTO_TOTAL").Value
        .Borders.LineStyle = 1
        .Interior.Color = RGB(255, 255, 0)
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
    
    rowTotal = rowTotal + 3
    ReDim arrayInfo(1 To 4, 1 To 1) As Variant
    arrayInfo(1, 1) = "Los precios incluyen IGV"
    arrayInfo(2, 1) = "Precios sujeto a cambios"
    arrayInfo(3, 1) = "Moneda Nuevos Soles"
    arrayInfo(4, 1) = "Contraentrega"
    
    ReDim arrayInfo2(1 To 13, 1 To 1) As Variant
    arrayInfo2(1, 1) = "NOTA IMPORTANTE"
    arrayInfo2(2, 1) = "Se realiza el envio ,previa verificaci?n  x horario y zona de cobertura"
    arrayInfo2(3, 1) = "Las cotizaciones son realizadas en base al stock de nuestro Almacen"
    arrayInfo2(4, 1) = "No realizamos entrega de ficha t?cnica .Rs"
    arrayInfo2(5, 1) = "Forma de pago:"
    arrayInfo2(6, 1) = "* Efectivo"
    arrayInfo2(7, 1) = "*Transferencia bancaria"
    arrayInfo2(8, 1) = "*Tarjetas de cr?dito"
    arrayInfo2(9, 1) = "*Abono en cheque"
    arrayInfo2(10, 1) = "Una vez realizado el abono en cheque ,enviarnos los datos de la transacci?n realizada ,"
    arrayInfo2(11, 1) = "las misma que es remitida a nuestro Dpto de Contabilidad(la verificaci?n demora en ser visualizadas"
    arrayInfo2(12, 1) = "hasta 3 d?as utiles) .Con la confirmacion de nuestro Dpto de Contabilidad ,se coordina inmediatamente con Ud"
    arrayInfo2(13, 1) = "la hora de entrega del pedido."
    
    With oSheet.Range("C" & rowTotal).Resize(4, 1)
        .Value = arrayInfo
        .Borders.LineStyle = 1
        .Font.Bold = True
        .Font.Size = 8
        .Font.name = "Arial"
    End With
    rowTotal = rowTotal + 1
    With oSheet.Range("C" & rowTotal)
        .Interior.Color = RGB(255, 230, 102)
        .Font.Color = RGB(255, 0, 0)
    End With
    rowTotal = rowTotal + 3
    With oSheet.Range("C" & rowTotal & ":G" & rowTotal)
        .MergeCells = True
        .Value = "Nuestro servicio del delivery tiene un recargo de s/5.00 que estaran incluidos dentro del pedido/factura/boleta"
        .Interior.Color = RGB(255, 255, 0)
        .Font.Size = 8
        .Font.Bold = True
    End With
    rowTotal = rowTotal + 1
    oSheet.Range("C" & rowTotal & ":G" & rowTotal).Interior.Color = RGB(255, 255, 0)
    With oSheet.Range("C" & rowTotal)
        .Value = "A partir de 1 de Agosto nos unimos al cuidado de nuestro planeta , por tal motivo estamos realizando el cobro por la entrega de bolsa plastica a 0.25 centimos bolsa mediana y 0.30 bolsa grande."
        .Font.Size = 8
        .Font.Bold = True
    End With
    rowTotal = rowTotal + 2
    Dim i As Integer
    For i = 1 To UBound(arrayInfo2)
        'Debug.Print arrayInfo2(i, 1)
        With oSheet.Range("B" & rowTotal)
            .Value = arrayInfo2(i, 1)
            .Font.name = "Arial"
            .Font.Bold = True
            .Font.Size = 8
        End With
        
        With oSheet.Range("B" & rowTotal & ":E" & rowTotal)
            .Borders(7).LineStyle = 1
            .Borders(8).LineStyle = 1
            .Borders(9).LineStyle = 1
            .Borders(10).LineStyle = 1
            
            If i = 1 Or i = 4 Or i = 5 Then .MergeCells = True
            If i = 1 Or i = 5 Then .HorizontalAlignment = -4108
            If i = 1 Then .Interior.Color = RGB(255, 230, 102)
            If i = 9 Then .Font.Color = RGB(102, 128, 255)
        End With
        rowTotal = rowTotal + 1
    Next
    
    oExcel.DisplayAlerts = False
    oBook.SaveAs FileNm '"C:\Documents and Settings\aescate\Escritorio\vbp_ventas_dev\drf.xls"
    'oBook.SaveCopyAs FileNm
    oExcel.DisplayAlerts = True
    oExcel.Quit
    
    MsgBox "Archivo guardado correctamente", vbExclamation, App.ProductName
    
    Exit Sub
Control:
    Debug.Print "Error.Number : " & Err.Number
    Select Case Err.Number
        Case 32755
            MsgBox "Archivo no se guardara, se cancelo o cerro la venta.", vbExclamation, App.ProductName
        Case 429
            MsgBox "Para ejecutar esta funcionalidad debe tener Excel instalado.", vbExclamation, App.ProductName
        Case Else
            MsgBox Err.Description & vbNewLine & Err.Number, vbOKOnly + vbExclamation, App.ProductName
    End Select
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo Handle
Dim strMensaje As String
strMensaje = objPedido.GrabarReclamo(CodigoEmpresa, CodigoLocal, NumeroPedido, txtObservacion.Text, objUsuario.Codigo)
If strMensaje = "" Then
    MsgBox "Se grabo satisfactoriamente", vbExclamation, App.ProductName
    'Unload Me
Else
    MsgBox strMensaje, vbCritical, App.ProductName
End If
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub CmdObs_Click()
    frm_ADM_ObsDetPedido.Show vbModal
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Control
    'cmdExcel debe mostrarse solo si es pedido guardado y tipo maquina cabina [003]
    'If objUsuario.TipoMaquina = objUsuario.TipoMaquinaCabina Then: cmdExcel.Visible = True: cmdExcel.Enabled = True
    'CodigoEmpresa = objUsuario.CodigoEmpresa
    'CargaCabecera
    'CargaDetalle
    'CargaFormaPago
    'CargaTiempos
    'CargaTransf

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub


Private Sub cargaDetalle(Optional vstrCodigoLocal As String, Optional vstrNumeropedido As String)
    Dim odynClon As oraDynaset
    
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO_2", _
                      "CANT", "FLG_FRACCIONAMIENTO", _
                      "prc_unit_kairo", "PCT_DSCT_UNIT", _
                      "PRC_UNIT_VTA", "MTO_SUBTOTAL")
                      
    arrCaption = Array("Codigo", "Descripci?n", _
                       "Cantidad", "Frac.", _
                       "PVP", "D.N.", _
                       "P.V.Px", "SubTotal")
                       
    arrAncho = Array(0, 7000, _
                     1000, 1000, _
                     1000, 1000, _
                     1000, 1000)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, _
                          dbgRight, dbgRight, _
                          dbgRight, dbgRight, _
                          dbgRight, dbgRight)
    
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    'grdDetalle.ColumnFooter = True
    
    Set rsDetalle = objPedido.ListaDetalle(CodigoEmpresa, vstrCodigoLocal, vstrNumeropedido)
    Set grdDetalle.DataSource = rsDetalle
    
    Dim dblPrecioUnit As Double
    Dim i As Integer

    dblPrecioUnit = 0
    Set odynClon = rsDetalle.Clone
    odynClon.MoveFirst
    For i = 0 To odynClon.RecordCount - 1
        dblPrecioUnit = dblPrecioUnit + odynClon("MTO_SUBTOTAL").Value
        odynClon.MoveNext
    Next i
        
    grdDetalle.Columns("FLG_FRACCIONAMIENTO").Visible = False
    Frame4.Caption = "Suma del Detalle del Pedido =>" & "  " & Format(dblPrecioUnit, "###,##0.00")
    lblNumeroItems(1).Caption = grdDetalle.ApproxCount
End Sub
Private Sub CargaTiempos(Optional vstrNumeropedido As String)
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("Generado", "Verificado", _
                      "Asignado", "Proforma", _
                      "Llevando", "Llegada", _
                      "Entregado", "Llegada a local", _
                      "Liberado", "Anulado", _
                      "Time")
                      
    arrCaption = Array("Generado", "Verificado", _
                       "Asignado", "Proforma", _
                       "Llevando", "Llegada", _
                       "Entregado", "Llegada a local", _
                       "Liberado", "Anulado", _
                       "Time")
                       
    arrAncho = Array(1000, 1000, _
                     1000, 1000, _
                     1000, 1000, _
                     1000, 1000, _
                     1000, 1000, _
                     1000)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft)
                          
    grdTiempos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    Set grdTiempos.DataSource = objPedido.ListaEstados(CodigoEmpresa, vstrNumeropedido)
End Sub
Private Sub CargaCabecera(Optional vstrCodigoLocal As String, Optional vstrNumeropedido As String)
On Error GoTo Handle
    Set rsCabecera = objPedido.ListaCabecera(CodigoEmpresa, vstrCodigoLocal, vstrNumeropedido)
    
    lblPedido.Text = vstrNumeropedido
    
    If rsCabecera("FLG_ENTREGA_TERCERO").Value = "1" Then
        LblConexSalud.Visible = True
        CmdEntregaTercero.Visible = True
      Else
        LblConexSalud.Visible = False
        CmdEntregaTercero.Visible = False
    End If
    
    lblTelefono(1).Caption = "" & rsCabecera("DES_AUX_CLI_TLF").Value
    lblFechaIngreso(1).Caption = "" & rsCabecera("FCH_CREA_CLI").Value
    lblDireccion(1).Caption = "" & rsCabecera("DES_SUFIJO").Value & " " & rsCabecera("DIRECCION").Value
    lblNombre(1).Caption = "" & rsCabecera("DES_CLIENTE").Value
    txtObservacion.Text = "" & rsCabecera("OBS_RECLAMO").Value
    CodigoCliente = "" & rsCabecera("COD_CLIENTE_DLV").Value
    lblReferencia(1).Caption = "" & rsCabecera("DES_REFERENCIA").Value
    If "" & rsCabecera("COD_TIPO_DOCUMENTO").Value = "FAC" Then
        lblRUC(1).Caption = "" & rsCabecera("RUC").Value & "/" & " " & rsCabecera("DES_RAZON_SOCIAL").Value
    Else
        lblRUC(1).Caption = ""
    End If
    lblOperadora(1).Caption = "" & rsCabecera("CAJERA").Value
    lblRuteador(1).Caption = "" & objPedido.fndevNomRuteador(objUsuario.CodigoEmpresa, NumeroPedido)
    lblMotorizado(1).Caption = "" & rsCabecera("DES_MOTORIZADO").Value
    lblBTLAsignada(1).Caption = "" & rsCabecera("COD_LOCAL_REF").Value
    lblBTLAsignada(2).Caption = "" & rsCabecera("COD_LOCAL_SAP_REF").Value
    lblDocumentoLocal(1).Caption = "" & rsCabecera("COD_TIPO_DOCUMENTO").Value
    LblModalidad(2).Caption = "" & rsCabecera("DES_MODALIDAD").Value
    LblCovenio(0).Caption = "" & rsCabecera("DES_CONVENIO").Value
    lblTransferencia(1).Caption = IIf("" & rsCabecera("COD_TIPO_DOC").Value = "PRO", "", "TRANSFERENCIA")
    lblDist(2).Caption = "" & rsCabecera("DES_DISTRITO").Value
    lblUrba(2).Caption = "" & rsCabecera("DES_URBANIZACION").Value
    ''LblCoPago(1).Caption = "" & rsCabecera("COPAGO").Value
    LblCoPago(1).Caption = IIf(IsNull(rsCabecera("PCT_BENEFICIARIO").Value), 0, rsCabecera("PCT_BENEFICIARIO").Value) * IIf(IsNull(rsCabecera("MTO_TOTAL").Value), 0, rsCabecera("MTO_TOTAL").Value) / 100
    
'    lblFechaPactada.Caption = IIf(IIf(IsNull(rsCabecera!FLG_FECHA_PACTADA), "", rsCabecera!FLG_FECHA_PACTADA) = "1", "SI - ", "NO") & IIf(IsNull(rsCabecera!FCH_HORA_PACT_ENTR), "", rsCabecera!FCH_HORA_PACT_ENTR)
    'lblFechaPactada.Caption = IIf(IIf(IsNull(rsCabecera!FLG_FECHA_PACTADA), "", rsCabecera!FLG_FECHA_PACTADA) = "1", "SI - ", "NO") & IIf(IsNull(rsCabecera!FCH_HORA_PACT_ENTR), "", (rsCabecera!DELIVERY_TYPE) & " " & rsCabecera!FCH_HORA_PACT_ENTR) & " - " & (rsCabecera!HORA_SEGUNDA_PACT_ENTR)
    lblFechaPactada.Caption = IIf(IsNull(rsCabecera!FCH_HORA_PACT_ENTR), "", rsCabecera!DELIVERY_TYPE & " " & Format(rsCabecera!FCH_HORA_PACT_ENTR, "dddd, dd mmmm") & Format(rsCabecera!FCH_HORA_PACT_ENTR, " - hh:mm am/pm") & Format(rsCabecera!HORA_SEGUNDA_PACT_ENTR, " a hh:mm am/pm"))
    
    frm_ADM_ObsDetPedido.pObsLocal = "" & rsCabecera("OBS_NOTA_LOCAL").Value
    frm_ADM_ObsDetPedido.pObsMotorizado = "" & rsCabecera("OBS_NOTA_MOTORIZADO").Value
    frm_ADM_ObsDetPedido.pObsRuta = "" & rsCabecera("OBS_NOTA_RUTEO").Value
    frm_ADM_ObsDetPedido.pObsVerificacion = "" & rsCabecera("OBS_NOTA_VERIFICACION").Value
    
    frm_DLV_EntregaTercero.pAuxNomb = "" & rsCabecera("DES_AUX_RECOGE_NOMBRE").Value
    frm_DLV_EntregaTercero.pAuxDirecc = "" & rsCabecera("DES_AUX_RECOGE_DIRECC").Value
    frm_DLV_EntregaTercero.pAuxRefer = "" & rsCabecera("DES_AUX_RECOGE_REF").Value
    frm_DLV_EntregaTercero.pAuxTelefono = "" & rsCabecera("DES_AUX_RECOGE_TLF").Value
    frm_DLV_EntregaTercero.pAuxDistrito = "" & rsCabecera("DISTRITO").Value
    
    strCodDireccionCli = "" & rsCabecera("COD_DIRECCION_CLI").Value
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
Private Sub CargaFormaPago(Optional vstrCodigoLocal As String, Optional vstrNumeropedido As String)
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("DES_FORMA_PAGO", _
                      "DES_HIJO", _
                      "VERIF_POS", _
                      "IMP_SIN_REDONDEO", _
                      "IMP_TIPO_CAMBIO", _
                      "IMP_MONEDA_NAC", _
                      "IMP_VUELTO", _
                      "NUM_TARJETA", _
                      "NUM_AUTORIZACION", _
                      "NUM_CUOTAS", _
                      "FCH_VENCIMIENTO", _
                      "NUM_DOCUMENTO_IDENT", _
                      "DES_NOM_TITULAR" _
                      )
                      
    arrCaption = Array("Descripci?n", _
                       "Detalle", _
                       "Verificaci?n", _
                       "Pagado", _
                       "T/C", _
                       "Importe Soles", _
                       "Vuelto", _
                       "Numero Tarjeta", _
                       "Autorizaci?n", _
                       "Cuotas", _
                       "F. Venc", _
                       "Doc. Ident.", _
                       "Titular" _
                       )
                       
    arrAncho = Array(1400, _
                     1400, _
                     1400, _
                     800, _
                     800, _
                     800, _
                     800, _
                     1800, _
                     1000, _
                     500, _
                     1000, _
                     1000, _
                     2000 _
                     )
                     
    arrAlineacion = Array(vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft _
                          )
                          
    grdFormaPago.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    Set grdFormaPago.DataSource = objPedido.ListaFormaPago(CodigoEmpresa, vstrCodigoLocal, vstrNumeropedido)
End Sub

Private Sub CargaTransf(Optional vstrNumeropedido As String)
    
    Dim rs As oraDynaset
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("ITEM", "NUM_PROFORMA", _
                      "LOCAL_ORIGEN", "COD_LOCAL", "DES_LOCAL", "DES_ESTADO_PEDIDO")
                      
    arrCaption = Array("Item", "Proforma", _
                       "Local Origen", "Cod Local", "Des. Local", "Estado")
                       
    arrAncho = Array(500, 1800, _
                     1600, 0, 2400, 1500)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
    
    grdTransferencias.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdTransferencias.Columns(3).Visible = False
    Set rs = objPedido.ListaTransfPedido(objUsuario.CodigoEmpresa, vstrNumeropedido, objUsuario.CodigoLocal)
    Set grdTransferencias.DataSource = rs
    Frame5.Caption = "Detalle de Transferencias" & "  " & "Total =>" & "  " & rs.RecordCount
End Sub


Function CaracterPosicion(ByVal Cadena As String, ByVal Caracter As String, ByVal Longuitud As Integer) As String
Dim strResto As String
Dim strAuxiliar As String
Dim strCadena As String
On Error GoTo Handle
strResto = Cadena
    While strResto <> ""
        strAuxiliar = left(strResto, Longuitud)
        strResto = right(Cadena, Len(strResto) - 4)
        strCadena = strCadena & strAuxiliar & Caracter
    Wend
    CaracterPosicion = strCadena
    Exit Function
Handle:
    strCadena = strCadena & strResto
    CaracterPosicion = strCadena
End Function

Public Sub ReCargaDetPedido()
    CodigoEmpresa = objUsuario.CodigoEmpresa
    If grdTransferencias Is Nothing Then Exit Sub
    If grdTransferencias.ApproxCount <= 0 Then
        CargaCabecera CodigoLocal, NumeroPedido
        cargaDetalle CodigoLocal, NumeroPedido
        CargaFormaPago CodigoLocal, NumeroPedido
        CargaTiempos NumeroPedido
        CargaTransf NumeroPedido
      Else
'        CargaCabecera grdTransferencias.Columns("LOCAL_ORIGEN").Value, grdTransferencias.Columns("NUM_PROFORMA").Value
'        CargaDetalle grdTransferencias.Columns("LOCAL_ORIGEN").Value, grdTransferencias.Columns("NUM_PROFORMA").Value
'        CargaFormaPago grdTransferencias.Columns("LOCAL_ORIGEN").Value, grdTransferencias.Columns("NUM_PROFORMA").Value
        CargaCabecera grdTransferencias.Columns("COD_LOCAL").Value, grdTransferencias.Columns("NUM_PROFORMA").Value
        cargaDetalle grdTransferencias.Columns("COD_LOCAL").Value, grdTransferencias.Columns("NUM_PROFORMA").Value
        CargaFormaPago grdTransferencias.Columns("COD_LOCAL").Value, grdTransferencias.Columns("NUM_PROFORMA").Value
        CargaTiempos grdTransferencias.Columns("NUM_PROFORMA").Value
        CargaTransf grdTransferencias.Columns("NUM_PROFORMA").Value
    End If
End Sub
    
Private Sub grdTransferencias_DblClick()
    On Error GoTo Handle
    Dim frm As New frm_VTA_DetallePedido
    If grdTransferencias.ApproxCount <= 0 Then Exit Sub
    
    frm.NumeroPedido = grdTransferencias.DataSource("NUM_PROFORMA")
    frm.CodigoLocal = grdTransferencias.DataSource("COD_LOCAL")
    frm.ReCargaDetPedido
    frm.Show vbModal
    Set frm = Nothing
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

