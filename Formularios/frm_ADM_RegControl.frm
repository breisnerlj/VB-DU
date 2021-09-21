VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_RegControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de la Competencia"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Mi Local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   9495
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   64880641
         CurrentDate     =   40007
      End
      Begin vbp_Ventas.ctlDataCombo cboCaja 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlTextBox txtDiaUltDoc1BTL 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
      Begin vbp_Ventas.ctlTextBox txtDiaUltDocUltBTL 
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
      Begin vbp_Ventas.ctlTextBox txtDia1Doc1BTL 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Tipo            =   7
         Enabled         =   0   'False
         TABAuto         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtDia1DocUltBTL 
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Tipo            =   7
         Enabled         =   0   'False
         TABAuto         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtValePromedio 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Tipo            =   4
         Enabled         =   0   'False
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
         Locked          =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caja"
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   33
         Top             =   780
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Vale Promedio"
         Height          =   255
         Left            =   5940
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Documento"
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
         Left            =   1680
         TabIndex        =   30
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Último Documento"
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
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Primer dia del Mes"
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
         Height          =   195
         Left            =   6720
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ulitmo dia del Mes"
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
         Height          =   195
         Left            =   600
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   1058
      ModoBotones     =   6
      EnabledEfecto   =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Competencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   9495
      Begin vbp_Ventas.ctlTextBox txtObservacion 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1085
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
      Begin vbp_Ventas.ctlDataCombo cboLocalCompe 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin vbp_Ventas.ctlDataCombo cboEmpresaCompetencia 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox txtDia1Doc1 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
      Begin vbp_Ventas.ctlTextBox txtDia1DocUlt 
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
      Begin vbp_Ventas.ctlDataCombo cboCajas 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin vbp_Ventas.ctlGrillaArray grdDocumento 
         Height          =   2655
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4683
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin vbp_Ventas.ctlTextBox txtDiaUltDocUlt 
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   4320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Tipo            =   7
         Enabled         =   0   'False
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
      Begin vbp_Ventas.ctlTextBox txtDiaUltDoc1 
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   4320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Tipo            =   7
         Enabled         =   0   'False
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Observación: "
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Números :"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   24
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo Documento"
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
         Left            =   3840
         TabIndex        =   23
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Documento"
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
         Left            =   1440
         TabIndex        =   22
         Top             =   1560
         Width           =   1620
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9360
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label6 
         Caption         =   "Cajas :"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa :"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Local :"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   630
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ulitmo dia del Mes"
         Enabled         =   0   'False
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
         Height          =   195
         Left            =   1800
         TabIndex        =   25
         Top             =   4320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_ADM_RegControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objControl As New clsCntCompetencia
Dim XDatos As New XArrayDB

Sub CargaDatos()
        
       Set cboEmpresaCompetencia.RowSource = objControl.ListaCompetencia(objUsuario.CodigoLocal)
       cboEmpresaCompetencia.BoundColumn = "RUC_PROVEEDOR"
       cboEmpresaCompetencia.ListField = "NOM_PROVEEDOR"
       cboEmpresaCompetencia.BoundText = ""
       Dim objMaquina As New clsMaquina
       Set cboCaja.RowSource = objMaquina.MaquinaLocal(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
       cboCaja.BoundColumn = "COD"
       cboCaja.ListField = "DES"
       cboCaja.BoundText = ""
        dtpFecha.Value = objUsuario.sysdate - 1
       dtpFecha.MaxDate = objUsuario.sysdate - 1
       XDatos.ReDim 0, -1, 0, 15
End Sub

Private Sub cboCaja_Change()
    BuscaDatosCaja
End Sub

Private Sub cboCajas_Change()
    bloquea
End Sub

Private Sub cboEmpresaCompetencia_Change()
        If Not fncGuion(cboEmpresaCompetencia.BoundText, 0, "-") = "" Then
       Set cboLocalCompe.RowSource = objControl.ListaCompetencia(objUsuario.CodigoLocal, fncGuion(cboEmpresaCompetencia.BoundText, 0, "-"))
       
        cboLocalCompe.BoundColumn = "COD_LOCAL"
        cboLocalCompe.ListField = "NOM_LOCAL"
        cboLocalCompe.Enabled = True
        Else
        cboLocalCompe.Enabled = False
    End If
            cboLocalCompe.BoundText = ""
       
        cboCajas.Enabled = False

        cboCajas.BoundText = ""
        bloquea
End Sub

Private Sub cboLocalCompe_Change()
    If Not cboLocalCompe.BoundText = "" Then
           Set cboCajas.RowSource = objControl.ListaCajas(fncGuion(cboLocalCompe.BoundText, 0, "-"), fncGuion(cboEmpresaCompetencia.BoundText, 0, "-"))
            cboCajas.BoundColumn = "ID"
            cboCajas.ListField = "DES_CAJA"
'            lblDireccion.Caption = fncGuion(fncGuion(cboLocalCompe.BoundText, 1, "-"), 1, "|")
            cboCajas.BoundText = ""
            cboCajas.Enabled = True
            bloquea
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdAgregar_Click()
If Val(Replace(txtDia1Doc1.Text, "-", "")) > Val(Replace(txtDia1DocUlt.Text, "-", "")) Then
    MsgBox "El Número inicial no puede ser menor al número final", vbCritical, App.ProductName
    Exit Sub
End If
If Val(Replace(txtDia1Doc1.Text, "-", "")) = Val(Replace(txtDia1DocUlt.Text, "-", "")) Then
    MsgBox "El Número inicial no puede ser igual al número final", vbCritical, App.ProductName
    Exit Sub
End If

If txtDia1Doc1.Text = "" Or Len(txtDia1Doc1.Text) <> 11 Then MsgBox "Debe de ingresar el primer documento", vbCritical, App.ProductName: txtDia1Doc1.Focus: Exit Sub
If txtDia1DocUlt.Text = "" Or Len(txtDia1DocUlt.Text) <> 11 Then MsgBox "Debe de ingresar el ultimo documento", vbCritical, App.ProductName: txtDia1DocUlt.Focus: Exit Sub

If Not Valida(fncGuion(cboEmpresaCompetencia.BoundText, 0, "-"), fncGuion(cboLocalCompe.BoundText, 0, "-"), cboCajas.Text) = True Then
            MsgBox "Ya se ingresa un registro repetido" & Chr(13) & _
                        "Empresa:" & cboEmpresaCompetencia.Text & Chr(13) & _
                        "Local:" & fncGuion(cboLocalCompe.BoundText, 0, "-") & Chr(13) & _
                        "Caja:" & cboCajas.Text & Chr(13), vbCritical, App.ProductName
    Exit Sub
End If
    Dim i As Integer
    i = XDatos.Count(1)
    XDatos.AppendRows
    ''datos
        XDatos(i, 0) = fncGuion(cboEmpresaCompetencia.BoundText, 0, "-") 'Codigo del Proveedor o Competencia
        XDatos(i, 1) = fncGuion(cboEmpresaCompetencia.Text, 1, "-") 'Descripcion del Proveedor o Competencia
        XDatos(i, 2) = fncGuion(cboLocalCompe.BoundText, 0, "-") ''Codigo de local de la competencia
        XDatos(i, 3) = fncGuion(cboLocalCompe.Text, 1, "-") ''Descripcion de local de la competencia
        XDatos(i, 4) = cboCajas.Text ''caja
        XDatos(i, 5) = txtDia1Doc1.Text ''Descripcion de local de la competencia
        XDatos(i, 6) = txtDia1DocUlt.Text ''Codigo de local de la competencia
        XDatos(i, 7) = txtDiaUltDoc1.Text ''Descripcion de local de la competencia
        XDatos(i, 8) = txtDiaUltDocUlt.Text ''Codigo de local de la competencia
        XDatos(i, 9) = txtObservacion.Text ''Observacion de el registro
        XDatos(i, 10) = cboEmpresaCompetencia.BoundText ''Original valor Empresa competencia
        XDatos(i, 11) = cboLocalCompe.BoundText ''original valor de sucursal empresa
        
    grdDocumento.Array1 = XDatos
    cboEmpresaCompetencia.BoundText = ""
End Sub

Private Sub ctlDataCombo1_Change()

End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
   On Error GoTo Control

    Select Case Index
        Case 1:
            Dim strControl As String
             strControl = Graba
             If Not strControl = "" Then
                MsgBox "Se grabo el control de la competencía con le N° " & strControl, vbExclamation, App.ProductName
                Unload Me
             Else
                
             End If
        Case 2:
             Unload Me
        Case 3:
            Unload Me
        Case 4:
            Unload Me
        Case Else
            MsgBox "No se encuentra implementado", vbCritical, App.ProductName
    End Select

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number

End Sub

Private Sub dtpFecha_Change()
    BuscaDatosCaja
End Sub


Private Sub Form_Load()
   On Error GoTo Control

    FormatoGrilla
    CargaDatos
    bloquea

   Exit Sub

Control:

    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub

Sub FormatoGrilla()

    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("", "Competencia", "", "Loc. Comp.", "Caja", "Primer Doc.", "Último Doc.", "Ult. dia, Doc 1", "Ult. dia, ult. Doc", "Destino")
    arrAncho = Array(0, 2500, 0, 3400, 700, 1200, 1200, 900, 900, 900)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter)
    grdDocumento.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDocumento.Columns(0).Visible = False
    grdDocumento.Columns(2).Visible = False
    grdDocumento.Columns(7).Visible = False
    grdDocumento.Columns(8).Visible = False
    grdDocumento.Columns(9).Visible = False
    
    grdDocumento.Columns(0).AllowSizing = False
    grdDocumento.Columns(1).AllowSizing = False
    grdDocumento.Columns(2).AllowSizing = False
    grdDocumento.Columns(3).AllowSizing = False
    grdDocumento.Columns(4).AllowSizing = False
    grdDocumento.Columns(5).AllowSizing = False
    grdDocumento.Columns(6).AllowSizing = False
    grdDocumento.Columns(7).AllowSizing = False
    grdDocumento.Columns(8).AllowSizing = False
    grdDocumento.Columns(9).AllowSizing = False
    
End Sub
Sub bloquea()
    If cboCajas.BoundText = "" Then
    
        txtDia1Doc1.Text = ""
        txtDia1DocUlt.Text = ""
        txtDiaUltDoc1.Text = ""
        txtDiaUltDocUlt.Text = ""
        txtObservacion.Text = ""
    
        txtDia1Doc1.Enabled = False
        txtDia1DocUlt.Enabled = False
        txtDiaUltDoc1.Enabled = False
        txtDiaUltDocUlt.Enabled = False
        txtObservacion.Enabled = False
        cmdAgregar.Enabled = False
        'If cboLocalCompe.BoundText = "" Then lblDireccion.Caption = ""
    Else
        txtDia1Doc1.Enabled = True
        txtDia1DocUlt.Enabled = True
        txtDiaUltDoc1.Enabled = True
        txtDiaUltDocUlt.Enabled = True
        txtObservacion.Enabled = True
        cmdAgregar.Enabled = True
    End If
End Sub

Function Valida(ByVal CodigoCompetencia As String, ByVal CodigoLocalCompetencia As String, ByVal CodigoCaja As String) As Boolean
Dim o As Integer
    While o < XDatos.Count(1)
        If XDatos(o, 0) = CodigoCompetencia And XDatos(o, 2) = CodigoLocalCompetencia And XDatos(o, 4) = CodigoCaja Then
            Valida = False
            Exit Function
        End If
        o = o + 1
    Wend
    Valida = True
End Function

Private Sub grdDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDelete
''        XDatos.DeleteRows grdDocumento.Bookmark
''        grdDocumento.Array1 = XDatos
''        grdDocumento.Rebind
        grdDocumento.Delete
        grdDocumento.Array1 = XDatos
        grdDocumento.Rebind

    Case vbKeyReturn
        txtDia1Doc1.Text = XDatos(grdDocumento.Bookmark, 5)  ''Descripcion de local de la competencia
        txtDia1DocUlt.Text = XDatos(grdDocumento.Bookmark, 6)  ''Codigo de local de la competencia
        txtDiaUltDoc1.Text = XDatos(grdDocumento.Bookmark, 7)   ''Descripcion de local de la competencia
        txtDiaUltDocUlt.Text = XDatos(grdDocumento.Bookmark, 8)  ''Codigo de local de la competencia
        txtObservacion.Text = XDatos(grdDocumento.Bookmark, 9)   ''Observacion de el registro
        cboEmpresaCompetencia.BoundText = XDatos(grdDocumento.Bookmark, 10) ''Original valor Empresa competencia
        cboLocalCompe.BoundText = XDatos(grdDocumento.Bookmark, 11)  ''original valor de sucursal empresa
        cboCajas.Text = XDatos(grdDocumento.Bookmark, 4)  ''caja

End Select

End Sub

Function Graba() As String
    If txtValePromedio.Text = "" Then MsgBox "El vale promedio no puede ser Vacio", vbCritical, App.ProductName: Exit Function
    If XDatos.Count(1) <= 0 Then MsgBox "Debe de ingresar por lo menos un local de la competencia", vbCritical, App.ProductName: Exit Function
    Dim CadRucCompetencia, CadLocalCompetencia, CadCaja, CadNumDoc1, CadNumDoc2, CadNumDoc3, CadNumDoc4, CadComentario As String
    Dim f As Integer
    While f < XDatos.Count(1)
        CadRucCompetencia = CadRucCompetencia & XDatos(f, 0) & "|"
        CadLocalCompetencia = CadLocalCompetencia & XDatos(f, 2) & "|"
        CadCaja = CadCaja & XDatos(f, 4) & "|"
        CadNumDoc1 = CadNumDoc1 & XDatos(f, 5) & "|"
        CadNumDoc2 = CadNumDoc2 & XDatos(f, 6) & "|"
        CadNumDoc3 = CadNumDoc3 & XDatos(f, 7) & "|"
        CadNumDoc4 = CadNumDoc4 & XDatos(f, 8) & "|"
        CadComentario = CadComentario & XDatos(f, 9) & "|"
        f = f + 1
    Wend
    
        CadRucCompetencia = CadRucCompetencia & objUsuario.Ruc & "|"
        CadLocalCompetencia = CadLocalCompetencia & objUsuario.CodigoLocal & "|"
        CadCaja = CadCaja & objUsuario.NombrePC & "|"
        CadNumDoc1 = CadNumDoc1 & txtDia1Doc1BTL.Text & "|"
        CadNumDoc2 = CadNumDoc2 & txtDia1DocUltBTL.Text & "|"
        CadNumDoc3 = CadNumDoc3 & txtDiaUltDoc1BTL.Text & "|"
        CadNumDoc4 = CadNumDoc4 & txtDiaUltDocUltBTL.Text & "|"
        CadComentario = CadComentario & "" & "|"
        
        
    Graba = objControl.Grabar("", objUsuario.CodigoLocal, "", objUsuario.Codigo, txtValePromedio.Text, CadRucCompetencia, CadLocalCompetencia, CadCaja, CadNumDoc1, CadNumDoc2, CadNumDoc3, CadNumDoc4, CadComentario, dtpFecha.Value)
End Function

Sub BuscaDatosCaja()
    Dim rs As oraDynaset
    Set rs = objControl.DatosCaja(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, cboCaja.BoundText, dtpFecha.Value)
    txtDia1Doc1BTL.Text = "" & rs("MIN_NUMERO").Value
    txtDia1DocUltBTL.Text = "" & rs("MAX_NUMERO").Value
    txtValePromedio.Text = "" & rs("VAL_PROMEDIO").Value
End Sub
