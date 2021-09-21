VERSION 5.00
Begin VB.Form frm_VTA_DatoAdicional 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4683
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtValor 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      Tipo            =   2
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
   End
   Begin VB.Label lblCodigo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "90412"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Debe ingresar la siguiente información requerida por el convenio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   3
      Top             =   540
      Width           =   5865
   End
   Begin VB.Label Label1 
      Caption         =   "Dato Adicional"
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
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frm_VTA_DatoAdicional.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H8000000B&
      Caption         =   "BONO CALVIAL D COM CJA X 30 COM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1110
      Width           =   3015
   End
End
Attribute VB_Name = "frm_VTA_DatoAdicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intValMin As Integer
Public output_valor As String
Dim objConvenio As New clsConvenio
Dim MuestraLista As Boolean

Public Sub subDatos(strConvenio As String, strCodigo As String, strDescrip As String, intTipo As Integer, intMax As Integer, intMin As Integer, Valor$)
    output_valor = ""
    lblCodigo.Caption = strCodigo
    lblDescripcion.Caption = strDescrip
    txtValor.Tipo = intTipo
    txtValor.MaxLength = intMax
    txtValor.Text = Valor
    intValMin = intMin
    Me.Caption = strDescrip 'strConvenio
    MuestraLista = IIf(objConvenio.TieneListaDatoAdicional(strCodigo) = 1, True, False)
    If MuestraLista = True Then
        Me.Height = 4155
        ctlGrilla1.Visible = True
       Set ctlGrilla1.DataSource = objConvenio.ListaValoresDatosAdicionales(strCodigo)
       txtValor.Visible = False
       lblDescripcion.Visible = False
       txtValor.TabStop = False
        ctlGrilla1.TabIndex = 0
    Else
        Me.Height = 1980
         ctlGrilla1.Visible = False
          txtValor.Visible = True
       lblDescripcion.Visible = True
       txtValor.TabIndex = 0
       txtValor.TabStop = True
    End If
    
    'Me.Show vbModal
End Sub

Private Sub ctlGrilla1_DblClick()
    Selecciona
End Sub

Private Sub ctlGrilla1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Selecciona
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Selecciona
        Case vbKeyEscape
                Unload Me
    End Select
End Sub

Private Sub Form_Load()
Formato_Grilla
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objConvenio = Nothing
End Sub

Private Sub Formato_Grilla()
    On Error GoTo handle
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant

        arrCampos = Array("COD_EUX_TIPO_CAMPO", "DES_EUX_TIPO_CAMPO")
        arrCaption = Array("Código", "Descripción")
        arrAncho = Array(1000, 5000)
        arrAlineacion = Array(dbgLeft, dbgLeft)
       ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
      
    
    Exit Sub
handle:
        MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Sub Selecciona()
    If MuestraLista = True Then
        output_valor = "" & ctlGrilla1.Columns(0).Value
    Else
            If intValMin <= Len(txtValor.Text) Then
                output_valor = Trim(txtValor.Text)
                Unload Me
                'objVenta.AgregaLineaDatoAdicional lblCodigo.Caption, lblDescripcion.Caption, txtValor.Tipo, txtValor.MaxLength, intValMin, txtValor.Text
            Else
                MsgBox "El valor ingresado debe ser minimo " & intValMin & " digitos", vbCritical, App.ProductName
                Exit Sub
            End If
    End If
    Unload Me

End Sub

