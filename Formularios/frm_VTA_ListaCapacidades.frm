VERSION 5.00
Begin VB.Form frm_VTA_ListaCapacidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Capacidades"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrillaArray tdbgListaCapacidades 
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5530
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   1400
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   4680
      Width           =   1400
   End
   Begin VB.Label lblMensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3960
      Width           =   6735
   End
   Begin VB.Label lblCapacidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Capacidad: "
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
      Left            =   5520
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblHorario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Horario Elegido: "
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
      Left            =   840
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frm_VTA_ListaCapacidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim data As Dictionary
Dim xListaArray As New XArrayDB
Public strLocal As String
Dim strSegmento As String
Dim strTipo As String
Dim objWS As New clsWebService
Public strFechaEsc As String
Public strHoraEsc As String
Private strCodLocalPriv As String
Private strFechaPriv As String
Private strHoraPriv As String
Private strTipoPriv As String

Function Datos(ByVal CodLocal As String, ByVal segmento As String, ByVal Tipo As String)
    strLocal = CodLocal: strSegmento = segmento: strTipo = Tipo
    Set data = objWS.listaCapacidades(strLocal, strSegmento, strTipo)
    'Set data = obj
End Function

Private Sub cmdValidar_Click()
    Dim lcResp As Dictionary
    If Len(Trim(strFechaPriv)) = 0 Then Exit Sub
    Set lcResp = objWS.validaCapacidades(strCodLocalPriv, strFechaPriv, strHoraPriv, strTipoPriv)
    If Not lcResp Is Nothing Then
        If IsObject(lcResp("data")) = False Then
            Exit Sub
        End If
        'For x = 1 To lcResp("data").Count()
            lblHorario = strFechaPriv & " " & strHoraPriv 'lcResp("data")(x)("day") & lcResp("data")(x)("schedules")(0)("value")
            lblCapacidad = "NO"
            If lcResp("data")("code") = "1" Then
                lblCapacidad = IIf(lcResp("data")("schedules")(1)("hasCapacity") = True, "SI", "NO")
                lblMensaje = lcResp("data")("schedules")(1)("message")
            End If
        'Next x
    End If
    If lblCapacidad = "SI" Then
        strFechaEsc = strFechaPriv
        strHoraEsc = strHoraPriv
        cmdAceptar.Enabled = True
    Else
        strFechaEsc = ""
        strHoraEsc = ""
        cmdAceptar.Enabled = False
    End If
End Sub

Private Sub cmdAceptar_Click()
    If Len(Trim(strFechaEsc)) <= 0 Or Len(Trim(strHoraEsc)) <= 0 Then MsgBox "Debe escoger un horario.", vbOKOnly, App.ProductName: Exit Sub
    If MsgBox("Se reemplazara las fecha/hora pactada." & Chr(13) & "Desea continuar ?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then Exit Sub
    frm_VTA_Documento.DTPicker1.Value = CDate(Format(strFechaEsc, "dd/mm/yyyy"))
    frm_VTA_Documento.DTPicker3.Value = CDate(Format(strHoraEsc, "hh:mm:ss AMPM"))
    'MsgBox "Se asigno la hora/fecha escogida", vbOKOnly, App.ProductName
    objVenta.flgDatosCapacidad = True
    frm_VTA_Documento.cmdAceptar.Enabled = True
    objVenta.bk_codLocalCapacidad = strLocal
    objVenta.bk_FechaCapacidad = CDate(Format(strFechaEsc, "dd/mm/yyyy"))
    objVenta.bk_HoraCapacidad = CDate(Format(strHoraEsc, "hh:mm:ss AMPM"))
    objVenta.bk_ServiceType = strTipo
    Unload Me
End Sub

Function setFormatGrid()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Fecha", "Inicio", "Fin", "Tiempo DLV", "Horario", "Mensaje", "Capacidad", "Hoy", "Capacidad2", "Valor", "ValorFinal")
    arrAncho = Array(2000, 1000, 1000, 700, 2000, 0, 0, 0, 0, 0, 0, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter)
    
    tdbgListaCapacidades.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Function

Private Sub Form_Load()
    Dim x, j, i As Integer
    Dim n, m As Integer
    
    setFormatGrid
    cmdValidar.Enabled = False
    cmdAceptar.Enabled = False
    
    xListaArray.ReDim 0, -1, 0, 11
    tdbgListaCapacidades.Array1 = xListaArray
    tdbgListaCapacidades.Rebind
    tdbgListaCapacidades.Refresh
    
    tdbgListaCapacidades.Columns("Mensaje").Visible = False
    tdbgListaCapacidades.Columns("Capacidad").Visible = False
    tdbgListaCapacidades.Columns("Hoy").Visible = False
    tdbgListaCapacidades.Columns("Capacidad2").Visible = False
    tdbgListaCapacidades.Columns("Valor").Visible = False
    tdbgListaCapacidades.Columns("ValorFinal").Visible = False
    If Not data Is Nothing Then
        If IsObject(data("data")) = False Then
            Exit Sub
        End If
        tdbgListaCapacidades.Array1 = xListaArray
        'se tenia como idea principal mostrar una fila y dentro de una celda un combo con los horarios
        'esto da error de tipo de datos .Add item
        'por lo que se mostrara 1 fila por cada horario repitiendo la fecha
        Dim Item As New TrueDBGrid70.ValueItem
        n = 0
        For x = 1 To data("data").Count()
            Debug.Print data("data")(x)("schedules").Count()
            For i = 1 To data("data")(x)("schedules").Count()
                xListaArray.AppendRows
                xListaArray(n, 0) = data("data")(x)("day")
                xListaArray(n, 1) = data("data")(x)("startHour")
                xListaArray(n, 2) = data("data")(x)("endHour")
                xListaArray(n, 3) = data("data")(x)("deliveryTime")
                xListaArray(n, 4) = data("data")(x)("schedules")(i)("time")
                xListaArray(n, 5) = data("data")(x)("message")
                xListaArray(n, 6) = data("data")(x)("hasCapacity")
                xListaArray(n, 7) = data("data")(x)("today")
                xListaArray(n, 8) = data("data")(x)("schedules")(i)("hasCapacity")
                xListaArray(n, 9) = data("data")(x)("schedules")(i)("value")
                xListaArray(n, 10) = data("data")(x)("schedules")(i)("valueEnd")
                n = n + 1
            Next i
'            xListaArray.AppendRows
'            xListaArray(x - 1, 0) = data("data")(x)("day")
'            xListaArray(x - 1, 1) = data("data")(x)("startHour")
'            xListaArray(x - 1, 2) = data("data")(x)("endHour")
'            xListaArray(x - 1, 3) = data("data")(x)("deliveryTime")
'
'            With tdbgListaCapacidades.Columns(4).ValueItems
'                .Presentation = dbgComboBox
'                .Translate = True
'                For i = 1 To data("data")(x)("schedules").Count()
'                    item.Value = "001" 'data("data")(x)("schedules")(i)("value")
'                    item.DisplayValue = "CASA" 'data("data")(x)("schedules")(i)("time")
'
'                    .Add item
'                Next i
'            End With
        Next x
    End If
    
    tdbgListaCapacidades.Array1 = xListaArray
    tdbgListaCapacidades.Rebind
    tdbgListaCapacidades.Refresh
End Sub

'Private Sub tdbgListaCapacidades_DblClick()
'    Dim Fecha As String
'    Dim hora As String
'    Dim Tipo As String
'    Dim CodLocal As String
'    Dim lcResp As Dictionary
'    Dim x, j, i As Integer
'    Dim n, m As Integer
'    strFechaEsc = ""
'    strHoraEsc = ""
'    If tdbgListaCapacidades.ApproxCount = 0 Then Exit Sub
'    Fecha = tdbgListaCapacidades.Columns("Fecha").Value
'    hora = tdbgListaCapacidades.Columns("Valor").Value
'    Tipo = "EXP"
'    CodLocal = strLocal
'    Set lcResp = objWS.validaCapacidades(CodLocal, Fecha, hora, Tipo)
'    If Not lcResp Is Nothing Then
'        If IsObject(lcResp("data")) = False Then
'            Exit Sub
'        End If
'        'For x = 1 To lcResp("data").Count()
'            lblHorario = Fecha & " " & hora 'lcResp("data")(x)("day") & lcResp("data")(x)("schedules")(0)("value")
'            lblCapacidad = "NO"
'            If lcResp("data")("code") = "1" Then
'                lblCapacidad = IIf(lcResp("data")("schedules")(1)("hasCapacity") = True, "SI", "NO")
'                lblMensaje = lcResp("data")("schedules")(1)("message")
'            End If
'        'Next x
'    End If
'    strFechaEsc = Fecha
'    strHoraEsc = hora
'End Sub

Private Sub tdbgListaCapacidades_RegistroSeleccionado(ByVal DatoColumna0 As String)
    strFechaPriv = ""
    strHoraPriv = ""
    strTipoPriv = ""
    strCodLocalPriv = ""
    If tdbgListaCapacidades.ApproxCount = 0 Then Exit Sub
    strFechaPriv = tdbgListaCapacidades.Columns("Fecha").Value
    strHoraPriv = tdbgListaCapacidades.Columns("Valor").Value
    strTipoPriv = strTipo
    strCodLocalPriv = strLocal
    objVenta.bk_ServiceType = strTipoPriv
    If Len(Trim(strFechaPriv)) > 0 Then
        cmdValidar.Enabled = True
    Else
        cmdValidar.Enabled = False
    End If
End Sub
