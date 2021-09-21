VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_VTA_Remitos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "frm_VTA_Remitos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraRemito 
      BackColor       =   &H80000004&
      Caption         =   "Remito"
      ForeColor       =   &H00C00000&
      Height          =   6675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9020
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton CmdDolares 
         Caption         =   "&Imprime Dolares"
         Height          =   465
         Left            =   5243
         TabIndex        =   24
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton CmdSoles 
         Caption         =   "&Imprime Soles"
         Height          =   465
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5880
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFchRemito 
         Height          =   375
         Left            =   7200
         TabIndex        =   18
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64552961
         CurrentDate     =   39021
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   465
         Left            =   7800
         TabIndex        =   16
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton CmdVistaRemito 
         Caption         =   "&Vista Preliminar"
         Height          =   465
         Left            =   4006
         TabIndex        =   3
         Top             =   5880
         Width           =   975
      End
      Begin VB.CheckBox ChkTodo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Todos"
         Height          =   255
         Left            =   4320
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmOkRemito 
         Height          =   300
         Left            =   5280
         Picture         =   "frm_VTA_Remitos.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   380
      End
      Begin vbp_Ventas.ctlGrillaArray grdRemito 
         Height          =   3015
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5318
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpFchIniRemesa 
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64552961
         CurrentDate     =   39016
      End
      Begin MSComCtl2.DTPicker dtpFchFinRemesa 
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64552961
         CurrentDate     =   39016
      End
      Begin vbp_Ventas.ctlTextBox TxtFds 
         Height          =   330
         Left            =   7200
         TabIndex        =   20
         Top             =   1245
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ColorDefault    =   -2147483633
         ColorDefault    =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   6
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
      Begin vbp_Ventas.ctlTextBox txtPrecinto 
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         Top             =   960
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
      Begin vbp_Ventas.ctlTextBox txtPrecintoDol 
         Height          =   330
         Left            =   1680
         TabIndex        =   26
         Top             =   1282
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
      Begin vbp_Ventas.ctlTextBox txtRemitoPI 
         Height          =   330
         Left            =   1680
         TabIndex        =   27
         Top             =   1605
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         Tipo            =   8
         MaxLength       =   30
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
      Begin VB.Label lblRemitoPI 
         BackStyle       =   0  'Transparent
         Caption         =   "Remito Pre-Impreso"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1640
         Width           =   1575
      End
      Begin VB.Label lblPreDolares 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Precinto Dolares"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Portavalor"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblPreSoles 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Precinto"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   998
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   255
         Left            =   6000
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " F3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   8160
         TabIndex        =   17
         Top             =   6360
         Width           =   285
      End
      Begin VB.Label LblTotalSobRemesa 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBFBFA&
         Caption         =   "Cantidad Sobres =>"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label LblCtdRemesa 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00DBFBFA&
         Caption         =   "Mandar a PortaValor =>"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label LblTotDolaresRemito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   7560
         TabIndex        =   12
         Top             =   5160
         Width           =   1125
      End
      Begin VB.Label LblTotSolesRemito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   6360
         TabIndex        =   11
         Top             =   5160
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fondo Sencillo"
         Height          =   195
         Left            =   6000
         TabIndex        =   10
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label LblBancoRemito 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFBFA&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_VTA_Remitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objRemito As New clsRemito
Dim objPrinter As New clsImpresiones

Public pdblSolx As Double
Public pdblSolx_Aux As Double
Public pdblDolx As Double
Public pdblDolx_Aux As Double
Public pstrTPortaValor As String
Public pstrPortaValor As String
Public pdblFDS As Double
Public pstrFecha As String
Public pstrMon As String
Public pstrCtdSob As String
Public pstrNumRecinto As String
Public pblnSalir As Boolean

Dim vCont As Double
Dim dblTotSobRem As Double
Dim dblContR As Double

Dim strCodRemAnt As String
Dim strMsgRemito As String

Dim strMensaje As String



Private Sub Form_Load()
On Error GoTo handle
    dtpFchIniRemesa.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    dtpFchFinRemesa.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    dtpFchRemito.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    SeteaGrilla
    
    LblBancoRemito.Caption = objRemito.BancoxBtl(objUsuario.CodigoLocal)
    pstrTPortaValor = objRemito.TipoPortavalor(objUsuario.CodigoLocal)
    pstrPortaValor = objRemito.Portavalor(objUsuario.CodigoLocal)
    
    dblTotSobRem = 0: dblContR = 0
    TxtFds.Text = "0"
    If pstrTPortaValor = "PP" Then
         CmdDolares.Visible = False
         CmdSoles.Visible = False
         CmdVistaRemito.Visible = True
         ''''
         txtPrecintoDol.Text = ""
         txtPrecintoDol.Visible = False
         lblPreDolares.Visible = False
         lblPreSoles.Caption = "Nº Precinto"
      Else
         CmdDolares.Visible = True
         CmdSoles.Visible = True
         CmdVistaRemito.Visible = False
         
         txtPrecintoDol.Text = ""
         txtPrecintoDol.Visible = True
         lblPreDolares.Visible = True
         lblPreSoles.Caption = "Nº Precinto Soles"
    End If
    
    If pstrPortaValor <> "BOTICA TORRES DE LIMATAMBO S.A.C." Then
       txtRemitoPI.Text = ""
       txtRemitoPI.Visible = True
       lblRemitoPI.Visible = True
    Else
        txtRemitoPI.Text = ""
        txtRemitoPI.Visible = False
        lblRemitoPI.Visible = False
    End If
    '' para que se carge automaticamente
    CmOkRemito_Click
    ChkTodo.Value = "1"
    ChkTodo_Click
    EstadoPantalla "NUEVO"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub CmOkRemito_Click()
On Error GoTo handle
    objLiquidacion.Remito.ReDim 0, -1, 0, 11
    grdRemito.Array1 = objLiquidacion.Remito
    ChkTodo.Value = 0
    ChkTodo_Click
    
    objLiquidacion.CargaListaRemesa Format(dtpFchIniRemesa.Value, "dd/mm/yyyy"), _
                                    Format(dtpFchFinRemesa.Value, "dd/mm/yyyy")
                                    
    grdRemito.Array1 = objLiquidacion.Remito
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Calcula_Marcado()
On Error GoTo handle
    Dim j%
    pdblSolx = 0: pdblDolx = 0: vCont = 0
    pdblSolx_Aux = 0: pdblDolx_Aux = 0
    
    grdRemito.MoveNext
    grdRemito.MovePrevious
    For j = 0 To objLiquidacion.Remito.UpperBound(1)
       If objLiquidacion.Remito(j, 9) = "-1" Then
         If objLiquidacion.Remito(j, 8) = "1" Then
            pdblSolx = pdblSolx + objLiquidacion.Remito(j, 5)
            pdblSolx_Aux = pdblSolx_Aux + objLiquidacion.Remito(j, 5)
            vCont = vCont + 1
            
         ElseIf objLiquidacion.Remito(j, 8) = "2" Then
            pdblDolx = pdblDolx + objLiquidacion.Remito(j, 5)
            pdblDolx_Aux = pdblDolx_Aux + objLiquidacion.Remito(j, 5)
            vCont = vCont + 1
            
         End If
       End If
    Next j
    LblCtdRemesa.Caption = vCont
    LblTotSolesRemito.Caption = "S/." & " " & Format(pdblSolx, "###,###0.00")
    LblTotDolaresRemito.Caption = "$" & " " & Format(pdblDolx, "###,###0.00")

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub ChkTodo_Click()
On Error GoTo handle
Dim k%
    k = 0
    'If objLiquidacion.Remito.UpperBound(1) = -1 Then MsgBox "No hay remesas generadas", vbExclamation, Caption: ChkTodo.Value = 0: Exit Sub
    'If grdRemito.ApproxCount <= 0 Then MsgBox "No hay remesas generadas", vbExclamation, Caption: ChkTodo.Value = 0: Exit Sub
    If ChkTodo.Value = "1" Then
         For k = 0 To objLiquidacion.Remito.UpperBound(1)
              objLiquidacion.Remito(k, 9) = "-1"
         Next k
         Call Calcula_Marcado
         
         LblTotalSobRemesa.Caption = "Cantidad Sobres =>" & "  " & objRemito.TotalSobres(objUsuario.CodigoEmpresa, _
                                                                                         objUsuario.CodigoLocal, _
                                                                                         CStr(Format(dtpFchIniRemesa.Value, "dd/mm/yyyy")), _
                                                                                         CStr(Format(dtpFchFinRemesa.Value, "dd/mm/yyyy")))
      Else
         For k = 0 To objLiquidacion.Remito.UpperBound(1)
             objLiquidacion.Remito(k, 9) = "0"
         Next k
         Call Calcula_Marcado
         
         LblTotalSobRemesa.Caption = "Cantidad Sobres =>"
    End If
    grdRemito.Rebind

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Sub Graba(ByRef NumeroRemito As String)
Dim gvarError As String
Dim strRetCodRemito As String
Dim arrRems As Variant
'On Error GoTo handle
   
   gvarError = objRemito.Graba(gclsOracle.ODataBase, _
                               pdblSolx, _
                               pdblDolx, _
                               Mid(Trim(LblBancoRemito.Caption), 1, 2), _
                               Trim(txtPrecinto.Text) & IIf(Trim(txtPrecintoDol.Text) = "", "", "," & Trim(txtPrecintoDol.Text)), _
                               CStr(Format(dtpFchRemito.Value, "dd/mm/yyyy")), _
                               pdblFDS, strRetCodRemito, Trim(txtRemitoPI.Text))
                                  
   strMensaje = gvarError
   If gvarError = "" Then
        strCodRemAnt = strRetCodRemito
        NumeroRemito = strRetCodRemito
        If pstrPortaValor = "BOTICA TORRES DE LIMATAMBO S.A.C." Then
           arrRems = Split(strRetCodRemito, ",")
           MsgBox "Se grabo satisfactoriamente el o los Remitos N°" & vbCrLf & Join(arrRems, vbCrLf), vbExclamation, App.ProductName
        Else
            MsgBox "Se Grabo el remito sastifactoriamente el remito " & strRetCodRemito, vbInformation, Caption
        End If
        'objLiquidacion.Remito.ReDim 0, -1, 0, 10
        'grdRemito.Array1 = objLiquidacion.Remito
     Else
        strCodRemAnt = ""
        NumeroRemito = ""
        MsgBox gvarError, vbCritical, Caption
   End If

'    Exit Sub
'handle:
'    Err.Raise Err.Number, "Graba", Err.Description
                    
End Sub

Private Sub CmdVistaRemito_Click()
On Error GoTo handle
   TxtFds.Enabled = True
   If TxtFds.Text = "" Then MsgBox "Ingrese el Fondo de Sencillo", vbCritical, App.ProductName: TxtFds.SetFocus: Exit Sub
   If txtPrecinto.Text = "" Then MsgBox "Ingrese el Nº Precinto", vbCritical, App.ProductName: txtPrecinto.SetFocus: Exit Sub
    
   If vCont <= 0 Then MsgBox "No hay marcado ningun monto a enviar a portavalor", vbExclamation, App.ProductName: Exit Sub
    pstrFecha = Format(dtpFchRemito.Value, "dd/mm/yyyy")
    pstrCtdSob = LblTotalSobRemesa.Caption
    pstrNumRecinto = txtPrecinto.Text
    
    If pstrTPortaValor = "PP" Then
        'PROSEGUR'
        pdblFDS = TxtFds.Text
        pdblSolx_Aux = pdblSolx + pdblFDS
        
        If pdblSolx_Aux > 0 Then
           pstrMon = "S"
        End If
        
        If pdblDolx_Aux > 0 Then
           pstrMon = "D"
        End If
        
        frm_VTA_Preview_Remito_PP.Show vbModal
        
     End If
    '********************************************************************'
    '********************************************************************'
    ' aca grabara'
     If frm_VTA_Remitos.pblnSalir = False Then
     
       'strMsgRemito = objRemito.ExisteRemito(strCodRemAnt)
        'Autor      : Arturo Escate
        'Fecha      : 25/11/2009
        'Proposito  : esto se comento porque el formato jalaba valores de la pantalla y no los grabados
        '==============================================================================================
        Dim strNumeroRemito As String
        Graba strNumeroRemito
        If Not strNumeroRemito = "" Then
           objPrinter.ProsegurNew objUsuario.CodigoLocal, strNumeroRemito
        End If
        '==============================================================================================
        If strMensaje = "" Then
                                     
'Impresion Soles y Dolares'
'Autor      : Arturo Escate
'Fecha      : 25/11/2009
'Proposito  : esto se comento porque el formato jalaba valores de la pantalla y no los grabados
'                objPrinter.Impresion_Prosegur frm_VTA_Preview_Remito_PP.pCliente, _
'                                              frm_VTA_Preview_Remito_PP.pRecibido, _
'                                              frm_VTA_Preview_Remito_PP.pPara, _
'                                              frm_VTA_Preview_Remito_PP.pLugar, _
'                                              frm_VTA_Preview_Remito_PP.pDesSol, _
'                                              frm_VTA_Preview_Remito_PP.pDesDol, _
'                                              frm_VTA_Preview_Remito_PP.pMtoSol, _
'                                              frm_VTA_Preview_Remito_PP.pMtoDol, _
'                                              frm_VTA_Preview_Remito_PP.pDirecc, _
'                                              frm_VTA_Preview_Remito_PP.pLocalidad, _
'                                              frm_VTA_Preview_Remito_PP.pPrecinto, _
'                                              frm_VTA_Preview_Remito_PP.pCtdSobres
'
                'Modificacion para el Detallado de Remesa que estan en el Remito
                'Modificado por Carlos Cieza bajo requeriemto
                
               If MsgBox("Se va a imprimir el detalle del remito - Colocar papel", _
                               vbQuestion + vbYesNo + vbDefaultButton1, "Confirme") = vbYes Then
                               
                        objPrinter.Imprime_Detalle_Remito objUsuario.CodigoEmpresa, _
                                                          objUsuario.CodigoLocal, _
                                                          strCodRemAnt
                                                          '"0000010178" PRUEBA DE REMITO EN DESARROLLO
                                                          

               Else
                   Exit Sub
               End If

                'Aca termina la modificacion del requerimiento y continua el resto del proceso

                NuevoPP
            Else
                Exit Sub
        End If
     End If
     
Exit Sub
handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName

End Sub

Private Sub CmdDolares_Click()
On Error GoTo handle
   TxtFds.Enabled = False
   If txtPrecinto.Text = "" Then MsgBox "Ingrese el Nº Precinto", vbCritical, App.ProductName: txtPrecinto.SetFocus: Exit Sub
    
   If vCont <= 0 Then MsgBox "No hay marcado ningun monto a enviar a portavalor", vbExclamation, App.ProductName: Exit Sub
    pstrFecha = Format(dtpFchRemito.Value, "dd/mm/yyyy")
    pstrCtdSob = LblTotalSobRemesa.Caption
    pstrNumRecinto = txtPrecinto.Text

    If pstrTPortaValor = "PH" Then
        'HERMES DOLARES'
        pdblFDS = TxtFds.Text
        pdblSolx_Aux = pdblSolx + pdblFDS
        If pdblDolx_Aux > 0 Then
           pstrMon = "D"
           frm_VTA_Preview_Remito_PH.Show vbModal
        Else
            MsgBox "No Hay Remesas Dolares o ya fue Impreso", vbExclamation, App.ProductName: grdRemito.Delete: txtPrecinto.Text = "": Exit Sub
        End If
    '********************************************************************'
    '********************************************************************'
    ' aca grabara'
    
     If frm_VTA_Remitos.pblnSalir = False Then
     
        strMsgRemito = objRemito.ExisteRemito(strCodRemAnt)
        If strMsgRemito = "0" Then
            'JESCATE_NEW
            Dim strNumeroRemito As String
            Graba strNumeroRemito
            objPrinter.HermesNew objUsuario.CodigoLocal, strNumeroRemito, "D"
        End If
     
        If strMensaje = "" Then
                'Imprime Dolares'
''''                objPrinter.Impresion_Hermes frm_VTA_Preview_Remito_PH.pLugar, _
''''                                            frm_VTA_Preview_Remito_PH.pFecha, _
''''                                            frm_VTA_Preview_Remito_PH.pEmpresa, _
''''                                            frm_VTA_Preview_Remito_PH.pDesMonto, _
''''                                            frm_VTA_Preview_Remito_PH.pBanco, _
''''                                            frm_VTA_Preview_Remito_PH.pCtaCte, _
''''                                            frm_VTA_Preview_Remito_PH.pMonto, _
''''                                            frm_VTA_Preview_Remito_PH.pPrecinto, _
''''                                            frm_VTA_Preview_Remito_PH.pDirEntrega, _
''''                                            frm_VTA_Preview_Remito_PH.pEntrega, _
''''                                            frm_VTA_Preview_Remito_PH.pEnvase
                NuevoPH "2"
           Else
               Exit Sub
        End If
     End If
   End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
   
End Sub

Private Sub CmdSoles_Click()
On Error GoTo handle
   TxtFds.Enabled = True
   If TxtFds.Text = "" Then MsgBox "Ingrese el Fondo de Sencillo", vbCritical, App.ProductName: TxtFds.SetFocus: Exit Sub
   If txtPrecinto.Text = "" Then MsgBox "Ingrese el Nº Precinto", vbCritical, App.ProductName: txtPrecinto.SetFocus: Exit Sub
    
   If vCont <= 0 Then MsgBox "No hay marcado ningun monto a enviar a portavalor", vbExclamation, App.ProductName: Exit Sub
    pstrFecha = Format(dtpFchRemito.Value, "dd/mm/yyyy")
    pstrCtdSob = LblTotalSobRemesa.Caption
    pstrNumRecinto = txtPrecinto.Text
    
    If pstrTPortaValor = "PH" Then
        'HERMES SOLES'
        pdblFDS = TxtFds.Text
        pdblSolx_Aux = pdblSolx + pdblFDS
        If pdblSolx_Aux > 0 Then
           pstrMon = "S"
           frm_VTA_Preview_Remito_PH.Show vbModal
        Else
           MsgBox "No Hay Remesas Soles o ya fue Impreso", vbExclamation, App.ProductName: grdRemito.Delete: txtPrecinto.Text = "": Exit Sub
        End If
    
    '********************************************************************'
    '********************************************************************'
    ' aca grabara'
    
     If frm_VTA_Remitos.pblnSalir = False Then
     
        strMsgRemito = objRemito.ExisteRemito(strCodRemAnt)
        If strMsgRemito = "0" Then
            'JESCATE_NEW
             Dim strNumeroRemito As String
             Graba strNumeroRemito
             If Not strNumeroRemito = "" Then
             objPrinter.HermesNew objUsuario.CodigoLocal, strNumeroRemito, "S"
              End If
        End If
        
        
        If strMensaje = "" Then
                'Imprime Soles'
                
                
''''                objPrinter.Impresion_Hermes frm_VTA_Preview_Remito_PH.pLugar, _
''''                                            frm_VTA_Preview_Remito_PH.pFecha, _
''''                                            frm_VTA_Preview_Remito_PH.pEmpresa, _
''''                                            frm_VTA_Preview_Remito_PH.pDesMonto, _
''''                                            frm_VTA_Preview_Remito_PH.pBanco, _
''''                                            frm_VTA_Preview_Remito_PH.pCtaCte, _
''''                                            frm_VTA_Preview_Remito_PH.pMonto, _
''''                                            frm_VTA_Preview_Remito_PH.pPrecinto, _
''''                                            frm_VTA_Preview_Remito_PH.pDirEntrega, _
''''                                            frm_VTA_Preview_Remito_PH.pEntrega, _
''''                                            frm_VTA_Preview_Remito_PH.pEnvase
                NuevoPH "1"
           Else
                Exit Sub
         End If
     End If
   End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
   
End Sub

Sub NuevoPP()
On Error GoTo handle
    vCont = 0
    pdblSolx = 0: pdblSolx_Aux = 0
    pdblDolx = 0: pdblDolx_Aux = 0
    TxtFds.Text = 0
    grdRemito.Delete
    grdRemito.Rebind
    ChkTodo.Value = 0
    txtPrecinto.Text = "": LblBancoRemito.Caption = ""
    LblTotSolesRemito.Caption = "": LblTotDolaresRemito.Caption = ""
    LblCtdRemesa.Caption = "": LblTotalSobRemesa.Caption = "Cantidad Sobres =>"
    frm_VTA_Remitos.grdRemito.Columns(5).FooterText = ""
    frm_VTA_Remitos.grdRemito.Columns(9).FooterText = ""
    pstrTPortaValor = "PP"

    Set objRemito = Nothing
    Set objRemito = New clsRemito
    objLiquidacion.Remito.ReDim 0, -1, 0, 11
    grdRemito.Array1 = objLiquidacion.Remito
    grdRemito.Rebind

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Sub NuevoPH(ByVal vstrMon As String)
On Error GoTo handle
     If vstrMon = "1" Then
        TxtFds.Text = 0
        pdblSolx = 0: pdblSolx_Aux = 0
'        txtPrecinto.Text = ""
        frm_VTA_Remitos.grdRemito.Columns(5).FooterText = ""
     Else
        TxtFds.Text = 0
        pdblDolx = 0: pdblDolx_Aux = 0
'        txtPrecinto.Text = ""
        frm_VTA_Remitos.grdRemito.Columns(9).FooterText = ""
     End If
     pstrTPortaValor = "PH"
            
     'grdRemito.Delete
        
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strMsgRemito = "0"
End Sub

Private Sub grdRemito_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo handle
     If ColIndex = 9 Then
        Call Calcula_Marcado
        
        If objLiquidacion.Remito(grdRemito.row, 9) Then
            dblTotSobRem = dblTotSobRem + objRemito.TotalSobresRemesa(objUsuario.CodigoEmpresa, _
                                                                      objUsuario.CodigoLocal, _
                                                                      objUsuario.NombrePC, _
                                                                      Format(dtpFchIniRemesa.Value, "dd/mm/yyyy"), _
                                                                      Format(dtpFchFinRemesa.Value, "dd/mm/yyyy"), _
                                                                      objLiquidacion.Remito(grdRemito.row, 8), _
                                                                      objLiquidacion.Remito(grdRemito.row, 0))

            LblTotalSobRemesa.Caption = "Cantidad Sobres =>" & "  " & dblTotSobRem

          Else
            dblTotSobRem = 0
            LblTotalSobRemesa.Caption = "Cantidad Sobres =>" & "  " & dblTotSobRem
        End If
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Label6_Click()

End Sub

Private Sub TxtFds_Change()
On Error GoTo handle
    If Not IsNumeric(TxtFds.Text) Then Exit Sub
    pdblFDS = TxtFds.Text

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub TxtFds_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    TxtFds.Tipo = Real
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub txtPrecinto_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    txtPrecinto.Tipo = AlfaNumerico

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Unload Me
    End Select
End Sub

Sub SeteaGrilla()
On Error GoTo handle
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim arrFoco As Variant

    arrCampos = Array("", "", _
                      "", "", _
                      "", "", _
                      "", "", _
                      "", "", _
                      "", "")
    arrCaption = Array("Remesa", "Cod.Liq", _
                       "Smb", "Sobre Nº", _
                       "Cajero", "Total", _
                       "Concepto", "Estado", _
                       "Moneda", "Chk", _
                       "CodDepen", "FchIni Caja")
    arrAncho = Array(1200, 1200, _
                     500, 1100, _
                     2200, 1200, _
                     900, 800, _
                     800, 1000, _
                     900, 1200)
    arrAlineacion = Array(dbgCenter, dbgCenter, _
                          dbgLeft, dbgCenter, _
                          dbgLeft, dbgRight, _
                          dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter)
    arrFoco = Array(False, False, _
                    False, False, _
                    False, False, _
                    False, False, _
                    False, True, _
                    False, False)
    grdRemito.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    
    grdRemito.ColumnFooter = True
    grdRemito.Columns(6).Visible = False
    grdRemito.Columns(7).Visible = False
    grdRemito.Columns(8).Visible = False
    grdRemito.Columns(9).ValueItems.Presentation = dbgCheckBox
    grdRemito.Columns(5).FooterFont.Size = 8: grdRemito.Columns(9).FooterFont.Size = 8
    grdRemito.Columns(5).FooterFont.Bold = False: grdRemito.Columns(9).FooterFont.Bold = False

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdSalir_Click()
    strMsgRemito = "0"
    Unload Me
End Sub

''''
Private Sub cmdGrabar_Click()
Dim strNumeroRemito As String
Dim i As Integer
Dim arrCad As Variant
On Error GoTo handle
   Select Case pstrTPortaValor
   Case "PH"
            If TxtFds.Text = "" Then MsgBox "Ingrese el Fondo de Sencillo", vbCritical, App.ProductName: TxtFds.SetFocus: Exit Sub
            If txtPrecinto.Text = "" Then MsgBox "Ingrese el Nº Precinto Soles", vbCritical, App.ProductName: txtPrecinto.SetFocus: Exit Sub
            If txtPrecintoDol.Text = "" Then MsgBox "Ingrese el Nº Precinto Dolares", vbCritical, App.ProductName: txtPrecintoDol.SetFocus: Exit Sub
            If grdRemito.ApproxCount <= 0 Then MsgBox "No tiene ninguna remesa para generar remito", vbCritical, App.ProductName: Exit Sub
            If pstrPortaValor <> "BOTICA TORRES DE LIMATAMBO S.A.C." Then
                If txtRemitoPI.Text = "" Then MsgBox "Ingrese el Nº Remito PreImpreso", vbCritical, App.ProductName: txtRemitoPI.SetFocus: Exit Sub
            End If
            Graba strNumeroRemito
            
            If Not strNumeroRemito = "" Then
                 If pstrPortaValor = "BOTICA TORRES DE LIMATAMBO S.A.C." Then
                    arrCad = Split(strNumeroRemito, ",")
                    For i = LBound(arrCad) To UBound(arrCad)
                        MsgBox "Por favor colocar el papel para imprimir el Remito en soles", vbInformation, App.ProductName
                        'objPrinter.HermesNew objUsuario.CodigoLocal, strNumeroRemito, "S"
                        objPrinter.HermesNew objUsuario.CodigoLocal, arrCad(i), "S"
                        MsgBox "Por favor colocar el papel para imprimir el Remito en dolares", vbInformation, App.ProductName
                        'objPrinter.HermesNew objUsuario.CodigoLocal, strNumeroRemito, "D"
                        objPrinter.HermesNew objUsuario.CodigoLocal, arrCad(i), "D"
                        
                        If MsgBox("Se va a imprimir el detalle del remito - Colocar papel", vbQuestion + vbYesNo + vbDefaultButton1, "Confirme") = vbYes Then
                           objPrinter.Imprime_Detalle_Remito objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, arrCad(i)
                        End If
                    Next
                 Else
                     MsgBox "Por favor colocar el papel para imprimir el Remito en soles", vbInformation, App.ProductName
                     objPrinter.HermesNew objUsuario.CodigoLocal, strNumeroRemito, "S"
                     MsgBox "Por favor colocar el papel para imprimir el Remito en dolares", vbInformation, App.ProductName
                     objPrinter.HermesNew objUsuario.CodigoLocal, strNumeroRemito, "D"
                     
                     If MsgBox("Se va a imprimir el detalle del remito - Colocar papel", vbQuestion + vbYesNo + vbDefaultButton1, "Confirme") = vbYes Then
                        objPrinter.Imprime_Detalle_Remito objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strNumeroRemito
                     End If
                 End If
            End If
            
    Case "PP"
            If TxtFds.Text = "" Then MsgBox "Ingrese el Fondo de Sencillo", vbCritical, App.ProductName: TxtFds.SetFocus: Exit Sub
            If txtPrecinto.Text = "" Then MsgBox "Ingrese el Nº Precinto Soles", vbCritical, App.ProductName: txtPrecinto.SetFocus: Exit Sub
            If grdRemito.ApproxCount <= 0 Then MsgBox "No tiene ninguna remesa para generar remito", vbCritical, App.ProductName: Exit Sub
            If pstrPortaValor <> "BOTICA TORRES DE LIMATAMBO S.A.C." Then
                If txtRemitoPI.Text = "" Then MsgBox "Ingrese el Nº Remito PreImpreso", vbCritical, App.ProductName: txtRemitoPI.SetFocus: Exit Sub
            End If
            Graba strNumeroRemito
            
            If Not strNumeroRemito = "" Then
               If pstrPortaValor = "BOTICA TORRES DE LIMATAMBO S.A.C." Then
                  arrCad = Split(strNumeroRemito, ",")
                  MsgBox "Por favor colocar el papel para imprimir el o los Remitos", vbInformation, App.ProductName
                  For i = LBound(arrCad) To UBound(arrCad)
                      'objPrinter.ProsegurNew objUsuario.CodigoLocal, strNumeroRemito
                      objPrinter.ProsegurNew objUsuario.CodigoLocal, arrCad(i)
                   
                      If MsgBox("Se va a imprimir el detalle del remito - Colocar papel", vbQuestion + vbYesNo + vbDefaultButton1, "Confirme") = vbYes Then
                         objPrinter.Imprime_Detalle_Remito objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, arrCad(i)
                      End If
                  Next
               Else
                   MsgBox "Por favor colocar el papel para imprimir el Remito", vbInformation, App.ProductName
                   objPrinter.ProsegurNew objUsuario.CodigoLocal, strNumeroRemito
                   If MsgBox("Se va a imprimir el detalle del remito - Colocar papel", vbQuestion + vbYesNo + vbDefaultButton1, "Confirme") = vbYes Then
                      objPrinter.Imprime_Detalle_Remito objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strNumeroRemito
                   End If
               End If
            End If
            
    Case Else
        MsgBox "No se ha definido el portavalor", vbCritical, App.ProductName
        Exit Sub
    End Select

    cmdGrabar.Enabled = False
    EstadoPantalla "BLOQUEADO"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub
Private Sub EstadoPantalla(ByVal Estado As String)
    Select Case Estado
    Case "NUEVO"
        txtPrecinto.Locked = False
        txtPrecintoDol.Locked = False
        dtpFchRemito.Enabled = True
        TxtFds.bloqueado = False
        txtRemitoPI.Locked = False

        txtPrecinto.Text = ""
        txtPrecintoDol.Text = ""
        dtpFchRemito.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
        TxtFds.Text = "0"
        txtRemitoPI.Text = ""
    Case "BLOQUEADO"
        txtPrecinto.Locked = True
        txtPrecintoDol.Locked = True
        txtRemitoPI.Locked = True
        dtpFchRemito.Enabled = False
        TxtFds.bloqueado = True
    Case Else
        MsgBox "Funcionalidad no implementada", vbCritical, App.ProductName
    End Select
    CmdSoles.Visible = False
    CmdVistaRemito.Visible = False
    CmdDolares.Visible = False
End Sub

Private Sub txtPrecintoDol_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    txtPrecintoDol.Tipo = AlfaNumerico

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
