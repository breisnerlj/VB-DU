VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_VTA_ConsultaDepositos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "frm_VTA_ConsultaDepositos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla grdDepositos 
      Height          =   5655
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9975
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFchIni 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60751873
      CurrentDate     =   39153
   End
   Begin vbp_Ventas.ctlToolBar ToolBar_AsisMoto 
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1058
   End
   Begin vbp_Ventas.ctlGrilla grdRemesa 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9975
   End
   Begin vbp_Ventas.ctlGrilla grdRemito 
      Height          =   5625
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   9922
   End
   Begin MSComCtl2.DTPicker dtpFchFin 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60751873
      CurrentDate     =   39153
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   2640
      TabIndex        =   6
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   465
   End
End
Attribute VB_Name = "frm_VTA_ConsultaDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objRemesa As New clsRemesa
Dim objRemito As New clsRemito
Dim objImpresion As New clsImpresiones
Dim strTPortaValor As String
Private pobjDepositos As New clsDepositos

Public Property Get Depositos() As clsDepositos
    Set Depositos = pobjDepositos
End Property

Public Property Set Depositos(objDepositos As clsDepositos)
    Set pobjDepositos = objDepositos
End Property

Private Sub Form_Load()

    SeteaGrillaRemesa
    SeteaGrillaRemito
    SetGrdDeposito
    
    dtpFchIni.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    dtpFchFin.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    
    ToolBar_AsisMoto.Buttons(1).Visible = False
    ToolBar_AsisMoto.Buttons(2).Visible = False
    ToolBar_AsisMoto.Buttons(10).Visible = False
    ToolBar_AsisMoto.Buttons(11).Visible = False
    
    If pobjDepositos.Control <> "2" Then
    
        If frm_VTA_Ctrl_Depositos.pblnMostrar = True Then
            grdRemesa.Visible = True
            grdRemito.Visible = False
            grdDepositos.Visible = False
          Else
            grdRemesa.Visible = False
            grdRemito.Visible = True
            grdDepositos.Visible = False
        End If
    Else
        grdRemesa.Visible = False
        grdRemito.Visible = False
    End If
    
    On Error GoTo handle
    strTPortaValor = objRemito.TipoPortavalor(objUsuario.CodigoLocal)
    Exit Sub
handle:
        MsgBox "No se encontro en el portavalor " & Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRemesa = Nothing
    Set objRemito = Nothing
    Set objImpresion = Nothing
    Set pobjDepositos = Nothing
End Sub

Private Sub ToolBar_AsisMoto_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    
   If pobjDepositos.Control <> "2" Then
        
        Select Case Index
                Case "1"
                      If frm_VTA_Ctrl_Depositos.pblnMostrar = True Then
                        Set grdRemesa.DataSource = objRemesa.Lista(objUsuario.CodigoEmpresa, _
                                                                   objUsuario.CodigoLocal, _
                                                                   CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                                   CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")))
                      Else
                        Set grdRemito.DataSource = objRemito.ListaGenerados(objUsuario.CodigoLocal, _
                                                                            CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                                            CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")))
                      End If
    
                Case "2"
                    If frm_VTA_Ctrl_Depositos.pblnMostrar = True Then
                        Set grdRemesa.DataSource = objRemesa.Lista(objUsuario.CodigoEmpresa, _
                                                                   objUsuario.CodigoLocal, _
                                                                   CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                                   CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")))
    
                      Else
                        Set grdRemito.DataSource = objRemito.ListaGenerados(objUsuario.CodigoLocal, _
                                                                            CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                                            CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")))
                    End If
    
                Case "3"
                    If grdRemito.ApproxCount <= 0 Then Exit Sub
                    On Error GoTo Pase
                    If frm_VTA_Ctrl_Depositos.pblnMostrar = True Then
                        'grdRemesa.MostrarImprimir
                      Else
                        'grdRemito.MostrarImprimir
                        If strTPortaValor = "PP" Then
                            ''25/11/2009, jescate
                             MsgBox "Se va imprimir el remito, por favor coloque el papel en la impresora", vbInformation, App.ProductName
                            objImpresion.ProsegurNew objUsuario.CodigoLocal, grdRemito.Columns("COD_REMITO").Value
                        Else
                            Dim j%
                            j = 1
                            For j = 1 To 2
                                If j = 1 Then
                                    MsgBox "Se va imprimir el remito de soles, por favor coloque el papel en la impresora", vbInformation, App.ProductName
                                    objImpresion.HermesNew objUsuario.CodigoLocal, grdRemito.Columns("COD_REMITO").Value, "S"
                                ElseIf j = 2 Then
                                    If MsgBox("Desea imprimir los dólares", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
                                    MsgBox "Se va imprimir el remito de dolares, por favor coloque el papel en la impresora", vbInformation, App.ProductName
                                    objImpresion.HermesNew objUsuario.CodigoLocal, grdRemito.Columns("COD_REMITO").Value, "D"
                                'ElseIf j > 2 Then Exit For
                                End If
                            Next j
                        End If
    
                        If MsgBox("Desea imprimir el reporte de remesas", vbYesNo + vbQuestion, App.ProductName) = vbYes Then
                                objImpresion.Imprime_Detalle_Remito objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, grdRemito.Columns("COD_REMITO").Value
                        End If
                    End If
                Exit Sub
Pase:
                MsgBox Err.Description, vbCritical, App.ProductName
                Case "4"
                    If frm_VTA_Ctrl_Depositos.pblnMostrar = True Then
                        grdRemesa.MostrarExcel
                      Else
                        grdRemito.MostrarExcel
                    End If
    
                Case "5"
                    If frm_VTA_Ctrl_Depositos.pblnMostrar = True Then
                        grdRemesa.MostrarEmail
                      Else
                        grdRemito.MostrarEmail
                    End If
    
                Case "6"
                     If frm_VTA_Ctrl_Depositos.pblnMostrar = True Then
                        If grdRemesa.ApproxCount <= 0 Then Exit Sub
                       Else
                        If grdRemito.ApproxCount <= 0 Then Exit Sub
                     End If
    
                     spAnula frm_VTA_Ctrl_Depositos.pblnMostrar
                     Set grdRemito.DataSource = objRemito.ListaGenerados(objUsuario.CodigoLocal, _
                                                                            CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                                            CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")))
    
                Case "7"
                    Unload Me
        End Select

   Else

        Select Case Index
            Case "1"
                Buscar
            Case "2"
                Buscar
            Case "3"
                If grdDepositos.ApproxCount <= 0 Then Exit Sub
                   grdDepositos.MostrarImprimir
            Case "4"
                If grdDepositos.ApproxCount <= 0 Then Exit Sub
                   grdDepositos.MostrarExcel
            Case "5"
                If grdDepositos.ApproxCount <= 0 Then Exit Sub
                   grdDepositos.MostrarEmail
            Case "6"
                Dim strError As String

                strError = pobjDepositos.Anula(objUsuario.CodigoEmpresa, _
                                               objUsuario.CodigoLocal, _
                                               Trim(grdDepositos.Columns("NRO_DEPOSITO").Value), _
                                               CStr(Format(grdDepositos.Columns("FCH_DEPOSITO").Value, "dd/mm/yyyy")))

                If strError = "" Then
                    MsgBox "Se anulo el depósito", vbInformation, App.ProductName
                    Buscar
                Else
                    MsgBox strError, vbCritical, App.ProductName
                End If
            Case "7"
                Unload Me
        End Select
  End If
End Sub

Private Sub spAnula(ByVal vstrFlgControl As Boolean)
    Dim gvarError As String
    
    Screen.MousePointer = vbHourglass
    If vstrFlgControl = True Then
    
        gvarError = objRemesa.Anula(objUsuario.CodigoEmpresa, _
                                    objUsuario.CodigoLocal, _
                                    grdRemesa.Columns("COD_MAQUINA").Value, _
                                    grdRemesa.Columns("COD_LIQUIDACION").Value, _
                                    grdRemesa.Columns("COD_REMESA").Value)
                        
            If gvarError = "" Then
                MsgBox "Se anulo la remesa", vbInformation, App.ProductName
                
                Set grdRemesa.DataSource = objRemesa.Lista(objUsuario.CodigoEmpresa, _
                                                           objUsuario.CodigoLocal, _
                                                           CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                           CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")))
                                                             
            Else
                MsgBox gvarError, vbCritical, App.ProductName
            End If
                        
    ElseIf vstrFlgControl = False Then
    
    
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
    If srtCodigoAUTH = "" Then frm_VTA_ObservaAutorizacion.Show vbModal
    strStore = ObjValidacion.Solicita("5", strAccion, strMensaje, srtCodigoAUTH, objUsuario.CodigoLocal, "", "", "", "", "", "", "1", grdRemito.Columns("COD_REMITO").Value, "", objUsuario.Codigo, frm_VTA_ObservaAutorizacion.OutObservacion, strCodigoAutorizacion, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", frm_VTA_ObservaAutorizacion.OutNumeroId)
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

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    
            gvarError = objRemito.Anula(objUsuario.CodigoLocal, _
                                        grdRemito.Columns("COD_REMITO").Value, "1")
                                    
            If gvarError = "" Then
                MsgBox "Se anulo el remito", vbInformation, App.ProductName
            Else
                MsgBox gvarError, vbCritical, App.ProductName
            End If
                                
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Buscar()

   On Error GoTo Control

        Set grdDepositos.DataSource = pobjDepositos.ListaDepositos(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, _
                                                                   CStr(Format(dtpFchIni.Value, "dd/mm/yyyy")), _
                                                                   CStr(Format(dtpFchFin.Value, "dd/mm/yyyy")))

   Exit Sub

Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub SeteaGrillaRemesa()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim i%

    arrCampos = Array("COD_REMESA", "EST_REMESA", "COD_MAQUINA", "FCH_DEPOSITO", "COD_LIQUIDACION", "ESTADO", _
                      "COD_REMITO", "FLG_ACTIVO", "COD_USUARIO_ORIGEN", "NOMB_ORIGEN", "IMP_TOTAL")
                      
    arrCaption = Array("Remesa", "Estado", "Maquina", "Fecha", "Liquidacion", "Estado", "Remito", _
                       "Activo", "Codigo", "Nombres", "Total")

    arrAncho = Array(1500, 600, 1100, 1800, 1800, 800, 1100, 900, 900, 2200, 900)
                     
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgCenter, vbAlignLeft, _
                          vbAlignLeft, dbgCenter, dbgCenter, vbAlignLeft, vbAlignLeft)
                          
    grdRemesa.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    For i = 0 To grdRemesa.Columns.Count - 1
        grdRemesa.Columns(i).Merge = True
    Next
    
    grdRemesa.Columns("FLG_ACTIVO").Visible = False

End Sub

Private Sub SeteaGrillaRemito()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_REMITO", "FCH_REMITO", "ESTADO", "FLG_ACTIVO", _
                      "NUM_PRECINTO", "IMP_TOTAL_SOLES", "IMP_TOTAL_DOLARES", "MTO_FDS")
                      
    arrCaption = Array("Remito", "Fecha", "Estado", "Activo", "Precinto", "Total S/.", "Total $", "FDS")

    arrAncho = Array(1100, 1800, 900, 900, 1500, 900, 900, 900)
                     
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
                          
    grdRemito.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

End Sub

Private Sub SetGrdDeposito()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_LOCAL", "COD_REMITO", "FCH_REMITO", "NRO_DEPOSITO", "FCH_DEPOSITO", _
                      "DES_BANCO", "MONEDA", "IMP_DEPOSITO", "DES_OBSERVACIONES")
                      
    arrCaption = Array("Local", "Remito", "Fch.Remito", "# Operacion", "Fch.Deposito", _
                       "Banco", "Moneda", " Monto Total", "Observaciones")

    arrAncho = Array(0, 1100, 1200, 1200, 1200, 2500, 1000, 1000, 2000)

    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, vbAlignLeft, vbAlignLeft, vbAlignLeft, dbgCenter, vbAlignLeft, vbAlignLeft)

    grdDepositos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    grdDepositos.Columns(0).Visible = False

    spGrilla_Traslate grdDepositos, "MONEDA", "1", "S/."
    spGrilla_Traslate grdDepositos, "MONEDA", "2", "$"
    
    grdDepositos.Columns(3).Merge = True

End Sub
