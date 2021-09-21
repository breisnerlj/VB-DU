VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private xRemesa As New XArrayDB
Private xRemito As New XArrayDB
Dim vstrLiquidacion As String
Dim objRemito As New clsRemito

Private Sub Class_Initialize()
    xRemesa.ReDim 0, -1, 0, 9  'Setea el array de los concepto de remesa
    xRemito.ReDim 0, -1, 0, 12 'Setea el array de los remitos
End Sub

Public Function ListaMtoCajaChica(ByVal CodigoCia As String, _
                                        ByVal CodigoLocal As String, _
                                        ByVal NumeroLiquidacion As String) As Double

    On Error GoTo CtrlErr
        ListaMtoCajaChica = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_RET_CAJA", CodigoCia, CodigoLocal, NumeroLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description

End Function

Public Function CierraCajaNoDoc(ByVal oradb As OraDatabase, _
                                ByVal vstrCia As String, _
                                ByVal vstrCodLocal As String, _
                                ByVal vstrCodLiquidacion As String, _
                                ByVal vstrCodUsuario As String)

    Dim i As Integer

    On Error GoTo Control
    
    For i = oradb.Parameters.Count - 1 To 0 Step -1
       oradb.Parameters.Remove i
    Next
    Dim gvarValores  As Variant
    Dim gvarIO  As Variant
    gvarValores = Array(vstrCia, _
                        vstrCodLocal, _
                        vstrCodLiquidacion, _
                        vstrCodUsuario)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    CierraCajaNoDoc = gclsOracle.SP("BTLPROD.PKG_LIQUIDACION.SP_CIERRA_CAJA_NO_DOC", _
                                                       gvarValores, _
                                                       gvarIO)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_CierraCajaNoDoc", Err.Description
End Function
                                

Public Function ListaCajasPrecerradas(ByVal vstrCia As String, _
                                        ByVal vstrCodBtl As String, _
                                        ByVal vstrFchIni As String) As oraDynaset

    On Error GoTo CtrlErr
        Set ListaCajasPrecerradas = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_LISTA_CAJA_PRECERRADA", 0, vstrCia, vstrCodBtl, vstrFchIni)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description

End Function

Public Function NomUsuario(ByVal vstrCodUsuario As String) As String
    On Error GoTo CtrlErr
        NomUsuario = gclsOracle.FN_Valor("CMR.PKG_UTIL.FN_NOM_USU", vstrCodUsuario, 1, 1)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

'**********************************************************************'
'--- Liquidación de Caja ---'
'**********************************************************************'
Public Function CalcEfectivo(ByVal vstrCia As String, _
                               ByVal vstrCodLocal As String, _
                               ByVal vstrMaquina As String, _
                               ByVal vstrCodFP As String, _
                               ByVal vstrCodFPHijo As String, _
                               ByVal vstrCodMoneda As String) As Double
    On Error GoTo CtrlErr
        CalcEfectivo = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_EFECTIVO", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrMaquina, _
                                                           vstrCodFP, _
                                                           vstrCodFPHijo, _
                                                           vstrCodMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcTarjeta(ByVal vstrCia As String, _
                              ByVal vstrCodLocal As String, _
                              ByVal vstrMaquina As String, _
                              ByVal vstrCodFP As String) As oraDynaset
                              
    On Error GoTo CtrlErr
        Set CalcTarjeta = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_CALC_TARJETA", 0, _
                                                             vstrCia, _
                                                             vstrCodLocal, _
                                                             vstrMaquina, _
                                                             vstrCodFP)
                                                             
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function
                              
Public Function CalcNotaCredito(ByVal vstrCia As String, _
                                  ByVal vstrCodLocal As String, _
                                  ByVal vstrMaquina As String, _
                                  ByVal vstrCodFP As String, _
                                  ByVal vstrCodFPHijo As String) As Double
    On Error GoTo CtrlErr
        CalcNotaCredito = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_NOTACREDITO", _
                                                             vstrCia, _
                                                             vstrCodLocal, _
                                                             vstrMaquina, _
                                                             vstrCodFP, _
                                                             vstrCodFPHijo)
                                                             
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcConvenio(ByVal vstrCia As String, _
                               ByVal vstrCodLocal As String, _
                               ByVal vstrCodMaquina As String, _
                               ByVal vstrCodLiquidacion As String) As oraDynaset
        On Error GoTo CtrlErr
            Set CalcConvenio = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_CALC_CONVENIO", 0, _
                                                               vstrCia, _
                                                               vstrCodLocal, _
                                                               vstrCodMaquina, _
                                                               vstrCodLiquidacion)
        
        Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcCheque(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrMaquina As String, _
                             ByVal vstrCodFP As String, _
                             ByVal vstrCodFPHijo As String, _
                             ByVal vstrCodMoneda As String) As Double

    On Error GoTo CtrlErr
        CalcCheque = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_CHEQUE", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrMaquina, _
                                                           vstrCodFP, _
                                                           vstrCodFPHijo, _
                                                           vstrCodMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcDocDscto(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrMaquina As String, _
                             ByVal vstrCodFP As String, _
                             ByVal vstrCodFPHijo As String) As Double
    On Error GoTo CtrlErr
          CalcDocDscto = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_DOC_DCTO", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrMaquina, _
                                                           vstrCodFP, _
                                                           vstrCodFPHijo)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalDonacion(ByVal vstrCia As String, _
                            ByVal vstrCodLocal As String, _
                            ByVal vstrCodMaquina As String, _
                            ByVal vstrCodFP As String) As oraDynaset
    On Error GoTo CtrlErr
    Set CalDonacion = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_CALC_DONACION", 0, _
                                                    vstrCia, _
                                                    vstrCodLocal, _
                                                    vstrCodMaquina, _
                                                    vstrCodFP)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function fnCalFormaPagos(ByVal vstrCia As String, _
                                ByVal vstrCodBtl As String, _
                                ByVal vstrCodMaquinas As String, _
                                ByVal vstrCodMoneda As String, _
                                ByVal vstrCodLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr
    Set fnCalFormaPagos = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_CALC_FORMAS_PAGOS", 0, vstrCia, vstrCodBtl, vstrCodMaquinas, vstrCodMoneda, vstrCodLiquidacion)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function


'------------------------------'
'---    REMESAS Y REMITOS ---  '
'------------------------------'

Public Function ListaUsurioxArqueo(ByVal vstrCia As String, _
                                   ByVal vstrCodLocal As String, _
                                   ByVal vstrFchIniDep As String, _
                                   ByVal vstrFchFinDep As String, _
                                   Optional ByVal vstrFlgCondicion As String = "0") As oraDynaset
    On Error GoTo CtrlErr
    Set ListaUsurioxArqueo = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_USU_X_ARQUEAR", 0, _
                                                  vstrCia, _
                                                  vstrCodLocal, _
                                                  vstrFchIniDep, _
                                                  vstrFchFinDep, _
                                                  vstrFlgCondicion)
            
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function ListaConcepto() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaConcepto = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_LISTA_CONCEPTO", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

'-- Propiedades --'

Public Property Get Remesa() As XArrayDB
    Set Remesa = xRemesa 'Devuelve la variable de la clase en forma de XArray
End Property

Public Property Get Remito() As XArrayDB
    Set Remito = xRemito 'Devuelve la variable de la clase en forma de XArray
End Property

'***********************************************'
'- Metodo Para Cargar la consulta de Conceptos -'
'***********************************************'
Public Sub CargaConceptoRemesa()
    Dim odynRemesa  As oraDynaset
    Dim strConcepto$, strDesConcepto$, strCta1$, strCta2$, strCta3$, strTipoCta$, strFP_P$, strFP_H$
    
    On Error GoTo CtrlErr
    
    'Llamada a la Consulta'
    Set odynRemesa = ListaConcepto
    
    odynRemesa.MoveFirst
    While Not odynRemesa.EOF
            
        strConcepto = odynRemesa("COD_CONCEPTO").Value
        strDesConcepto = odynRemesa("DES_CONCEPTO").Value
        strCta1 = odynRemesa("NUM_CTA_CTE_01").Value
        strCta2 = odynRemesa("NUM_CTA_CTE_02").Value
        strCta3 = odynRemesa("NUM_CTA_CTE_03").Value
        strTipoCta = odynRemesa("FLG_TIPO_CTA").Value
        strFP_P = odynRemesa("COD_FORMA_PAGO").Value
        strFP_H = odynRemesa("COD_HIJO").Value
            
        AgregaConceptoRemesa strConcepto, strDesConcepto, strCta1, strCta2, strCta3, strTipoCta, strFP_P, strFP_H
        odynRemesa.MoveNext
    Wend
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Sub


'*************************************************************************************
' Metodo para agregar un concepto al arreglo xRemesa
'*************************************************************************************
Public Function AgregaConceptoRemesa(ByVal vstrCodigo$, ByVal vstrDescripcion$, _
                                     ByVal vstrCta1$, ByVal vstrCta2$, _
                                     ByVal vstrCta3$, ByVal vstrTipoCta$, _
                                     ByVal vstrFP_P$, ByVal vstrFP_H$) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xRemesa.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xRemesa.Count(1)
    While i < aux
        If xRemesa(i, 0) = vstrCodigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xRemesa.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xRemesa.AppendRows
    End If
j:
    If xRemesa.Count(1) = 0 Then ultimo = 0: xRemesa.AppendRows
        xRemesa(ultimo, 0) = vstrCodigo          ' Asigna el codigo de concepto
        xRemesa(ultimo, 1) = vstrDescripcion     ' Asigna la descripcion del concepto
        xRemesa(ultimo, 2) = vstrCta1            ' Asigna si la cuenta 1
        xRemesa(ultimo, 3) = vstrCta2            ' Asigna si la cuenta 2
        xRemesa(ultimo, 4) = vstrCta3            ' Asigna si la cuenta 3
        xRemesa(ultimo, 5) = "0"                 ' Importe soles
        xRemesa(ultimo, 6) = "0"                 ' Importe Dolares
        xRemesa(ultimo, 7) = vstrTipoCta
        xRemesa(ultimo, 8) = vstrFP_P            ' Forma de pago de la remesa padre
        xRemesa(ultimo, 9) = vstrFP_H            ' Forma de pago de la remesa hijo
    Set AgregaConceptoRemesa = xRemesa
    
End Function


'****************************************************'
'- Metodo Para Cargar la remesas por asignar remito -'
'****************************************************'
Public Sub CargaListaRemesa(ByVal vstrFchIni As String, _
                            ByVal vstrFchFin As String)

 Dim odynRemito As oraDynaset
 Dim strRemesa As String, strLiquidacion As String
 Dim strMoneda As String, strSobre As String
 Dim strCodUsuario As String, strNomUsu As String, StrImp As String
 Dim strConcepto As String, strEstRemesa As String
 Dim strSmbMon As String, strFchIniCaja As String

    On Error GoTo CtrlErr
    Set odynRemito = objRemito.ListaRemesas(objUsuario.CodigoEmpresa, _
                                            objUsuario.CodigoLocal, _
                                            "1", _
                                            "0", _
                                            vstrFchIni, _
                                            vstrFchFin)
                                            
    odynRemito.MoveFirst
    While Not odynRemito.EOF
       strRemesa = odynRemito("COD_REMESA").Value
       strLiquidacion = "" & odynRemito("COD_LIQUIDACION").Value
       strMoneda = odynRemito("COD_MONEDA").Value
       strSobre = odynRemito("NRO_SOBRE").Value
       strCodUsuario = odynRemito("COD_USUARIO_ORIGEN").Value
       strNomUsu = odynRemito("NOM_USUARIO_ORIGEN").Value
       StrImp = odynRemito("IMP_SOBRE").Value
       strConcepto = odynRemito("COD_CONCEPTO").Value
       strEstRemesa = odynRemito("EST_REMESA").Value
       strSmbMon = odynRemito("SMB_MONEDA").Value
       strFchIniCaja = odynRemito("FCH_INICIO").Value
       
       AgregaListaRemesa strRemesa, strLiquidacion, _
                         strSmbMon, strSobre, _
                         strCodUsuario, strNomUsu, _
                         StrImp, strConcepto, _
                         strEstRemesa, strMoneda, _
                         strFchIniCaja
       odynRemito.MoveNext
    Wend
    frm_VTA_Remitos.grdRemito.Rebind
    ' -- Footer --'
    Dim odynClon As oraDynaset
    Dim dblValSol#, dblValDol#
    Dim k%
    
    Set odynClon = odynRemito.Clone
    dblValSol = 0: dblValDol = 0
    
    odynClon.MoveFirst
    While Not odynClon.EOF
       If odynClon("COD_MONEDA").Value = "1" Then
            dblValSol = dblValSol + odynClon("IMP_SOBRE").Value
       ElseIf odynClon("COD_MONEDA").Value = "2" Then
            dblValDol = dblValDol + odynClon("IMP_SOBRE").Value
       End If
        odynClon.MoveNext
    Wend
    
    frm_VTA_Remitos.grdRemito.Columns(5).FooterText = "S/." & " " & Format(dblValSol, "###,###0.00")
    frm_VTA_Remitos.grdRemito.Columns(9).FooterText = "$." & " " & Format(dblValDol, "###,###0.00")
    
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
    
End Sub

'*************************************************************************************
' Metodo para agregar un concepto al arreglo xRemesa
'*************************************************************************************
Public Function AgregaListaRemesa(ByVal vstrRemesa$, ByVal vstrLiquid$, _
                                  ByVal vstrMon$, ByVal vstrSobre$, _
                                  ByVal vstrCodUsuario$, _
                                  ByVal vstrUSUARIO$, ByVal vstrImp$, _
                                  ByVal vstrConcepto$, ByVal vstrEstado$, _
                                  ByVal vstrSmbMon$, ByVal vstrFchIniCaja$) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xRemito.Count(1) < 0 Then Exit Function
    
'    Dim i As Integer
'    Dim encontro As Boolean


'    aux = xRemito.Count(1)
'    While i < aux
'        If xRemito(i, 10) = Trim(vstrRemesa & vstrMon) Then
'            ultimo = i
'            encontro = True
'GoTo j
'        Else
'            encontro = False
'            ultimo = xRemito.Count(1) 'Llena con la ultima posicion disponible
'        End If
'        i = i + 1
'    Wend
'    If encontro = False Then
'        xRemito.AppendRows
'    End If
'j:
    
    'If xRemesa.Count(1) = 0 Then 'ultimo = 0: xRemito.AppendRows
        xRemito.AppendRows
        xRemito(xRemito.UpperBound(1), 0) = vstrRemesa
        xRemito(xRemito.UpperBound(1), 1) = vstrLiquid
        xRemito(xRemito.UpperBound(1), 2) = vstrMon
        xRemito(xRemito.UpperBound(1), 3) = vstrSobre
        xRemito(xRemito.UpperBound(1), 4) = vstrUSUARIO
        xRemito(xRemito.UpperBound(1), 5) = vstrImp
        xRemito(xRemito.UpperBound(1), 6) = vstrConcepto
        xRemito(xRemito.UpperBound(1), 7) = vstrEstado
        xRemito(xRemito.UpperBound(1), 8) = vstrSmbMon
        xRemito(xRemito.UpperBound(1), 9) = Empty
        xRemito(xRemito.UpperBound(1), 10) = vstrCodUsuario
        xRemito(xRemito.UpperBound(1), 11) = vstrFchIniCaja
    Set AgregaListaRemesa = xRemito
  'End If
End Function

'--- Saca el total declarado del local por las remesas ---'
Public Function fnTotDeclLocal(ByVal vstrCia As String, _
                               ByVal vstrCodBtl As String, _
                               ByVal vstrCodMaquina As String, _
                               ByVal vstrCodLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr
    Set fnTotDeclLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_TOT_DEC_LOCAL", 0, vstrCia, vstrCodBtl, vstrCodMaquina, vstrCodLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

'-- Lista los arqueos cerrados --'
Public Function fndevLiqCerrados(ByVal vstrCia As String, _
                                 ByVal vstrCodBtl As String, _
                                 ByVal vstrFchIni As String) As oraDynaset
    On Error GoTo CtrlErr
    Set fndevLiqCerrados = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_DEV_ARQ_CERRADOS", 0, vstrCia, vstrCodBtl, vstrFchIni)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

'-- Muestra Nombre Usuario --'
Public Function fnDevNomUsu(ByVal vstrCodUsuario As String) As String
    On Error GoTo CtrlErr
    fnDevNomUsu = gclsOracle.FN_Valor("CMR.PKG_UTIL.FN_NOM_USU", vstrCodUsuario)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

'-- Graba Liquidacion de caja --'
Public Function Graba(ByVal oradb As OraDatabase, _
                      ByVal vstrCia As String, _
                      ByVal vstrCodBtl As String, _
                      ByVal vstrCodMaquina As String, _
                      ByVal vstrCodDepend As String, _
                      ByVal vstrCodUsuario As String, _
                      ByVal vstrCadImpLiq As String, _
                      Optional vstrObs As String, _
                      Optional vstrCadCodFPago As String, _
                      Optional vstrCadCodFPagoLiq As String, _
                      Optional vstrCadRetEfect As String, _
                      Optional vstrRetCaja As String, _
                      Optional vstrDevCodLiq As String, _
                      Optional vstrCodLiquidacion As String, _
                      Optional vstrCodRemesaP As String, _
                      Optional vstrCodRemesaH As String, _
                      Optional vstrCadImpRemesa As String, _
                      Optional vstrCadImpRemesaSoles As String, _
                      Optional vstrCadCodDocumento As String, _
                      Optional vstrCadMinDocumento As String, _
                      Optional vstrCadMaxDocumento As String) As String
                      
    Dim i As Integer
    
    On Error GoTo Control
    
    For i = oradb.Parameters.Count - 1 To 0 Step -1
       oradb.Parameters.Remove i
    Next
    Dim gvarValores  As Variant
    Dim gvarIO  As Variant
    gvarValores = Array(vstrCia, _
                        vstrCodBtl, _
                        vstrCodMaquina, _
                        vstrCodDepend, _
                        vstrCodUsuario, _
                        vstrCadImpLiq, _
                        vstrObs, _
                        vstrCadCodFPago, _
                        vstrCadCodFPagoLiq, _
                        vstrCadRetEfect, _
                        vstrRetCaja, _
                        vstrDevCodLiq, _
                        vstrCodLiquidacion, _
                        vstrCodRemesaP, _
                        vstrCodRemesaH, _
                        vstrCadImpRemesa, _
                        vstrCadImpRemesaSoles, _
                        vstrCadCodDocumento, _
                        vstrCadMinDocumento, _
                        vstrCadMaxDocumento)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   Salida, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    Graba = gclsOracle.SP("BTLPROD.PKG_LIQUIDACION.SP_GRABA", _
                              gvarValores, _
                              gvarIO)
    
    vstrLiquidacion = gvarValores(13)
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_Graba", Err.Description
                      
End Function
                      
'Public Property Let CodLiquidacion(vstrLiquidacion As String)
'    CodLiquidacion = vstrLiquidacion
'End Property
                      
Public Property Get CodLiquidacion() As String
    CodLiquidacion = vstrLiquidacion
End Property
                      
Public Function Anula(ByVal oradb As OraDatabase, _
                      ByVal vstrCia As String, _
                      ByVal vstrCodBtl As String, _
                      ByVal vstrCodLiquidacion As String, Optional ByVal FlagCredito As String = "0", Optional ByVal CodigoUsuario As String = "") As String

    Dim i As Integer

    On Error GoTo Control
    
    For i = oradb.Parameters.Count - 1 To 0 Step -1
       oradb.Parameters.Remove i
    Next
    Dim gvarValores  As Variant
    Dim gvarIO  As Variant
    gvarValores = Array(vstrCia, _
                        vstrCodBtl, _
                        vstrCodLiquidacion, FlagCredito, CodigoUsuario)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    Anula = gclsOracle.SP("BTLPROD.PKG_LIQUIDACION.SP_ANULA", _
                                                       gvarValores, _
                                                       gvarIO)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_Anula", Err.Description

End Function




'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'
'--- Impresion de la Liquidacion de Caja ---'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'

'** Hecho el 21/12/2006 Crueda

Public Function ImpLiquidacionCabecera(ByVal vstrCia As String, _
                                       ByVal vstrCodLocal As String, _
                                       ByVal vstrCodMaquina As String, _
                                       ByVal vstrCodLiquidacion As String) As oraDynaset
    On Error GoTo Control
    Set ImpLiquidacionCabecera = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.LISTACAB_LIQ", 0, vstrCia, vstrCodLocal, vstrCodMaquina, vstrCodLiquidacion)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_ImpLiquidacionCabecera", Err.Description
End Function

Public Function ImpLiquidacionDetalle(ByVal vstrCodLiquidacion As String) As oraDynaset
    On Error GoTo Control
    Set ImpLiquidacionDetalle = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.LISTADET_LIQ", 0, vstrCodLiquidacion)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_ImpLiquidacionDetalle", Err.Description
End Function

Public Function ImpLiquidacionRemesa(ByVal vstrCia As String, _
                                     ByVal vstrCodLocal As String, _
                                     ByVal vstrCodMaquina As String, _
                                     ByVal vstrCodLiquidacion As String) As oraDynaset
    On Error GoTo Control
    Set ImpLiquidacionRemesa = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.LISTA_REMESA_LIQ", 0, vstrCia, vstrCodLocal, vstrCodMaquina, vstrCodLiquidacion)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_ImpLiquidacionRemesa", Err.Description
End Function

'Public Function TipoMaquina(ByVal vstrCia As String, _
'                            ByVal vstrCodLocal As String, _
'                            ByVal vstrCodMquina As String) As String
'On Error GoTo CrltErr
'    Set TipoMaquina = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION._TIPO_MAQUINA", vstrCia, _
'                                                                                 vstrCodLocal, _
'                                                                                 vstrCodMquina)
'    Exit Function
'CrltErr:
'    Err.Raise Err.Number, "clsLiquidacion_TipoMaquina", Err.Description
'End Function

Public Function DifCajaImp(ByVal vstrCodLiquidacion As String) As String
    On Error GoTo Control
       DifCajaImp = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.DIF_CAJA_IMP", vstrCodLiquidacion)
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_DifCajaImp", Err.Description
End Function

'Public Function FnExisteDoc(ByVal vstrCia As String, _
'                            ByVal vstrCodLocal As String, _
'                            ByVal vstrCodLiquidacion As String) As String
'    On Error GoTo Control
'       FnExisteDoc = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_VERIF_DOC_CAJA", vstrCia, vstrCodLocal, vstrCodLiquidacion)
'
'    Exit Function
'Control:
'    Err.Raise Err.Number, "clsLiquidacion_FnExisteDoc", Err.Description
'
'End Function

Public Function Det_Liquidacion_Venta(ByVal vstrCia As String, _
                                      ByVal vstrCodLocal As String, _
                                      ByVal vstrCodLiquidacion As String, _
                                      ByVal vstrIndicador As String) As oraDynaset

    On Error GoTo Control
      Set Det_Liquidacion_Venta = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.DET_LIQUIDACION_VENTA", 0, vstrCia, _
                                                                                                         vstrCodLocal, _
                                                                                                         vstrCodLiquidacion, _
                                                                                                         vstrIndicador)
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_Det_Liquidacion_Venta", Err.Description
End Function

Public Function Lista_Sec_x_Doc(ByVal vstrCia As String, _
                                ByVal vstrCodLocal As String, _
                                ByVal vstrCodLiquidacion As String) As oraDynaset

    On Error GoTo Control
      Set Lista_Sec_x_Doc = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_LISTA_SEC_X_DOC", 0, vstrCia, _
                                                                                                  vstrCodLocal, _
                                                                                                  vstrCodLiquidacion)
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_Lista_Sec_x_Doc", Err.Description
End Function

Public Function SecxDoc()
    On Error GoTo Control
        SecxDoc = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_VAR_SEC_X_DOC")
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_SecxDoc", Err.Description
End Function

Public Function VtaxModalidad()
    On Error GoTo Control
        VtaxModalidad = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_VAR_VTA_X_MOD")
    Exit Function
Control:
    Err.Raise Err.Number, "clsLiquidacion_VtaxModalidad", Err.Description
End Function

Public Function CalcServicio(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrCodLiquidacion As String, _
                             ByVal vstrCodServicio As String) As Double
    On Error GoTo CtrlErr
          CalcServicio = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_SERVICIOS", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion, _
                                                           vstrCodServicio)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcCajero(ByVal vstrCia As String, _
                           ByVal vstrCodLocal As String, _
                           ByVal vstrCodLiquidacion As String, _
                           ByVal vstrCargoAbono As String) As Double
    On Error GoTo CtrlErr
          CalcCajero = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_CAJERO", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion, _
                                                           vstrCargoAbono)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcSoat(ByVal vstrCia As String, _
                         ByVal vstrCodLocal As String, _
                         ByVal vstrCodLiquidacion As String) As Double
    On Error GoTo CtrlErr
          CalcSoat = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_SOAT", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcVisaCash(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrCodLiquidacion As String, _
                             ByVal vstrCargoAbono As String) As Double
    On Error GoTo CtrlErr
          CalcVisaCash = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_VISACASH", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion, _
                                                           vstrCargoAbono)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcDonacion(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrCodLiquidacion As String) As Double
    On Error GoTo CtrlErr
          CalcDonacion = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_DONACIONES", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcEfect(ByVal vstrCia As String, _
                          ByVal vstrCodLocal As String, _
                          ByVal vstrCodLiquidacion As String, _
                          ByVal vstrCodHijo As String) As Double
    On Error GoTo CtrlErr
          CalcEfect = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_EFECT", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion, _
                                                           vstrCodHijo)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcCupones(ByVal vstrCia As String, _
                            ByVal vstrCodLocal As String, _
                            ByVal vstrCodLiquidacion As String) As Double
    On Error GoTo CtrlErr
          CalcCupones = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_CUPONES", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function CalcCheques(ByVal vstrCia As String, _
                            ByVal vstrCodLocal As String, _
                            ByVal vstrCodLiquidacion As String) As Double
    On Error GoTo CtrlErr
          CalcCheques = gclsOracle.FN_Valor("BTLPROD.PKG_LIQUIDACION.FN_CALC_W_CHEQUES", _
                                                           vstrCia, _
                                                           vstrCodLocal, _
                                                           vstrCodLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Function

Public Function ListaDocsEmitidos(ByVal vstrCia As String, _
                                  ByVal vstrCodLiquidacion As String, _
                                  Optional ByVal vstrCodEstado As String) As oraDynaset

    On Error GoTo CtrlErr
        Set ListaDocsEmitidos = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_DOC_EMITIDOS", 0, vstrCia, vstrCodLiquidacion, vstrCodEstado)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description

End Function

Public Function ListaTarjetas(ByVal vstrCia As String, _
                              ByVal vstrCodLocal As String, _
                              ByVal vstrCodLiquidacion As String) As oraDynaset

    On Error GoTo CtrlErr
        Set ListaTarjetas = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_CALC_TARJETA", 0, vstrCia, vstrCodLocal, vstrCodLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description

End Function


Public Function ListaDetLiqFormaPago(ByVal vstrCia As String, _
                                     ByVal vstrLocal As String, _
                                     ByVal vstrLiquidacion As String, _
                                     ByVal vstrCodFP As String, _
                                     Optional ByVal vstrHijo As String = "") As oraDynaset
    On Error GoTo CtrlErr
    Set ListaDetLiqFormaPago = gclsOracle.FN_Cursor("BTLPROD.PKG_LIQUIDACION.FN_DET_LIQ_FORMA_PAGO", 0, vstrCia, _
                                                                                                        vstrLocal, _
                                                                                                        vstrLiquidacion, _
                                                                                                        vstrCodFP) ',
                                                                                                        'vstrHijo)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.ListaDetLiqFormaPago", Err.Description
End Function

'#################################################################
Public Function RepoListaDatos(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr
    
    Set RepoListaDatos = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_DATOS", 0, vstrCodigoLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaDatos", Err.Description
End Function

Public Function RepoListaXTipDoc(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr
    
    Set RepoListaXTipDoc = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_X_TIPDOC", 0, vstrCodigoLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaXTipDoc", Err.Description
End Function

Public Function RepoListaXForPag(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr
    
    Set RepoListaXForPag = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_X_FORPAG", 0, vstrCodigoLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaXForPag", Err.Description
End Function

Public Function RepoListaXModVta(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr
    
    Set RepoListaXModVta = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_X_MODALIDAD", 0, vstrCodigoLiquidacion)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaXModVta", Err.Description
End Function

Public Function RepoListaDocTarjeta(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr

    Set RepoListaDocTarjeta = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_DOC_TARJETA", 0, vstrCodigoLiquidacion)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaDocTarjeta", Err.Description
End Function

Public Function RepoListaDocCajeExpress(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr

    Set RepoListaDocCajeExpress = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_DOC_CAJEXPRESS", 0, vstrCodigoLiquidacion)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaDocCajeExpress", Err.Description
End Function

Public Function RepoListaRemesas(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr

    Set RepoListaRemesas = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_REMESAS", 0, vstrCodigoLiquidacion)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaRemesas", Err.Description
End Function

Public Function RepoListaNCMalgen(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr

    Set RepoListaNCMalgen = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_NC_MALGEN", 0, vstrCodigoLiquidacion)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaNCMalgen", Err.Description
End Function

Public Function RepoListaErrores(ByVal vstrCodigoLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr

    Set RepoListaErrores = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_LIQUIDACION.FN_LISTA_ERROR", 0, vstrCodigoLiquidacion)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLiquidación.RepoListaErrores", Err.Description
End Function

