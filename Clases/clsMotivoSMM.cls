VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMotivoSMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim objMaximo As New clsMaximoManual
Dim strCodigo As String
Dim strDescripcion As String
Dim intCtdDiasVigencia As Integer
Dim intFlgActivo As Integer

Public Property Get Codigo() As String
Codigo = strCodigo
End Property

Public Property Let Codigo(ByVal vNewValue As String)
    strCodigo = vNewValue
    
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

Public Property Let Descripcion(ByVal vNewValue As String)
    strDescripcion = vNewValue
End Property

Public Property Get CtdDiasVigencia() As Integer
    CtdDiasVigencia = intCtdDiasVigencia
End Property

Public Property Let CtdDiasVigencia(ByVal vNewValue As Integer)
    intCtdDiasVigencia = vNewValue
End Property

Public Property Get FlgActivo() As Integer
    FlgActivo = intFlgActivo
End Property

Public Property Let FlgActivo(ByVal vNewValue As Integer)
    intFlgActivo = vNewValue
End Property


Public Property Get Maximo() As clsMaximoManual
    Set Maximo = objMaximo
End Property

Public Property Set Maximo(ByVal vNewValue As clsMaximoManual)
    Set Maximo = vNewValue
End Property

Function Lista() As oraDynaset

    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTIVO_SMM.FN_LISTA", 0)

End Function

Function ListaActivos() As oraDynaset

    Set ListaActivos = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTIVO_SMM.FN_LISTA", 0, "", "1")

End Function


Sub Registrar()

Dim varrValores As Variant
Dim varrIO As Variant
Dim strError As String

varrValores = Array(Descripcion, Maximo.Codigo, CtdDiasVigencia, FlgActivo)
varrIO = Array(entrada, entrada, entrada, entrada)

strError = gclsOracle.SP("BTLPROD.PKG_MOTIVO_SMM.SP_GRABAR_NUEVO_SMM", varrValores, varrIO)

If strError <> "" Then
    Err.Raise -1, "", strError
End If

End Sub


Sub Modificar()

Dim varrValores As Variant
Dim varrIO As Variant
Dim strError As String

varrValores = Array(Maximo.Codigo, CtdDiasVigencia, FlgActivo)
varrIO = Array(entrada, entrada, entrada)

strError = gclsOracle.SP("BTLPROD.PKG_MOTIVO_SMM.SP_MODIFICA_MOTIVO_SMM", varrValores, varrIO)

If strError <> "" Then
    Err.Raise -1, "", strError
End If

End Sub

Sub myNew(pstrCodigo As String)
Dim datos As oraDynaset

    Set datos = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTIVO_SMM.FN_LISTA", 0, pstrCodigo)
    Codigo = "" & datos("COD_MOTIVO_SMM").Value
    Descripcion = "" & datos("DES_MOTIVO_SMM").Value
    CtdDiasVigencia = "" & datos("CTD_DIAS_VIGENCIA").Value
    Maximo.Codigo = "" & datos("COD_MAXIMO").Value
    FlgActivo = "" & datos("FLG_ACTIVO").Value
  
End Sub

