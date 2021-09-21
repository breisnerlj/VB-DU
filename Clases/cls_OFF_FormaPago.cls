VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_OFF_FormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xFormaPago As New XArrayDB
Private objIni As New cls_ArchivoIni

Private intCodigo As Integer
Private strDescripcion As Integer



Public Property Get Codigo() As Integer
    Codigo = intCodigo
End Property
Public Property Let Codigo(ByVal newValue As Integer)
    intCodigo = newValue
End Property
Public Property Get Descripcion() As Integer
    Descripcion = strDescripcion
End Property
Public Property Let Descripcion(ByVal newValue As Integer)
    strDescripcion = newValue
End Property


Private Sub BuscaDatosFormaPago()
    
    Dim arrCodFormaPago As Variant
    Dim arrDesFormaPago As Variant

    arrCodFormaPago = Split(objIni.LeerIni(gstrIni, "general", "COD_FORMAPAGO", ""), ",")
    arrDesFormaPago = Split(objIni.LeerIni(gstrIni, "general", "DES_FORMAPAGO", ""), ",")
    
    Set xFormaPago = ArrayaXarray(arrCodFormaPago, arrDesFormaPago)

End Sub

Private Sub Class_Initialize()
    
    xFormaPago.ReDim 0, -1, 0, 1
    BuscaDatosFormaPago
End Sub


Public Property Get FormaPago() As XArrayDB
    Set FormaPago = xFormaPago
End Property


Function DescripcionFormaPago(ByVal strFormaPago As String) As String
    
    Dim a As Long
    
    a = xFormaPago.Find(0, 0, strFormaPago)
    DescripcionFormaPago = UCase(xFormaPago(a, 1))

End Function




Public Function AgregaFormaPago(ByVal pstrCodigo As String, _
                    ByVal pstrDescripcion As String) As XArrayDB
    


    Dim ultimo As Integer
    Dim aux As Integer
    If xFormaPago.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xFormaPago.Count(1)
    While i < aux
        If xFormaPago(i, 0) = pstrCodigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xFormaPago.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xFormaPago.AppendRows
    End If
j:
    If xFormaPago.Count(1) = 0 Then ultimo = 0: xFormaPago.AppendRows
    
    xFormaPago(ultimo, 0) = pstrCodigo
    xFormaPago(ultimo, 1) = pstrDescripcion
    Set AgregaFormaPago = xFormaPago


End Function


