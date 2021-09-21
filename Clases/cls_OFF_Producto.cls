VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_OFF_Producto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strCodigo As String
Private strDescripcion As String
Private strLaboratorio As String
Private strFlgFracciona As String
Private intCtdFracciona As Integer
Private strMoneda As String
Private dblPrecioPublico As Double
Private strConReceta As String
Private strCodigoBarra As String
Private strPartidaArancelaria As String
Private dblPctIgv As Double
Private strAsoSustituto As String
Private dblStock As Double
Private strDescripcionCorta As String
Private dblPctDescuento As Double
Private strFlgRegalo As String
Private xProducto As New XArrayDB





Public Property Get ColProductos() As Integer
    ColProductos = 15
End Property
Public Property Get ColAlternativos() As Integer
    ColAlternativos = 15
End Property



Public Property Get Codigo() As String
    Codigo = strCodigo
End Property
Public Property Let Codigo(ByVal newValue As String)
    strCodigo = newValue
End Property
Public Property Get ColCodigo() As Integer
    ColCodigo = 0
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property
Public Property Let Descripcion(ByVal newValue As String)
    strDescripcion = newValue
End Property
Public Property Get ColDescripcion() As Integer
    ColDescripcion = 1
End Property

Public Property Get Laboratorio() As String
    Laboratorio = strLaboratorio
End Property
Public Property Let Laboratorio(ByVal newValue As String)
    strLaboratorio = newValue
End Property
Public Property Get ColLaboratorio() As Integer
    ColLaboratorio = 2
End Property

Public Property Get FlgFracciona() As String
    FlgFracciona = strFlgFracciona
End Property
Public Property Let FlgFracciona(ByVal newValue As String)
    strFlgFracciona = newValue
End Property
Public Property Get ColFlgFracciona() As Integer
    ColFlgFracciona = 3
End Property

Public Property Get CtdFracciona() As Integer
    CtdFracciona = intCtdFracciona
End Property
Public Property Let CtdFracciona(ByVal newValue As Integer)
    intCtdFracciona = newValue
End Property
Public Property Get ColCtdFracciona() As Integer
    ColCtdFracciona = 4
End Property

Public Property Get Moneda() As String
    Moneda = strMoneda
End Property
Public Property Let Moneda(ByVal newValue As String)
    strMoneda = newValue
End Property
Public Property Get ColMoneda() As Integer
    ColMoneda = 5
End Property

Public Property Get PrecioPublico() As Double
    PrecioPublico = dblPrecioPublico
End Property
Public Property Let PrecioPublico(ByVal newValue As Double)
    dblPrecioPublico = newValue
End Property
Public Property Get ColPrecioPublico() As Integer
    ColPrecioPublico = 6
End Property

Public Property Get ConReceta() As String
    ConReceta = strConReceta
End Property
Public Property Let ConReceta(ByVal newValue As String)
    strConReceta = newValue
End Property
Public Property Get ColConReceta() As Integer
    ColConReceta = 7
End Property

Public Property Get CodigoBarra() As String
    CodigoBarra = strCodigoBarra
End Property
Public Property Let CodigoBarra(ByVal newValue As String)
    strCodigoBarra = newValue
End Property
Public Property Get ColCodigoBarra() As Integer
    ColCodigoBarra = 8
End Property

Public Property Get PartidaArancelaria() As String
    PartidaArancelaria = strPartidaArancelaria
End Property
Public Property Let PartidaArancelaria(ByVal newValue As String)
    strPartidaArancelaria = newValue
End Property
Public Property Get ColPartidaArancelaria() As Integer
    ColPartidaArancelaria = 9
End Property

Public Property Get PctIgv() As Double
    PctIgv = dblPctIgv
End Property
Public Property Let PctIgv(ByVal newValue As Double)
    dblPctIgv = newValue
End Property
Public Property Get ColPctIgv() As Integer
    ColPctIgv = 10
End Property

Public Property Get AsoSustituto() As String
    AsoSustituto = strAsoSustituto
End Property
Public Property Let AsoSustituto(ByVal newValue As String)
    strAsoSustituto = newValue
End Property
Public Property Get ColAsoSustituto() As Integer
    ColAsoSustituto = 11
End Property

Public Property Get Stock() As Double
    Stock = dblStock
End Property
Public Property Let Stock(ByVal newValue As Double)
    dblStock = newValue
End Property
Public Property Get ColStock() As Integer
    ColStock = 12
End Property

Public Property Get DescripcionCorta() As String
    DescripcionCorta = strDescripcionCorta
End Property
Public Property Let DescripcionCorta(ByVal newValue As String)
    strDescripcionCorta = newValue
End Property
Public Property Get ColDescripcionCorta() As Integer
    ColDescripcionCorta = 13
End Property

Public Property Get PctDescuento() As Double
    PctDescuento = dblPctDescuento
End Property
Public Property Let PctDescuento(ByVal newValue As Double)
    dblPctDescuento = newValue
End Property
Public Property Get ColPctDescuento() As Integer
    ColPctDescuento = 14
End Property

Public Property Get FlgRegalo() As String
    FlgRegalo = strFlgRegalo
End Property
Public Property Let FlgRegalo(ByVal newValue As String)
    strFlgRegalo = newValue
End Property
Public Property Get ColFlgRegalo() As Integer
    ColFlgRegalo = 15
End Property





'Public Function LeerTxtProductos() As XArrayDB
'On Error GoTo Handle
'
'    Dim rst As New ADODB.Recordset
'    Dim xdb As New XArrayDB
'    Dim i As Integer
'
'    xdb.ReDim 0, -1, 0, ColProductos
'
'    'rst.Open "Select * from precios.txt ", gstrConexion, adOpenStatic, adLockOptimistic, adCmdText
'    rst.Open "Select * from precios.txt ", gstrConexion, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not rst.EOF Then
'
'        i = 0
'
'        rst.MoveFirst
'        While Not rst.EOF
'
'            xdb.AppendRows
'            xdb(i, ColCodigo) = rst.Fields("CODIGO") & ""
'            xdb(i, ColDescripcion) = rst.Fields("DESCRIPCION") & ""
'            xdb(i, ColLaboratorio) = rst.Fields("LABORATORIO") & ""
'            xdb(i, ColFlgFracciona) = rst.Fields("FLG_FRACCIONA") & ""
'            xdb(i, ColCtdFracciona) = Val(rst.Fields("CTD_FRACCIONA") & "")
'            xdb(i, ColMoneda) = rst.Fields("MONEDA") & ""
'            xdb(i, ColPrecioPublico) = Val(rst.Fields("PRECIO_PUBLICO") & "")
'            xdb(i, ColConReceta) = rst.Fields("CON_RECETA") & ""
'            xdb(i, ColCodigoBarra) = rst.Fields("CODIGO_BARRA") & ""
'            xdb(i, ColPartidaArancelaria) = rst.Fields("PARTIDA_ARANCELARIA") & ""
'            xdb(i, ColPctIgv) = Val(rst.Fields("PCT_IGV") & "")
'            xdb(i, ColAsoSustituto) = rst.Fields("ASO_SUSTITUTO") & ""
'            xdb(i, ColStock) = Val(rst.Fields("STOCK") & "")
'            xdb(i, ColDescripcionCorta) = rst.Fields("DESCRIPCION_CORTA") & ""
'            xdb(i, ColPctDescuento) = Val(rst.Fields("PCT_DESCUENTO") & "")
'            xdb(i, ColFlgRegalo) = rst.Fields("FLG_REGALO") & ""
'
'            i = i + 1
'            rst.MoveNext
'        Wend
'
'    End If
'
'    Set LeerTxtProductos = xdb
'
'    Set rst = Nothing
'    Set xdb = Nothing
'
'    Exit Function
'Handle:
'    Err.Raise Err.Number, "cls_OFF_Productos.LeerTxtProductos", Err.Description
'End Function

'Public Function LeerTxtAlternativos() As XArrayDB
'On Error GoTo Handle
'
'    Dim rst As New ADODB.Recordset
'    Dim xdb As New XArrayDB
'    Dim i As Integer
'
'    xdb.ReDim 0, -1, 0, ColAlternativos
'
'    'rst.Open "select * from precios.txt order by ASO_SUSTITUTO, DESCRIPCION", gstrConexion, adOpenStatic, adLockOptimistic, adCmdText
'    rst.Open "select * from precios.txt order by ASO_SUSTITUTO, DESCRIPCION", gstrConexion, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not rst.EOF Then
'
'        i = 0
'
'        rst.MoveFirst
'        While Not rst.EOF
'
'            xdb.AppendRows
'            xdb(i, ColCodigo) = rst.Fields("CODIGO") & ""
'            xdb(i, ColDescripcion) = rst.Fields("DESCRIPCION") & ""
'            xdb(i, ColLaboratorio) = rst.Fields("LABORATORIO") & ""
'            xdb(i, ColFlgFracciona) = rst.Fields("FLG_FRACCIONA") & ""
'            xdb(i, ColCtdFracciona) = Val(rst.Fields("CTD_FRACCIONA") & "")
'            xdb(i, ColMoneda) = rst.Fields("MONEDA") & ""
'            xdb(i, ColPrecioPublico) = Val(rst.Fields("PRECIO_PUBLICO") & "")
'            xdb(i, ColConReceta) = rst.Fields("CON_RECETA") & ""
'            xdb(i, ColCodigoBarra) = rst.Fields("CODIGO_BARRA") & ""
'            xdb(i, ColPartidaArancelaria) = rst.Fields("PARTIDA_ARANCELARIA") & ""
'            xdb(i, ColPctIgv) = Val(rst.Fields("PCT_IGV") & "")
'            xdb(i, ColAsoSustituto) = rst.Fields("ASO_SUSTITUTO") & ""
'            xdb(i, ColStock) = Val(rst.Fields("STOCK") & "")
'            xdb(i, ColDescripcionCorta) = rst.Fields("DESCRIPCION_CORTA") & ""
'            xdb(i, ColPctDescuento) = Val(rst.Fields("PCT_DESCUENTO") & "")
'            xdb(i, ColFlgRegalo) = rst.Fields("FLG_REGALO") & ""
'
'            i = i + 1
'            rst.MoveNext
'        Wend
'
'    End If
'
'    Set LeerTxtAlternativos = xdb
'
'    Set rst = Nothing
'    Set xdb = Nothing
'
'    Exit Function
'Handle:
'    Err.Raise Err.Number, "cls_OFF_Productos.LeerTxtAlternativos", Err.Description
'End Function

Public Function LeerTxtProductos() As ADODB.Recordset
    Dim rst As ADODB.Recordset

    On Error GoTo Handle
    
    Set rst = New ADODB.Recordset
'    rst.Open "Select * from precios.txt ", gstrConexion, adOpenStatic, adLockReadOnly, adCmdText
    rst.CursorLocation = adUseClient
    rst.Open strPreciosXML, gstrConexion, adOpenStatic, adLockReadOnly
    'Set rst.ActiveConnection = Nothing
    
    Set LeerTxtProductos = rst
    
    Set rst = Nothing
    Exit Function
Handle:
    Err.Raise Err.Number, "cls_OFF_Productos.LeerTxtProductos", Err.Description
End Function

Public Function LeerTxtAlternativos() As ADODB.Recordset
    Dim rst As ADODB.Recordset

    On Error GoTo Handle
    
    Set rst = New ADODB.Recordset
'    rst.Open "Select * from precios.txt ", gstrConexion, adOpenStatic, adLockReadOnly, adCmdText
    rst.CursorLocation = adUseClient
    rst.Open strPreciosXML, gstrConexion, adOpenStatic, adLockReadOnly
    'Set rst.ActiveConnection = Nothing
    
    Set LeerTxtAlternativos = rst
    
    Set rst = Nothing
    Exit Function
Handle:
    Err.Raise Err.Number, "cls_OFF_Productos.LeerTxtAlternativos", Err.Description
End Function


Public Function AgregaProducto(Codigo As String, _
                            Descripcion As String, _
                            Laboratorio As String, _
                            FlagFracciona As String, _
                            CtdFracciona As Integer, _
                            Moneda As String, _
                            PrecioPublico As Double, _
                            ConReceta As String, _
                            CodigoBarra As String, _
                            PartidaArancelaria As String, _
                            PctIgv As Double, _
                            AsoSustituto As String, _
                            Stock As Integer, _
                            DescripcionCorta As String, _
                            PctDescuento As Double, _
                            FlgRegalo As String) As XArrayDB
    Dim ultimo As Integer
    Dim aux As Integer
    If xProducto.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xProducto.Count(1)
    While i < aux
        If xProducto(i, 0) = Codigo And xProducto(i, 11) = FlgRegalo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xProducto.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xProducto.AppendRows
    End If
j:
    If xProducto.Count(1) = 0 Then ultimo = 0: xProducto.AppendRows
    
    
    xProducto(ultimo, 0) = Codigo
    xProducto(ultimo, 1) = Descripcion
    xProducto(ultimo, 2) = Laboratorio
    xProducto(ultimo, 3) = FlagFracciona
    xProducto(ultimo, 4) = CtdFracciona
    xProducto(ultimo, 5) = Moneda
    xProducto(ultimo, 6) = PrecioPublico
    xProducto(ultimo, 7) = ConReceta
    xProducto(ultimo, 8) = CodigoBarra
    xProducto(ultimo, 9) = PartidaArancelaria
    xProducto(ultimo, 10) = PctIgv
    xProducto(ultimo, 11) = AsoSustituto
    xProducto(ultimo, 12) = Stock
    xProducto(ultimo, 13) = DescripcionCorta
    xProducto(ultimo, 14) = PctDescuento
    xProducto(ultimo, 15) = FlgRegalo
    Set AgregaProducto = xProducto
    
End Function

Private Sub Class_Initialize()
    xProducto.ReDim 0, -1, 0, 15
End Sub


Public Function LimpiaProducto()
    xProducto.ReDim 0, -1, 0, 15

End Function


Public Property Get Producto() As XArrayDB
    Set Producto = xProducto
End Property



Public Function IndicadorReceta(ByVal pstrIndicador As String) As String

    On Error GoTo CntrlError
    
    Select Case pstrIndicador
        Case 1
            IndicadorReceta = "VENTA LIBRE"
        Case 2
            IndicadorReceta = "PRESENTAR RECETA"
        Case 3
            IndicadorReceta = "RETENER RECETA"
        Case 4
            IndicadorReceta = "CON RECETA (OPCIONAL)"
        Case Else
            IndicadorReceta = "*"
    End Select
    
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "cls_OFF_Producto.IndicadorReceta", Err.Description



End Function


