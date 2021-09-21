VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function ListaHistorial(Optional ByVal CodLocal As String) As oraDynaset
    Set ListaHistorial = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_LISTA_HISTORIAL", 0, CodLocal)
End Function
Public Function ListaDetalleHistorial(Optional ByVal codGuia As String) As oraDynaset
    Set ListaDetalleHistorial = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_LISTA_DETALLE_HISTORIAL", 0, codGuia)
End Function
Public Function cabeceraRecepcion(Optional ByVal codGuia As String) As oraDynaset
    Set cabeceraRecepcion = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_CABECERA_RECEPCION", 0, codGuia)
End Function
Public Function listaProductoSolicitud(Optional ByVal codSolicitud As String) As oraDynaset
    Set listaProductoSolicitud = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_LISTA_PRODUCTO_SOLICITUD", 0, codSolicitud)
End Function
Public Function listaProductoGenerico(Optional ByVal desProducto As String) As oraDynaset
    Set listaProductoGenerico = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_LISTA_PRODUCTO_GENERICO", 0, desProducto)
End Function
Public Function periodoGenerado(Optional ByVal codSolicitud As String) As oraDynaset
    Set periodoGenerado = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_PERIODO_GENERADO", 0, codSolicitud)
End Function
Public Function detalleRecepcion(Optional ByVal codGuia As String) As oraDynaset
    Set detalleRecepcion = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_DETALLE_RECEPCION", 0, codGuia)
End Function
Public Function llenaFechas(Optional ByVal CodLocal As String) As oraDynaset
    Set llenaFechas = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_LLENA_FECHAS", 0)
End Function
Public Function Parametro(Optional ByVal CodLocal As String) As oraDynaset
    Set Parametro = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_PARAMETRO", 0)
End Function
Public Function listaGrilla(Optional ByVal codArea As String) As oraDynaset
    Set listaGrilla = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_LISTA_GRILLA", 0, codArea)
End Function
Public Function listaGuias(Optional ByVal codArea As String) As oraDynaset
    Set listaGuias = gclsOracle.FN_Cursor("SSGG.PKG_SERVICIOS_GENERALES.FN_LISTAGUIA", 0, codArea)
End Function

'******* Configuracion de Grilla**********
Sub spGrilla_Carga(ByRef rgrd As Object, _
                   ByVal vvarCaption As Variant, _
                   Optional ByVal vvarWidth As Variant, _
                   Optional ByVal vvarAlignment As Variant, _
                   Optional ByVal vvarDataSource As Variant, _
                   Optional ByVal vblnAllowSizing As Boolean)
    Dim columna As TrueDBGrid70.Column
    Dim i As Integer
    rgrd.Columns.Clear
    DoEvents
    For i = 0 To UBound(vvarCaption)
        Set columna = rgrd.Columns.Add(i)
        If Not IsMissing(vvarWidth) Then rgrd.Columns(i).Width = vvarWidth(i)
        If Not IsMissing(vvarAlignment) Then rgrd.Columns(i).Alignment = vvarAlignment(i)
        If Not IsMissing(vvarDataSource) Then rgrd.Columns(i).DataField = vvarDataSource(i)
        rgrd.Columns(i).AllowSizing = vblnAllowSizing
        rgrd.Columns(i).Visible = True
        rgrd.Columns(i).Caption = vvarCaption(i)
        rgrd.Columns(i).WrapText = True
    Next i
    rgrd.AllowAddNew = False
    rgrd.Splits(0).AllowColSelect = False
    rgrd.Splits(0).AllowRowSelect = True
    rgrd.Splits(0).AllowRowSizing = False
    rgrd.Splits(0).AllowSizing = False
    rgrd.Splits(0).Style.VerticalAlignment = dbgVertCenter
    rgrd.AllowUpdate = False
    rgrd.HoldFields
    Set columna = Nothing
End Sub


Sub spGrilla_Titulos(ByRef rgrd As TDBGrid, ByVal vvarTitulo As Variant)
    Dim i As Integer
    For i = 0 To UBound(vvarTitulo)
        rgrd.Columns(i).Caption = vvarTitulo(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_Alinea(ByRef rgrd As TDBGrid, ByVal vvarAlinea)
    Dim i As Integer
    For i = 0 To UBound(vvarAlinea)
        rgrd.Columns(i).Alignment = vvarAlinea(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_Ancho(ByRef rgrd As TDBGrid, ByVal vvarAncho)
    Dim i As Integer
    For i = 0 To UBound(vvarAncho)
        rgrd.Columns(i).Width = vvarAncho(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_DatoCampo(ByRef rgrd As TDBGrid, ByVal vvarDatoCampo)
    Dim i As Integer
    For i = 0 To rgrd.Columns.Count - 1
        If i > UBound(vvarDatoCampo) Then rgrd.Columns(i).Visible = False Else rgrd.Columns(i).DataField = vvarDatoCampo(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_CheckBox(ByRef rgrd As TDBGrid, ByVal vvarColumn)
    rgrd.Columns(vvarColumn).ValueItems.Presentation = dbgCheckBox
End Sub

Sub spGrilla_CboBox(ByRef rgrd As TDBGrid, ByVal vvarColumn As Variant, ByVal vvarValue As Variant, ByVal vrst As oraDynaset)
    Dim ValueItem As New TrueDBGrid70.ValueItem
    Dim i As Integer
    rgrd.Columns(vvarColumn).ValueItems.Clear
    vrst.MoveFirst
    While Not vrst.EOF
        ValueItem.Value = vrst(vvarValue).Value
        rgrd.Columns(vvarColumn).ValueItems.Add ValueItem
        vrst.MoveNext
    Wend
    If vrst.RecordCount = 0 Then Exit Sub
    rgrd.Columns(vvarColumn).ValueItems.Presentation = dbgComboBox
    rgrd.Columns(vvarColumn).ValueItems.MaxComboItems = vrst.RecordCount
    rgrd.Columns(vvarColumn).DropDownList = True
    Set ValueItem = Nothing
End Sub

Sub spGrilla_Traslate(ByRef rgrd As TDBGrid, ByVal vvarColumn As Variant, ByVal vvarValue As Variant, ByVal vvarDisplayValue As Variant)
    Dim ValueItem As New TrueDBGrid70.ValueItem
    On Error GoTo ERROR
    rgrd.Columns(vvarColumn).ValueItems.Translate = True
    ValueItem.DisplayValue = vvarDisplayValue
    ValueItem.Value = vvarValue
    rgrd.Columns(vvarColumn).ValueItems.Add ValueItem
    Set ValueItem = Nothing
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Sub spGrilla_Locked_Columns(ByRef rgrd As TDBGrid, Optional ByVal vintColumna%)
    Dim i As Integer
    rgrd.AllowUpdate = False
    With rgrd
        For i = 0 To .Columns.Count - 1
            .Columns(i).Locked = True
            .Columns(i).AllowFocus = False
            If Not IsMissing(vintColumna) Then If (i = vintColumna) Then .Columns(i).Locked = False
            DoEvents
        Next i
        DoEvents
    End With
End Sub

Sub spGrilla_UnLocked_Column(ByRef rgrd As TDBGrid, ByVal vvarColumna As Variant)
    DoEvents
    rgrd.AllowUpdate = True
    With rgrd
            .Columns(vvarColumna).Locked = False
            .Columns(vvarColumna).AllowFocus = True
    End With
End Sub
'*Funcion que verifica la resolucion de la pantalla del equipo local,

'devuelve un valor long

Public Function verifica_resolucion() As Long

    Dim largo As Long, alto As Long, PixelX As Long, PixelY As Long

    PixelX = Screen.TwipsPerPixelX

    PixelY = Screen.TwipsPerPixelY

    largo = Screen.Width \ PixelX

    alto = Screen.Height \ PixelY

    verifica_resolucion = largo + alto

    'verifica_resolucion = 800

    

End Function



