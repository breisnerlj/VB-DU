VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private byteNumElementosLocalAsig As Byte
Private xLocalAsig As New XArrayDB

Public Function Lista(Optional Codigo As String = "") As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_ZONA.FN_LISTA", 0, Codigo)
End Function
Public Function ListaLocalDiferencia(ByVal Cia As String, Optional Codigo As String = "") As oraDynaset
    Set ListaLocalDiferencia = gclsOracle.FN_Cursor("BTLPROD.PKG_ZONA.FN_LISTA_LOCAL_N", 0, Cia, "")
End Function

Public Function Graba(ByRef Codigo As String, _
                      Optional Descripcion As String = "", _
                      Optional Abreviatura As String = "", _
                      Optional Activo As String = "", _
                      Optional Usuario As String = "" _
                      ) As String

    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim a As String
    arrValores = Array(Codigo, Descripcion, Abreviatura, Activo, Usuario)
    arrDireccion = Array(entrada_salida, "1", "1", "1", "1")
    On Error GoTo handle
    Graba = gclsOracle.SP("BTLPROD.PKG_ZONA.SP_GRABA", arrValores, arrDireccion)
    Codigo = arrValores(0)
    Exit Function
handle:
    Graba = "Error al ejecutar procedimiento"
End Function

Public Function ListaLocal(Cia As String, ZONA As String, Optional Localx As String, Optional CIAREF As String) As oraDynaset
    Set ListaLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_ZONA.FN_LISTA_LOCAL", 0, Cia, ZONA, Localx, CIAREF)
End Function
Public Function GrabaLocal(Optional Cia As String, _
                      Optional ZONA As String = "", _
                      Optional CadenaLocales As String = "", _
                      Optional Activo As String = "", _
                      Optional Usuario As String = "" _
                      ) As String

    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim a As String
    arrValores = Array(Cia, ZONA, CadenaLocales, Activo, Usuario)
    arrDireccion = Array("1", "1", "1", "1", "1")
    On Error GoTo handle
    GrabaLocal = gclsOracle.SP("BTLPROD.PKG_ZONA.SP_GRABA_LOCAL", arrValores, arrDireccion)
    
    Exit Function
handle:
    GrabaLocal = "Error al ejecutar procedimiento"
End Function


Public Function ListaZonaLocal(ByVal pCia As String, ByVal pCodLocal As String) As oraDynaset
    On Error GoTo handle
    
    
    Set ListaZonaLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_ZONA.FN_LISTA_ZONA_LOCAL", 0, pCia, pCodLocal)
    
    Exit Function
handle:
    Err.Raise Err.Number, "clsZona.ListaZonaLocal", Err.Description

End Function



Public Function Lista_Locales(Optional CodCia As String = "", Optional CodZona As String = "", Optional CodLocal As String = "") As oraDynaset
    On Error GoTo CtrlError
    Set Lista_Locales = gclsOracle.FN_Cursor("BTLPROD.PKG_ZONA.FN_LISTA_LOCAL", 0, CodCia, CodZona, CodLocal)
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsZona.Lista_Locales", Err.Description
End Function

Public Sub InicializaLista()
    Dim i As Integer
    If xLocalAsig.Count(1) > 0 Then
        For i = 0 To xLocalAsig.Count(1) - 1
            Call xLocalAsig.DeleteRows(xLocalAsig.UpperBound(1))
        Next
    End If
End Sub

Private Sub Class_Initialize()
    xLocalAsig.ReDim 0, -1, 0, 2  'Setea el array
End Sub

Public Property Get LocalAsig() As XArrayDB
    Set LocalAsig = xLocalAsig 'Devuelve la variable de la clase en forma de XArray
End Property

Public Function AgregaLocal(ByVal vstrcia$, _
                            ByVal CodZona$, _
                            ByVal vstrCadCodBtl$) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xLocalAsig.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean

    aux = xLocalAsig.Count(1)
    While i < aux
        If xLocalAsig(i, 2) = vstrCadCodBtl Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xLocalAsig.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xLocalAsig.AppendRows
    End If
j:

    If xLocalAsig.Count(1) = 0 Then ultimo = 0: xLocalAsig.AppendRows
        'xLocalAsig(xLocalAsig.UpperBound(1), 0) = vstrcia         ' Asigna el codigo de empresa
        'xLocalAsig(xLocalAsig.UpperBound(1), 1) = CodZona     ' Asigna el cod zona
        'xLocalAsig(xLocalAsig.UpperBound(1), 2) = vstrCadCodBtl   ' Asigna cadena de codigos btl
        xLocalAsig(ultimo, 0) = vstrcia         ' Asigna el codigo de empresa
        xLocalAsig(ultimo, 1) = CodZona     ' Asigna el cod zona
        xLocalAsig(ultimo, 2) = vstrCadCodBtl   ' Asigna cadena de codigos btl
    Set AgregaLocal = xLocalAsig
    'xLocalAsig.Clear
End Function


Public Function Graba_Local(ByVal oradb As OraDatabase, _
                            ByVal Cia As String, _
                            ByVal CodZona As String, _
                            ByVal Activo As Integer) As String
On Error GoTo CtrlErr
Dim i As Integer

    byteNumElementosLocalAsig = IIf(LocalAsig.UpperBound(1) < 0, 1, LocalAsig.UpperBound(1)) + 1

    For i = oradb.Parameters.Count - 1 To 0 Step -1
       oradb.Parameters.Remove i
    Next

    oradb.Parameters.Add "A_CIA", objUsuario.CodigoEmpresa, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_COD_ZONA", CodZona, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_FLG_ACTIVO", Activo, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_COD_USUARIO", objUsuario.Codigo, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.AddTable "A_CAD_COD_LOCAL", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosLocalAsig, 200

    Dim CadLocal As OracleInProcServer.OraParamArray
    Set CadLocal = oradb.Parameters("A_CAD_COD_LOCAL")

    '---- Mandando el arreglo al SP -----'
    Debug.Print "***********************************"
    For i = 0 To LocalAsig.UpperBound(1)
        CadLocal(i) = LocalAsig(i, 2)
        Debug.Print CadLocal(i)
    Next

    oradb.ExecuteSQL " BEGIN BTLPROD.PKG_ZONA.SP_GRABA_LOCAL(:A_CIA, :A_COD_ZONA, :A_CAD_COD_LOCAL, :A_FLG_ACTIVO, :A_COD_USUARIO) ; END;"
Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsZona.Graba_Local", Err.Description
End Function

Public Function Lista_Locales_Nombres(Optional CodCia As String = "", Optional CodZona As String = "", Optional DesLocal As String = "") As oraDynaset
    On Error GoTo CtrlError
    Set Lista_Locales_Nombres = gclsOracle.FN_Cursor("BTLPROD.PKG_ZONA.FN_LISTA_LOCAL_NOMBRE", 0, CodCia, CodZona, DesLocal)
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsZona.Lista_Locales_Nombres", Err.Description
End Function



Public Function DevuelveZona(ByVal CodigoLocal As String) As String
    On Error GoTo CtrlError
Dim StrCodigoZona As String
StrCodigoZona = "" & gclsOracle.FN_Valor("BTLPROD.FN_DEV_ZONA_LOCAL", CodigoLocal)
DevuelveZona = StrCodigoZona
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsZona.DevuelveZona", Err.Description
End Function
