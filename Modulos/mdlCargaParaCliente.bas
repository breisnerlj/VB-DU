Attribute VB_Name = "mdlCargaParaCliente"
Option Explicit

    Global gRsLocal As OracleInProcServer.oraDynaset
    Global gRsPais As OracleInProcServer.oraDynaset
    Global gRsDepartamento As oraDynaset
'    Global rsProvincia As oraDynaset
'    Global rsDistrito As oraDynaset
    Global gRsEstadoCivil As OracleInProcServer.oraDynaset
    Global gRsTipoDocumento As OracleInProcServer.oraDynaset
    Global gRsSexo As OracleInProcServer.oraDynaset
    Global gRsSufijo As OracleInProcServer.oraDynaset
    Global gRsSuFijoDirecc As OracleInProcServer.oraDynaset
    Global gRsTipoDireccion As OracleInProcServer.oraDynaset
    Global gRsContacto As OracleInProcServer.oraDynaset
    
Public Sub CargaDatosClienteProv()
    
    Set gRsLocal = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_LOCAL.FN_LISTA ( '" & objUsuario.CodigoEmpresa & "', ''); END; ", "SALIDA", 0)
    Set gRsPais = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_UBIGEO.FN_LISTA_PAIS; END; ", "SALIDA", 0)
    Set gRsDepartamento = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO; END; ", "SALIDA", 0)
'    Set rsProvincia = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, strDepartamento, "", strProvincia)
'    Set rsDistrito = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, strDepartamento, strProvincia, "[ SELECCIONAR ]")
'    Set gRsEstadoCivil = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_ESTADO_CIVIL.FN_LISTA (''); END; ", "SALIDA", 0)
    Set gRsEstadoCivil = Nothing
    Set gRsTipoDocumento = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_DOCUMENTO (''); END; ", "SALIDA", 0)
'    Set gRsSexo = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_CLIENTE.FN_LISTA_SEXO; END; ", "SALIDA", 0)
    Set gRsSexo = Nothing
'    Set gRsSufijo = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_CLIENTE.FN_LISTA_SUFIJO; END; ", "SALIDA", 0)
    Set gRsSufijo = Nothing
'    Set gRsSuFijoDirecc = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_CLIENTE.FN_LST_SUFIJOS_DIREC; END; ", "SALIDA", 0)
    Set gRsSuFijoDirecc = Nothing
    Set gRsTipoDireccion = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_DIRECCION_CEN; END; ", "SALIDA", 0)
    Set gRsContacto = gclsOracle.ODataBase.CreatePlsqlDynaset(" BEGIN :SALIDA := BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_CONT; END; ", "SALIDA", 0)

End Sub

