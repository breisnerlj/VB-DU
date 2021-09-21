VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Public Function Graba(ByVal Codigo As String, ''                      ByVal Estado As String, ''                      ByVal CodigoLocal As String, ''                      ByVal Nombre As String, ''                      ByVal Apellido As String, ''                      ByVal Numero As String, ''                      ByVal Alias As String, ''                      ByVal Activo As String, ''                      ByVal Usuario As String) As String
''Dim gvarValores  As Variant
''Dim gvarIO  As Variant
''    On Error GoTo CtrlErr
''    gvarValores = Array(Codigo, ''                        Estado, ''                        CodigoLocal, ''                        Nombre, ''                        Apellido, ''                        Numero, ''                        Alias, ''                        Activo, ''                        Usuario)
''
''    gvarIO = Array(entrada, ''                   entrada, ''                   entrada, ''                   entrada, ''                   entrada, ''                   entrada, ''                   entrada, ''                   entrada, ''                   entrada)
''
''    Graba = gclsOracle.SP("BTLPROD.PKG_MOTORIZADO.SP_GRABA", ''                                            gvarValores, gvarIO)
''
''    Exit Function
''CtrlErr:
''    Err.Raise Err.Number, "clsMotorizado", Err.Description
''End Function

Public Function Graba(ByVal Codigo As String, ByVal Estado As String, ByVal LocalAsignado As String, ByVal Nombre As String, ByVal Apellido As String, ByVal NroRef As String, ByVal Alias As String, ByVal Activo As Integer) As String
Dim gvarValores  As Variant
Dim gvarIO  As Variant
    On Error GoTo CtrlErr
    gvarValores = Array(Codigo, Estado, LocalAsignado, Nombre, Apellido, NroRef, Alias, Activo, objUsuario.Codigo)
    
    gvarIO = Array(entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
    
    Graba = gclsOracle.SP("BTLPROD.PKG_MOTORIZADO.SP_GRABA", gvarValores, gvarIO)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado.Graba", Err.Description
End Function

Public Function ListaCodigo() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaCodigo = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA_CODIGO", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado", Err.Description
End Function

Public Function Lista(Optional CodMotorizado As String = "", _
                      Optional CodLocal As String = "", _
                      Optional CodEstado As String) As oraDynaset
    On Error GoTo CtrlError
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA", 0, _
                                                CodMotorizado, _
                                                CodLocal, _
                                                CodEstado)
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsMotorizado.Lista", Err.Description
End Function

Public Function Lista_Inc_Todos(Optional ByVal CodigoLocal As String) As oraDynaset
    On Error GoTo CtrlErr
    Set Lista_Inc_Todos = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA_INC_TODOS", 0, CodigoLocal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado.Lista_Inc_Todos", Err.Description
End Function

Public Function ListaEstado(Optional ByVal Codigo As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaEstado = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA_ESTADO", 0, Codigo)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado", Err.Description
End Function

Public Function ListaAsistencia(Optional vstrCodigo As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaAsistencia = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA_ASISTENCIA", 0, vstrCodigo)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado", Err.Description
End Function

Public Function ListaDisponible(Optional CodigoLocal As String, _
                                Optional FchIni As String, _
                                Optional FchFin As String, _
                                Optional FlgTodos As String, _
                                Optional CodMaquina As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaDisponible = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA_DISPONIBLE", 0, CodigoLocal, FchIni, FchFin, CodMaquina, FlgTodos)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado", Err.Description
End Function

Public Function GrabaAsistencia(ByVal vstrCodMotorizado As String, _
                                       ByVal vstrFchAsistencia As String, _
                                       ByVal vstrFchHrIng As String, _
                                       ByVal vstrFchHrSal As String, _
                                       ByVal vstrFchHrIngRef As String, _
                                       ByVal vstrFchHrSalRef As String, _
                                       ByVal vstrFlgActivo As String, _
                                       ByVal vstrCodUsuario As String) As String
    On Error GoTo CtrlErr
    Dim varValores  As Variant
    Dim varIO  As Variant
    varValores = Array(vstrCodMotorizado, _
                        vstrFchAsistencia, _
                        vstrFchHrIng, _
                        vstrFchHrSal, _
                        vstrFchHrIngRef, _
                        vstrFchHrSalRef, _
                        vstrFlgActivo, _
                        vstrCodUsuario)
    
    varIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
    
    GrabaAsistencia = gclsOracle.SP("BTLPROD.PKG_MOTORIZADO.SP_GRABA_ASIST_MOTORIZADO", _
                                            varValores, varIO)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado", Err.Description
End Function

Public Function GrabaAusencia(ByVal Codigo As String, _
                         ByVal Ausencia As Boolean) As String
    On Error GoTo CtrlErr
    Dim varValores  As Variant
    Dim varIO  As Variant
    varValores = Array(Codigo, _
                       IIf(Ausencia = True, "1", "0"))
    
    varIO = Array(entrada, _
                  entrada)
    
    GrabaAusencia = gclsOracle.SP("BTLPROD.PKG_MOTORIZADO.SP_GRABA_AUSENCIA", _
                                            varValores, varIO)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado", Err.Description
End Function

Public Function GrabaEstado(ByVal Codigo As String, _
                         ByVal flgTipo As String, _
                         Usuario) As String
    On Error GoTo CtrlErr
    Dim varValores  As Variant
    Dim varIO  As Variant
    

    varValores = Array(Codigo, _
                       flgTipo, _
                       Usuario)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada)
    
    GrabaEstado = gclsOracle.SP("BTLPROD.PKG_MOTORIZADO.SP_GRABA_ESTADO", _
                                            varValores, varIO)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado", Err.Description
End Function

Public Function ListaApellidos(Optional DesApellidos As String = "", _
                               Optional CodEstado As String = "") As oraDynaset
    On Error GoTo CtrlError
    Set ListaApellidos = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA_APELLIDOS", 0, _
                                                   DesApellidos, CodEstado)
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsMotorizado.ListaApellidos", Err.Description
End Function

Public Function LstEstMotorizado_Act_Ina() As oraDynaset
    On Error GoTo CtrlError
    Set LstEstMotorizado_Act_Ina = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_EST_MOTORIZADO_ACT_INA", 0)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsMotorizado.LstEstMotorizado_Act_Ina", Err.Description
End Function

Public Function ListaRuteadores() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaRuteadores = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTORIZADO.FN_LISTA_RUTEADORES ", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "ListaRuteadores", Err.Description
End Function

Public Function GrabaObservaciones(ByVal Codigo As String, _
                                   ByVal Fecha As Date, _
                                   ByVal Observacion As String)
    Dim varValores  As Variant
    Dim varIO  As Variant
    
    varValores = Array(Codigo, _
                       Format(Fecha, "DD/MM/YYYY HH:MM:SS"), _
                       Observacion, _
                       objUsuario.Codigo)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada, _
                  entrada)
    
    GrabaObservaciones = gclsOracle.SP("BTLPROD.PKG_MOTORIZADO.SP_GRABA_OBSERVACIONES", _
                                            varValores, varIO)
    
End Function

