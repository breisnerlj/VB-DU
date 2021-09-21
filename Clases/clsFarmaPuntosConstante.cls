VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFPConstante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarEstadoTarjeta As clsFPEstadoTarjeta
Private mvarTIPODOCIDENTIDAD As clsFPTipoDocIdentidad
Private mvarEXITO As String
Private mvarNO_CONEXION_ORBIS As String
Private mvarPARAMETRO_INCOMPLETO As String
Private mvarPARAMETRO_INCORRECTO As String
Private mvarVENTA_YA_ENVIADA As String
Private mvarKEY_ORBIS As String
Private mvarDEVOLUCION_YA_APLICADA As String
Private mvarYA_EXISTE_REGISTRO_TARJETA As String
Private mvarTIMEOUT_REQUEST As String
Private mvarTIMEOUT_CONNECT As String
Private mvarTIMEOUT As String
Private mvarTARJETA_NO_AFILIADA As String

Private Sub Class_Initialize()
    Set mvarEstadoTarjeta = New clsFPEstadoTarjeta
    Set mvarTIPODOCIDENTIDAD = New clsFPTipoDocIdentidad

    mvarEXITO = "0"
    mvarNO_CONEXION_ORBIS = "901"
    mvarPARAMETRO_INCOMPLETO = "1"
    mvarPARAMETRO_INCORRECTO = "2"
    mvarVENTA_YA_ENVIADA = "38"
    mvarKEY_ORBIS = "10"
    mvarDEVOLUCION_YA_APLICADA = "18"
    mvarYA_EXISTE_REGISTRO_TARJETA = "14"
    mvarTIMEOUT_REQUEST = "com.sun.xml.internal.ws.request.timeout"
    mvarTIMEOUT_CONNECT = "com.sun.xml.internal.ws.connect.timeout"
    mvarTIMEOUT = 3000
    mvarTARJETA_NO_AFILIADA = "41"
End Sub

Public Property Get EstadoTarjeta() As clsFPEstadoTarjeta
    Set EstadoTarjeta = mvarEstadoTarjeta
End Property

Public Property Get TipoDocIdentidad() As clsFPTipoDocIdentidad
    Set TipoDocIdentidad = mvarTIPODOCIDENTIDAD
End Property

Public Property Get EXITO() As String
    EXITO = mvarEXITO
End Property

Public Property Get NO_CONEXION_ORBIS() As String
    NO_CONEXION_ORBIS = mvarNO_CONEXION_ORBIS
End Property

Public Property Get PARAMETRO_INCOMPLETO() As String
    PARAMETRO_INCOMPLETO = mvarPARAMETRO_INCOMPLETO
End Property

Public Property Get PARAMETRO_INCORRECTO() As String
    PARAMETRO_INCORRECTO = mvarPARAMETRO_INCORRECTO
End Property

Public Property Get VENTA_YA_ENVIADA() As String
    VENTA_YA_ENVIADA = mvarVENTA_YA_ENVIADA
End Property

Public Property Get KEY_ORBIS() As String
    KEY_ORBIS = mvarKEY_ORBIS
End Property

Public Property Get DEVOLUCION_YA_APLICADA() As String
    DEVOLUCION_YA_APLICADA = mvarDEVOLUCION_YA_APLICADA
End Property

Public Property Get YA_EXISTE_REGISTRO_TARJETA() As String
    YA_EXISTE_REGISTRO_TARJETA = mvarYA_EXISTE_REGISTRO_TARJETA
End Property

Public Property Get TIMEOUT_REQUEST() As String
    TIMEOUT_REQUEST = mvarTIMEOUT_REQUEST
End Property

Public Property Get TIMEOUT_CONNECT() As String
    TIMEOUT_CONNECT = mvarTIMEOUT_CONNECT
End Property

Public Property Get TIMEOUT() As String
    TIMEOUT = mvarTIMEOUT
End Property

Public Property Get TARJETA_NO_AFILIADA() As String
    TARJETA_NO_AFILIADA = mvarTARJETA_NO_AFILIADA
End Property
