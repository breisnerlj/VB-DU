VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAfiliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarNombre As String
Private mvarApPaterno As String
Private mvarApMaterno As String
Private mvarFechaNacimiento As String
Private mvarGenero As String
Private mvarEmail As String
Private mvarTelefono As String
Private mvarDireccion As String
Private mvarDepartamento As String
Private mvarProvincia As String
Private mvarDistrito As String
Private mvarTipoDireccion As String
Private mvarReferencias As String
Private mvarTipoLugar As String
Private mvarCelular As String
Private mvarInformacionAdicional As String
Private mvarDireccionNormalizada As String
Private mvarDNI As String
Private mvarTarjetas As New XArrayDB
Private mvarTipoDni As String
Private mvarCodCliente_Ink As String
Private mvarPuntosDisponibles_Ink As String

Public Property Let puntosDisponibles(ByVal vData As String)
    mvarPuntosDisponibles_Ink = vData
End Property
Public Property Get puntosDisponibles() As String
    puntosDisponibles = mvarPuntosDisponibles_Ink
End Property
Public Property Let CodCliente(ByVal vData As String)
    mvarCodCliente_Ink = vData
End Property

Public Property Get CodCliente() As String
    CodCliente = mvarCodCliente_Ink
End Property

Public Property Let Nombre(ByVal vData As String)
    mvarNombre = vData
End Property

Public Property Get Nombre() As String
     Nombre = mvarNombre
End Property

Public Property Let Genero(ByVal vData As String)
    mvarGenero = vData
End Property

Public Property Get Genero() As String
     Genero = mvarGenero
End Property

Public Property Let Email(ByVal vData As String)
    mvarEmail = vData
End Property

Public Property Get Email() As String
     Email = mvarEmail
End Property

Public Property Let Telefono(ByVal vData As String)
    mvarTelefono = vData
End Property

Public Property Get Telefono() As String
     Telefono = mvarTelefono
End Property

Public Property Let Departamento(ByVal vData As String)
    mvarDepartamento = vData
End Property

Public Property Get Departamento() As String
     Departamento = mvarDepartamento
End Property

Public Property Let Provincia(ByVal vData As String)
    mvarProvincia = vData
End Property

Public Property Get Provincia() As String
     Provincia = mvarProvincia
End Property

Public Property Let Distrito(ByVal vData As String)
    mvarDistrito = vData
End Property

Public Property Get Distrito() As String
     Distrito = mvarDistrito
End Property

Public Property Let TipoDireccion(ByVal vData As String)
    mvarTipoDireccion = vData
End Property

Public Property Get TipoDireccion() As String
     TipoDireccion = mvarTipoDireccion
End Property

Public Property Let Referencias(ByVal vData As String)
    mvarReferencias = vData
End Property

Public Property Get Referencias() As String
     Referencias = mvarReferencias
End Property

Public Property Let TipoLugar(ByVal vData As String)
    mvarTipoLugar = vData
End Property

Public Property Get TipoLugar() As String
     TipoLugar = mvarTipoLugar
End Property

Public Property Let Celular(ByVal vData As String)
    mvarCelular = vData
End Property

Public Property Get Celular() As String
     Celular = mvarCelular
End Property

Public Property Let InformacionAdicional(ByVal vData As String)
    mvarInformacionAdicional = vData
End Property

Public Property Get InformacionAdicional() As String
     InformacionAdicional = mvarInformacionAdicional
End Property

Public Property Let TipoDni(ByVal vData As String)
    mvarTipoDni = vData
End Property

Public Property Get TipoDni() As String
     TipoDni = mvarTipoDni
End Property

Public Property Let DNI(ByVal vData As String)
    mvarDNI = vData
End Property

Public Property Get DNI() As String
     DNI = mvarDNI
End Property

Public Property Let Direccion(ByVal vData As String)
    mvarDireccion = vData
End Property

Public Property Get Direccion() As String
     Direccion = mvarDireccion
End Property

Public Property Let FechaNacimiento(ByVal vData As String)
    mvarFechaNacimiento = vData
End Property

Public Property Get FechaNacimiento() As String
     FechaNacimiento = mvarFechaNacimiento
End Property

Public Property Let DireccionNormalizada(ByVal vData As String)
    mvarDireccionNormalizada = vData
End Property

Public Property Get DireccionNormalizada() As String
    ' IdTipoAvenida|Dirección|IdDepartamento|IdProvincia|IdDistrito|IdTipoDirección|Referencias
     DireccionNormalizada = mvarTipoLugar & "|" & _
                            mvarDireccion & "|" & _
                            mvarDepartamento & "|" & _
                            mvarProvincia & "|" & _
                            mvarDistrito & "|" & _
                            mvarTipoDireccion & "|" & _
                            mvarReferencias
End Property

Public Property Let ApParterno(ByVal vData As String)
    If Not (vData = "") Then
        If (InStr(1, vData, "@") >= 1) Then
            Dim paterno() As String
            paterno = Split(vData, "@")
            mvarApPaterno = paterno(0)
        Else
            mvarApPaterno = vData
        End If
    End If
    End Property

Public Property Get ApParterno() As String
     ApParterno = mvarApPaterno
End Property

Public Property Let ApMarterno(ByVal vData As String)
    If Not (vData = "") Then
        If (InStr(1, vData, "@") >= 1) Then
            Dim materno() As String
            materno = Split(vData, "@")
            mvarApMaterno = materno(1)
        Else
            mvarApMaterno = vData
        End If
    End If
End Property

Public Property Get ApMarterno() As String
     ApMarterno = mvarApMaterno
End Property

Public Property Get Apellidos() As String
    Dim tmp As String
    If (mvarApMaterno <> "" Or mvarApMaterno <> "") Then
        tmp = Trim$(mvarApPaterno) & "@" & Trim$(mvarApMaterno)
    End If
    Apellidos = tmp
End Property

Public Property Let Tarjetas(ByVal vData As XArrayDB)
    Set mvarTarjetas = vData
End Property

Public Property Get Tarjetas() As XArrayDB
     Set Tarjetas = mvarTarjetas
End Property
