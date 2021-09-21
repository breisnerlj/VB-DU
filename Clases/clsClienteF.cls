VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClienteF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvartipodni As String
Private mvarnrodni As String
Private mvarnombre As String
Private mvarapepaterno As String
Private mvarapematerno As String
Private mvarfecnac As Date
Private mvarSexo As String
Private mvardirprefijo As String
Private mvardireccion As String
Private mvardirdepa As String
Private mvardirprov As String
Private mvardirdist As String
Private mvardirtipo As String
Private mvardirreferencia As String
Private mvaremail As String
Private mvartelefono As String
Private mvarmovil As String
Private mvartarjetas As XArrayDB
Private objOrbis As clsOrbisFarma
Private mvaraccountnumber As String

Public Property Get TarjetasAsociadas() As XArrayDB
    Set TarjetasAsociadas = mvartarjetas
End Property

Public Property Let TelefonoMovil(ByVal vData As String)
    mvarmovil = vData
End Property

Public Property Get TelefonoMovil() As String
    TelefonoMovil = mvarmovil
End Property

Public Property Let TelefonoFijo(ByVal vData As String)
    mvartelefono = vData
End Property

Public Property Get TelefonoFijo() As String
    TelefonoFijo = mvartelefono
End Property

Public Property Let Email(ByVal vData As String)
    mvaremail = vData
End Property

Public Property Get Email() As String
    Email = mvaremail
End Property

Public Property Let DireccionReferencia(ByVal vData As String)
    mvardirreferencia = vData
End Property

Public Property Get DireccionReferencia() As String
    DireccionReferencia = mvardirreferencia
End Property

Public Property Let Nombres(ByVal vData As String)
    mvarnombre = vData
End Property

Public Property Get Nombres() As String
    Nombres = mvarnombre
End Property

Public Property Let DireccionTipo(ByVal vData As String)
    mvardirtipo = vData
End Property

Public Property Get DireccionTipo() As String
    DireccionTipo = mvardirtipo
End Property

Public Property Let DireccionProvincia(ByVal vData As String)
    mvardirprov = vData
End Property

Public Property Get DireccionProvincia() As String
    DireccionProvincia = mvardirprov
End Property

Public Property Let DireccionDistrito(ByVal vData As String)
    mvardirdist = vData
End Property

Public Property Get DireccionDistrito() As String
    DireccionDistrito = mvardirdist
End Property

Public Property Let DireccionDepartamento(ByVal vData As String)
    mvardirdepa = vData
End Property

Public Property Get DireccionDepartamento() As String
    DireccionDepartamento = mvardirdepa
End Property

Public Property Let direccion(ByVal vData As String)
    mvardireccion = vData
End Property

Public Property Get direccion() As String
    direccion = mvardireccion
End Property

Public Property Let DireccionPrefijo(ByVal vData As String)
    mvardirprefijo = vData
End Property

Public Property Get DireccionPrefijo() As String
    DireccionPrefijo = mvardirprefijo
End Property

Public Property Let Sexo(ByVal vData As String)
    mvarSexo = vData
End Property

Public Property Get Sexo() As String
    Sexo = mvarSexo
End Property

Public Property Let FechaNacimiento(ByVal vData As Date)
    mvarfecnac = vData
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = mvarfecnac
End Property

Public Property Let ApellidoMaterno(ByVal vData As String)
    mvarapematerno = vData
End Property

Public Property Get ApellidoMaterno() As String
    ApellidoMaterno = mvarapematerno
End Property

Public Property Let ApellidoPaterno(ByVal vData As String)
    mvarapepaterno = vData
End Property

Public Property Get ApellidoPaterno() As String
    ApellidoPaterno = mvarapepaterno
End Property

Private Sub Class_Initialize()
    Set objOrbis = New clsOrbisFarma
    Set mvartarjetas = New XArrayDB
End Sub

Private Sub Class_Terminate()
    Set objOrbis = Nothing
End Sub

Public Property Get tipoDni() As String
    tipoDni = mvartipodni
End Property

Public Property Get numeroDni() As String
    numeroDni = mvarnrodni
End Property

Public Property Get conTarjeta() As Boolean
    conTarjeta = Not (mvartarjetas.Count(1) = 0)
End Property

Private Function getItem(vData As String, vSeparador As String, vIndice As Byte) As String
    Dim tmp() As String
    On Error GoTo Control
    tmp = Split(vData, vSeparador)
    getItem = Trim$(tmp(vIndice))
    Exit Function
Control:
    getItem = ""
End Function

Private Function getAllItems(vData As String, vSeparador As String) As XArrayDB
    Dim tmp() As String, i As Byte
    Dim arr As New XArrayDB
    On Error GoTo Control
    tmp = Split(vData, vSeparador)
    arr.ReDim 0, UBound(tmp), 0, 0
    For i = LBound(tmp) To UBound(tmp)
        arr.Value(i, 0) = Trim$(tmp(i))
    Next i
    Set getAllItems = arr
    Exit Function
Control:
    Set getAllItems = New XArrayDB
End Function

Private Function getDate(vData As String) As Date
    Dim year As Integer, month As Byte, day As Byte, tmpDate As String, Fecha As Date
    
    day = Mid(vData, 7, 2)
    month = Mid(vData, 5, 2)
    year = Mid(vData, 1, 4)
    tmpDate = day & "/" & month & "/" & year
    
    If IsDate(Format(tmpDate, "dd/mm/yyyy")) Then
        Fecha = Format(tmpDate, "dd/mm/yyyy")
    Else
        Fecha = ""
    End If
    
    getDate = Fecha
End Function

Public Function CargarCliente(tipoDni As String, nroDni As String) As Boolean
    Dim oAS As stAffiliationSearch
    
    On Error GoTo Control
    
    mvartipodni = tipoDni
    mvarnrodni = nroDni
    mvaraccountnumber = IIf(tipoDni = "002", "D0", "E") & nroDni
    
    Set oAS = objOrbis.getAffiliationSearch(mvaraccountnumber, objUsuario.Codigo)
    
    If oAS.errorid = "0" Then
        mvarnombre = oAS.CardName
        mvarapepaterno = getItem(oAS.cardlastname, "@", 0)
        mvarapematerno = getItem(oAS.cardlastname, "@", 1)
        mvarfecnac = getDate(oAS.CardBirthdate)
        mvarSexo = IIf("" & oAS.CardGender = "", "F", oAS.CardGender)
        mvardirprefijo = getItem(oAS.cardaddress, "|", 0) 'IdTipoAvenida
        mvardireccion = getItem(oAS.cardaddress, "|", 1) 'Dirección
        mvardirdepa = getItem(oAS.cardaddress, "|", 2) 'IdDepartamento
        mvardirprov = getItem(oAS.cardaddress, "|", 3) 'IdProvincia
        mvardirdist = getItem(oAS.cardaddress, "|", 4) 'IdDistrito
        mvardirtipo = getItem(oAS.cardaddress, "|", 5) 'IdTipoDirección
        mvardirreferencia = getItem(oAS.cardaddress, "|", 6) 'Referencias
        mvaremail = oAS.CardEmail
        mvartelefono = oAS.CardPhone
        mvarmovil = oAS.CardCellPhone
        Set mvartarjetas = getAllItems(oAS.cardrelated, ",")
        CargarCliente = True
    Else
        CargarCliente = False
        Err.Raise oAS.errorid, "clsClienteF", oAS.message
    End If
    
    Set oAS = Nothing
    Exit Function
Control:
    Set oAS = Nothing
    CargarCliente = False
    Err.Raise Err.Number, "clsClienteF", Err.Description
End Function

Public Function AsociarTarjeta(NumeroTarjeta As String) As Boolean
    Dim cardlastname As String, cardaddress As String
    Dim oCA As stCardAffiliation
    
    On Error GoTo Control
    If Trim$(NumeroTarjeta) = "" Then
        AsociarTarjeta = True
        Exit Function
    End If
    
    cardlastname = mvarapepaterno & "@" & mvarapematerno
    ' IdTipoAvenida|Dirección|IdDepartamento|IdProvincia|IdDistrito|IdTipoDirección|Referencias
    cardaddress = mvardirprefijo & "|" & _
                  mvardireccion & "|" & _
                  mvardirdepa & "|" & _
                  mvardirprov & "|" & _
                  mvardirdist & "|" & _
                  mvardirtipo & "|" & _
                  mvardirreferencia
    
    Set oCA = objOrbis.setCardAffiliation(NumeroTarjeta, mvaraccountnumber, mvarnombre, cardlastname, Format$(mvarfecnac, "YYYYMMDD"), mvarSexo, "", "", mvaremail, mvartelefono, mvarmovil, cardaddress, "", objUsuario.Codigo)
    If Not oCA.errorid = "0" Then
        AsociarTarjeta = False
        Err.Raise oCA.errorid, "clsClienteF", oCA.message
    End If
    AsociarTarjeta = True
    Exit Function
Control:
    Set oCA = Nothing
    AsociarTarjeta = False
    Err.Raise Err.Number, "clsClienteF", Err.Description
End Function
