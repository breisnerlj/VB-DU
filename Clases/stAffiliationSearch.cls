VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stAffiliationSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvartransactionid As String 'copia local
Private mvarcardnumber As String 'copia local
Private mvaraccountnumber As String 'copia local
Private mvarcardname As String 'copia local
Private mvarcardlastname As String 'copia local
Private mvarcardbirthdate As String 'copia local
Private mvarcardgender As String 'copia local
Private mvarcarddoctor As String 'copia local
Private mvarcardcontactpreferences As String 'copia local
Private mvarcardemail As String 'copia local
Private mvarcardphone As String 'copia local
Private mvarcardcellphone As String 'copia local
Private mvarcardaddress As String 'copia local
Private mvarcardrelated As String 'copia local
Private mvarcardfields As String 'copia local
Private mvarerrorid As String 'copia local
Private mvarmessage As String 'copia local
Private mvartransactiondate As String 'copia local

Public Function toString() As String
End Function

Public Property Let transactiondate(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.transactiondate = 5
    mvartransactiondate = vData
End Property

Public Property Get transactiondate() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.transactiondate
    transactiondate = mvartransactiondate
End Property

Public Property Let message(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.message = 5
    mvarmessage = vData
End Property

Public Property Get message() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.message
    message = mvarmessage
End Property

Public Property Let errorid(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.errorid = 5
    mvarerrorid = vData
End Property

Public Property Get errorid() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.errorid
    errorid = mvarerrorid
End Property

Public Property Let cardfields(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardfields = 5
    mvarcardfields = vData
End Property

Public Property Get cardfields() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardfields
    cardfields = mvarcardfields
End Property

Public Property Let cardrelated(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardrelated = 5
    mvarcardrelated = vData
End Property

Public Property Get cardrelated() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardrelated
    cardrelated = mvarcardrelated
End Property

Public Property Let cardaddress(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardaddress = 5
    mvarcardaddress = vData
End Property

Public Property Get cardaddress() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardaddress
    cardaddress = mvarcardaddress
End Property

Public Property Let cardcellphone(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardcellphone = 5
    mvarcardcellphone = vData
End Property

Public Property Get cardcellphone() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardcellphone
    cardcellphone = mvarcardcellphone
End Property

Public Property Let cardphone(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardphone = 5
    mvarcardphone = vData
End Property

Public Property Get cardphone() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardphone
    cardphone = mvarcardphone
End Property

Public Property Let cardemail(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardemail = 5
    mvarcardemail = vData
End Property

Public Property Get cardemail() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardemail
    cardemail = mvarcardemail
End Property

Public Property Let cardcontactpreferences(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardcontactpreferences = 5
    mvarcardcontactpreferences = vData
End Property

Public Property Get cardcontactpreferences() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardcontactpreferences
    cardcontactpreferences = mvarcardcontactpreferences
End Property

Public Property Let carddoctor(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.carddoctor = 5
    mvarcarddoctor = vData
End Property

Public Property Get carddoctor() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.carddoctor
    carddoctor = mvarcarddoctor
End Property

Public Property Let cardgender(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardgender = 5
    mvarcardgender = vData
End Property

Public Property Get cardgender() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardgender
    cardgender = mvarcardgender
End Property

Public Property Let cardbirthdate(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardbirthdate = 5
    mvarcardbirthdate = vData
End Property

Public Property Get cardbirthdate() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardbirthdate
    cardbirthdate = mvarcardbirthdate
End Property

Public Property Let cardlastname(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardlastname = 5
    mvarcardlastname = vData
End Property

Public Property Get cardlastname() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardlastname
    cardlastname = mvarcardlastname
End Property

Public Property Let cardname(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardname = 5
    mvarcardname = vData
End Property

Public Property Get cardname() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardname
    cardname = mvarcardname
End Property

Public Property Let accountnumber(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.accountnumber = 5
    mvaraccountnumber = vData
End Property

Public Property Get accountnumber() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.accountnumber
    accountnumber = mvaraccountnumber
End Property

Public Property Let cardnumber(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cardnumber = 5
    mvarcardnumber = vData
End Property

Public Property Get cardnumber() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cardnumber
    cardnumber = mvarcardnumber
End Property

Public Property Let TransactionId(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.transactionid = 5
    mvartransactionid = vData
End Property

Public Property Get TransactionId() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.transactionid
    TransactionId = mvartransactionid
End Property
