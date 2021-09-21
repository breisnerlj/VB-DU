VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stTransactionUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvartransactionid As String 'copia local
Private mvartransactionitems As String 'copia local
Private mvarcardnumber As String 'copia local
Private mvaraccountnumber As String 'copia local
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



Public Property Let transactionitems(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.transactionitems = 5
    mvartransactionitems = vData
End Property


Public Property Get transactionitems() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.transactionitems
    transactionitems = mvartransactionitems
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



