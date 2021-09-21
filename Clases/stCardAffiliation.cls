VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stCardAffiliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvartransactionid As String 'copia local
Private mvarcardnumber As String 'copia local
Private mvaraccountnumber As String 'copia local
Private mvarcardbalance As String 'copia local
Private mvarcarditems As String 'copia local
Private mvarerrorid As String 'copia local
Private mvarmessage As String 'copia local
Private mvartransactiondate As String 'copia local

Public Function toString() As String
End Function

Public Property Let TransactionDate(ByVal vData As String)
    mvartransactiondate = vData
End Property

Public Property Get TransactionDate() As String
    TransactionDate = mvartransactiondate
End Property

Public Property Let Message(ByVal vData As String)
    mvarmessage = vData
End Property

Public Property Get Message() As String
    Message = mvarmessage
End Property

Public Property Let ErrorId(ByVal vData As String)
    mvarerrorid = vData
End Property

Public Property Get ErrorId() As String
    ErrorId = mvarerrorid
End Property

Public Property Let CardItems(ByVal vData As String)
    mvarcarditems = vData
End Property

Public Property Get CardItems() As String
    CardItems = mvarcarditems
End Property

Public Property Let CardBalance(ByVal vData As String)
    mvarcardbalance = vData
End Property

Public Property Get CardBalance() As String
    CardBalance = mvarcardbalance
End Property

Public Property Let AccountNumber(ByVal vData As String)
    mvaraccountnumber = vData
End Property

Public Property Get AccountNumber() As String
    AccountNumber = mvaraccountnumber
End Property

Public Property Let CardNumber(ByVal vData As String)
    mvarcardnumber = vData
End Property

Public Property Get CardNumber() As String
    CardNumber = mvarcardnumber
End Property

Public Property Let TransactionId(ByVal vData As String)
    mvartransactionid = vData
End Property

Public Property Get TransactionId() As String
    TransactionId = mvartransactionid
End Property
