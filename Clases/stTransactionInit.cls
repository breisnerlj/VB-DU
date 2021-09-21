VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stTransactionInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvartransactionid As String
Private mvarcardnumber As String
Private mvaraccountnumber As String
Private mvarcardname As String
Private mvarcardbalance As String
Private mvarcarditems As String
Private mvarerrorid As String
Private mvarmessage As String
Private mvartransactiondate As String

Public Property Let transactiondate(ByVal vData As String)
    mvartransactiondate = vData
End Property

Public Property Get transactiondate() As String
    transactiondate = mvartransactiondate
End Property

Public Property Let message(ByVal vData As String)
    mvarmessage = vData
End Property

Public Property Get message() As String
    message = mvarmessage
End Property

Public Property Let errorid(ByVal vData As String)
    mvarerrorid = vData
End Property

Public Property Get errorid() As String
    errorid = mvarerrorid
End Property

Public Property Let carditems(ByVal vData As String)
    mvarcarditems = vData
End Property

Public Property Get carditems() As String
    carditems = mvarcarditems
End Property

Public Property Let cardbalance(ByVal vData As String)
    mvarcardbalance = vData
End Property

Public Property Get cardbalance() As String
    cardbalance = mvarcardbalance
End Property

Public Property Let cardname(ByVal vData As String)
    mvarcardname = vData
End Property

Public Property Get cardname() As String
    cardname = mvarcardname
End Property

Public Property Let accountnumber(ByVal vData As String)
    mvaraccountnumber = vData
End Property

Public Property Get accountnumber() As String
    accountnumber = mvaraccountnumber
End Property

Public Property Let cardnumber(ByVal vData As String)
    mvarcardnumber = vData
End Property

Public Property Get cardnumber() As String
    cardnumber = mvarcardnumber
End Property

Public Property Let TransactionId(ByVal vData As String)
    mvartransactionid = vData
End Property

Public Property Get TransactionId() As String
    TransactionId = mvartransactionid
End Property

Public Function toString() As String
    toString = "mvartransactionid . = " & mvartransactionid & vbCrLf & _
               "mvarcardnumber .... = " & mvarcardnumber & vbCrLf & _
               "mvaraccountnumber . = " & mvaraccountnumber & vbCrLf & _
               "mvarcardname ...... = " & mvarcardname & vbCrLf & _
               "mvarcardbalance ... = " & mvarcardbalance & vbCrLf & _
               "mvarcarditems ..... = " & mvarcarditems & vbCrLf & _
               "mvarerrorid ....... = " & mvarerrorid & vbCrLf & _
               "mvarmessage ....... = " & mvarmessage & vbCrLf & _
               "mvartransactiondate = " & mvartransactiondate
End Function

