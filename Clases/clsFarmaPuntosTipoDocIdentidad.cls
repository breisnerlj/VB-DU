VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFPTipoDocIdentidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarDNI As String 'copia local
Private mvarCE As String 'copia local

'''Public Property Let CE(ByVal vData As String)
'''    mvarCE = vData
'''End Property

Public Property Get CE() As String
    CE = mvarCE
End Property

'''Public Property Let DNI(ByVal vData As String)
'''    mvarDNI = vData
'''End Property

Public Property Get DNI() As String
    DNI = mvarDNI
End Property

Private Sub Class_Initialize()
    mvarDNI = "D"
    mvarCE = "E"
End Sub
