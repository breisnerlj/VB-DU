VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFPEstadoTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarINVALIDA As String 'copia local
Private mvarINACTIVA As String 'copia local
Private mvarBLOQUEADA As String 'copia local
Private mvarBLOQUEADA_REDIMIR As String 'copia local
Private mvarACTIVA As String 'copia local
Private mvarSIN_ESTADO As String 'copia local

'''Public Property Let SIN_ESTADO(ByVal vData As String)
'''    mvarSIN_ESTADO = vData
'''End Property

Public Property Get SIN_ESTADO() As String
    SIN_ESTADO = mvarSIN_ESTADO
End Property

'''Public Property Let ACTIVA(ByVal vData As String)
'''    mvarACTIVA = vData
'''End Property

Public Property Get ACTIVA() As String
    ACTIVA = mvarACTIVA
End Property

'''Public Property Let BLOQUEADA_REDIMIR(ByVal vData As String)
'''    mvarBLOQUEADA_REDIMIR = vData
'''End Property

Public Property Get BLOQUEADA_REDIMIR() As String
    BLOQUEADA_REDIMIR = mvarBLOQUEADA_REDIMIR
End Property

'''Public Property Let BLOQUEADA(ByVal vData As String)
'''    mvarBLOQUEADA = vData
'''End Property

Public Property Get BLOQUEADA() As String
    BLOQUEADA = mvarBLOQUEADA
End Property

'''Public Property Let INACTIVA(ByVal vData As String)
'''    mvarINACTIVA = vData
'''End Property

Public Property Get INACTIVA() As String
    INACTIVA = mvarINACTIVA
End Property

'''Public Property Let INVALIDA(ByVal vData As String)
'''    mvarINVALIDA = vData
'''End Property

Public Property Get INVALIDA() As String
    INVALIDA = mvarINVALIDA
End Property

Private Sub Class_Initialize()
    mvarINVALIDA = "3"
    mvarINACTIVA = "4"
    mvarBLOQUEADA = "13"
    mvarBLOQUEADA_REDIMIR = "902"
    mvarACTIVA = "903"
    mvarSIN_ESTADO = "900"
End Sub
