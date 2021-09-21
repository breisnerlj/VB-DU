VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private xOpciones As New XArrayDB

Private Sub Class_Initialize()
    xOpciones.ReDim 0, -1, 0, 3  'Setea el array de Opciones
End Sub

Public Property Get Opcionesxdb() As XArrayDB
    Set Opcionesxdb = xOpciones 'Devuelve la variable de la clase en forma de XArray
End Property

Public Function Inicializa()
    xOpciones.ReDim 0, -1, 0, 3   'Setea el array de Opciones
End Function




