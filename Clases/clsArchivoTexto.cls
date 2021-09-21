VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchivoTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim NumArch As Integer

 



Private Sub Class_Initialize()
NumArch = FreeFile


Open "C:\CajeroCorresponsal.txt" For Append As NumArch


End Sub


Public Sub WriteArchivo(ByVal pp As String)
Write #NumArch, pp



End Sub




Private Sub Class_Terminate()
Close #NumArch


End Sub
