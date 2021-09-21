Attribute VB_Name = "mdl_OFF_Variables"
Option Explicit
Global objOFFUsuario As New cls_OFF_Usuario
Global objOFFCliente As New cls_OFF_Cliente
Global objOFFVenta As New cls_OFF_Venta

Public Const COD_ESTADO_EMI = "EMI"
Public Const COD_ESTADO_ANU = "ANU"
Public Const COD_TIPO_BOL = "BOL"
Public Const COD_TIPO_FAC = "FAC"
Public Const COD_TIPO_TKB = "TKB"
Public Const COD_TIPO_TKF = "TKF"
Public Const COD_TIPO_GRL = "GRL"
Public Const COD_TIP_MOV_VENTA = "VTA"
Public Const COD_TIP_MOV_ANULACION = "AVT"
Public Const COD_TIP_MOV_REACTIVACION = "RVT"
Public Const COD_PERFIL_QFI = "0700"
Public Const COD_PERFIL_QFII = "0709"
Public Const COD_FPAGO_EFE_SOLES = "0"
Public Const COD_FPAGO_EFE_DOLAR = "1"
Public Const COD_FPAGO_TARJETA = "2"



Public Declare Function ShellExecute Lib "Shell32.Dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal pOperation As String, ByVal pFile As String, ByVal pParameters As String, ByVal pdirectory As String, ByVal nShowCmd As Long) As Long


