VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADM_VerSobran 
   Caption         =   "Lista de Sobrantes"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin vbp_Ventas.ctlGrilla grdVerSobran 
      Height          =   4695
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   67764225
         CurrentDate     =   41291
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   67764225
         CurrentDate     =   41291
      End
      Begin VB.Label Label2 
         Caption         =   "Fec. Fin :"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Ini :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_ADM_VerSobran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
