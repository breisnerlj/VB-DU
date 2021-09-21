VERSION 5.00
Begin VB.Form frmRecargasdirec 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Esperando respuesta"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "00 sg."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   7080
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmRecargasdirec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim PMensaje As String
Dim WstrMensaje As String
Dim WstrLocal As String
Public Function carga(ByVal a_ccod_local_in As String, _
            ByVal A_cnumpedido_in As String, _
            ByVal a_ctelefono_in As String, _
            ByVal A_cmonto_in As String, _
            ByVal A_cusu_crea_in As String, _
            ByVal a_cterminal_in As String, _
            ByVal A_ctipo_rcd_in As String _
) As String

WstrLocal = a_ccod_local_in
WstrMensaje = conectaOracle(a_ccod_local_in, _
            A_cnumpedido_in, _
            a_ctelefono_in, _
            A_cmonto_in, _
            A_cusu_crea_in, _
            a_cterminal_in, _
            A_ctipo_rcd_in)

Timer1.Interval = 1000
Timer1.Enabled = True
i = 0
Me.Show vbModal
carga = PMensaje
End Function



Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
'i = 0
    'Form1.Timer1.Interval = 60
    Dim Mensaje As String
    i = i + 1
    'Label2.Visible = True
    
    Label1.Caption = "" & i & " sg."
    If (i Mod 5) = 0 Then
        Mensaje = Respuesta(WstrLocal, WstrMensaje)
        If fncPalote(fncPalote(Mensaje, 1, "|"), 0, "|") <> "" Then
            
                PMensaje = Mensaje
            
            
            
            i = 60
        End If
    End If
    
    If i = 60 Then
        Unload Me
    End If
    'Label2.Visible = False
End Sub
