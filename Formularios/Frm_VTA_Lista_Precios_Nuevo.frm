VERSION 5.00
Begin VB.Form Frm_VTA_Lista_Precios_Nuevo 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Precios"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Productos que empiezan por :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar Impresión"
         Height          =   855
         Left            =   3720
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   4800
         TabIndex        =   4
         Top             =   5760
         Width           =   1455
      End
      Begin VB.ListBox LstProductos 
         Height          =   6360
         ItemData        =   "Frm_VTA_Lista_Precios_Nuevo.frx":0000
         Left            =   120
         List            =   "Frm_VTA_Lista_Precios_Nuevo.frx":0007
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir Precios"
         Height          =   855
         Left            =   3720
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos los Productos"
         Height          =   495
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Frm_VTA_Lista_Precios_Nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Control
    setteaFormulario Me
    Cargar_Lista_Letras
'    Me.Width = 6750
'    Me.Height = 7320

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Cargar_Lista_Letras()
    Dim i As Variant
    Dim j As Integer
    i = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "Ñ", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    LstProductos.Enabled = False
    While j < UBound(i)
        LstProductos.AddItem i(j), j
        j = j + 1
    Wend
End Sub

Private Sub chkTodos_Click()
    LstProductos.Enabled = IIf(chkTodos.Value = 1, False, True)
End Sub

Private Sub cmdImprimir_Click()
Dim u As Integer

'On Error GoTo Control

'Set loses = CreateObject("OracleInProcServer.XOraSession")
'Set lodb = loses.OpenDatabase("BTLRAC", "BTLCADENA" & "/" & "BTLCADENA", 0&)

    If MsgBox("¿ Listo para imprimir ?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
    
        If chkTodos.Value = 1 Then
            Imprime ""
        Else
        
            Dim intX As Integer
                With LstProductos
                    For intX = 0 To .ListCount - 1
                        If .Selected(intX) = True Then
                            Imprime .List(intX)
                            .List(intX) = .List(intX) & "    Impreso ..."
                        End If
                    Next
                End With
        End If
    End If
'lodb.Close
'Set lodb = Nothing
'Set loses = Nothing

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Imprime(ByVal v_CodProducto As String)
Dim rsResultado As oraDynaset
Dim i%
Dim v_Pagina As Integer

'On Error GoTo Control
    v_Pagina = 1
    'Set rsResultado = objBonificado.Lista_Precios(v_CodProducto, g.cod_local)
    Set rsResultado = gclsOracle.FN_Cursor("BTLPROD.PKG_GESTION.FN_LISTA_PRECIOS_NUEVO", 0, objUsuario.CodigoLocal, v_CodProducto)
    
    If rsResultado.RecordCount > 0 Then
        i = 1
        Printer.FontName = "Draft 17cpi"
        Printer.PaperSize = vbPRPSLetter
        'Printer.Height = 2970
        'Printer.Print "."
        sp_Cabecera (v_Pagina)
        While Not rsResultado.EOF
            Printer.Print rsResultado("TEXTO")
            
            If (i Mod 68) = 0 Then
                v_Pagina = v_Pagina + 1
                'Printer.PaperSize = vbPRPSLetter
                Printer.NewPage
                Printer.PaperSize = vbPRPSLetter
                sp_Cabecera (v_Pagina)
            End If
            rsResultado.MoveNext
            i = i + 1
        Wend
        Printer.EndDoc
    End If
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdSalir_Click()
    Printer.EndDoc
    Unload Me
End Sub

Private Sub sp_Cabecera(v_Pagina As Integer)
Dim strNumRuc$
Dim strDirecc$
Dim strRazSoc$
'On Error GoTo Control

    strNumRuc = pfstr_Val_Parametro("NUMRUC_BTL")
    strDirecc = pfstr_Val_Direccion_Local(objUsuario.CodigoLocal)
    strRazSoc = pfstr_Val_Razon_Social(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)

    Printer.FontName = "Draft 12cpi"
    Printer.Print strRazSoc & Space(77 - Len(strRazSoc)) & "FECHA : " & Date
    Printer.Print Mid(strDirecc, 1, 64) & Space(77 - IIf(Len(strDirecc) >= 64, 64, Len(strDirecc))) & "PAGINA: " & Format(v_Pagina, "####")
    Printer.Print "RUC: " & strNumRuc & Space(39 - 16) & "LISTADO DE PRECIOS"
    Printer.Print Space(39) & "------------------"
    Printer.Print
    Printer.FontName = "Draft 17cpi"
    Printer.Print String(137, "-")
    Printer.Print "CODIGO DESCRIPCION" & Space(87 - 18) & "LABORATORIO" & Space(27) & "PRECIO (S/.)"
    Printer.Print String(137, "-")
    
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub CmdCancelar_Click()
    Printer.EndDoc
End Sub

