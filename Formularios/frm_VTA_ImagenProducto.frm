VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm_VTA_ImagenProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imagen Producto"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "Siguiente -->"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<-- Anterior"
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4815
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   6615
         ExtentX         =   11668
         ExtentY         =   8493
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.PictureBox pctrImagenProducto 
         Height          =   4815
         Left            =   0
         ScaleHeight     =   4755
         ScaleWidth      =   6555
         TabIndex        =   1
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.Label Label8 
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   7200
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CLASE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label6 
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   6840
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "GRUPO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   6480
      Width           =   4815
   End
   Begin VB.Label Label3 
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   6120
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LABORATORIO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_VTA_ImagenProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strCodigoProducto As String
Private strPath As String
Private strFileUrl As String
Private indiceImgActual As Integer
Private cantidadImagenes As Long
Private listaImagenes() As String
Private strProducto As String
Private strLaboratorio As String
Private strGrupo As String
Private strClase As String
Public flgWeb As Boolean
Private strResponse As String
Public cMsg As Integer

Public Property Get codigoproducto() As String
    codigoproducto = strCodigoProducto
End Property

Public Property Let codigoproducto(ByVal s_codProducto As String)
    strCodigoProducto = s_codProducto
End Property

Public Property Get imagePath() As String
    imagePath = strPath
End Property

Public Property Let imagePath(ByVal s_imagePath As String)
    strPath = s_imagePath
End Property

Public Property Get imagenFileUrl() As String
    imagenFileUrl = strFileUrl
End Property

Public Property Let imagenFileUrl(ByVal s_url As String)
    strFileUrl = s_url
End Property
Public Property Get Producto() As String
    Producto = strProducto
End Property
Public Property Let Producto(ByVal s_producto As String)
    strProducto = s_producto
End Property
Public Property Get Laboratorio() As String
    Laboratorio = strLaboratorio
End Property
Public Property Let Laboratorio(ByVal s_laboratorio As String)
    strLaboratorio = s_laboratorio
End Property
Public Property Get grupo() As String
    grupo = strGrupo
End Property
Public Property Let grupo(ByVal s_grupo As String)
    strGrupo = s_grupo
End Property
Public Property Get clase() As String
    clase = strClase
End Property
Public Property Let clase(ByVal s_clase As String)
    strClase = s_clase
End Property



Private Sub cmdAnterior_Click()
   
    indiceImgActual = indiceImgActual - 1
    Call cargarImagenFrame(strPath & listaImagenes(indiceImgActual))
    cmdSiguiente.Enabled = True
    
    If indiceImgActual <= 0 Then
        cmdAnterior.Enabled = False
    Else
        cmdAnterior.Enabled = True
    End If

End Sub

Private Sub cmdSiguiente_Click()
    
    indiceImgActual = indiceImgActual + 1
   
    Call cargarImagenFrame(strPath & listaImagenes(indiceImgActual))
    cmdAnterior.Enabled = True
    
    If indiceImgActual >= cantidadImagenes Then
        cmdSiguiente.Enabled = False
    Else
        cmdSiguiente.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    MousePointer = vbDefault
   On Error GoTo LoadError
    
    If Not strCodigoProducto = "" Then
        If flgWeb Then
            Dim urlWeb As String
            Dim urlPath As String
            pctrImagenProducto.Visible = False
            WebBrowser1.Visible = False
            '"https://s3-us-west-2.amazonaws.com/inkafarmaproductimages/newimages/"
            urlPath = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "IMGPRODWEB", objUsuario.CodigoEmpresa)
            urlWeb = urlPath & strCodigoProducto & "L.png"
            
            WebBrowser1.Navigate urlWeb
            
            'WebBrowser1.Navigate "about:<html><body scroll=no><img width='100%' src =""V:\003736*.png""></img></body></html>"
        Else
            Call cargaFrame
        End If
        Label3.Caption = strProducto
        Label4.Caption = strLaboratorio
        Label6.Caption = strGrupo
        Label8.Caption = strClase
    End If
    Exit Sub
    
LoadError:
    MsgBox "Se produjo un error al mostrar imagen del producto." & vbNewLine & Err.Description
    Debug.Print Err.Description
    Unload Me
End Sub

Private Sub cargaFrame()
    pctrImagenProducto.Visible = True
    WebBrowser1.Visible = False
    Dim archivosEnCarpeta As String
    Dim lCtr As Long
    Dim urlImagenProducto As String
    
    If imagePath = "" Then
        imagePath = CStr(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FILE_IMAGEN_PRODUCTO"))
    End If
                
    If Dir(imagePath, vbDirectory) = "" Then
        MsgBox "La carpeta compartida de imagenes no existe o no es accesible" & vbCrLf & _
               imagePath, vbCritical + vbOKOnly, App.ProductName
    End If
                
    imagenFileUrl = imagePath & strCodigoProducto & "\"
                
    If Dir(imagenFileUrl, vbDirectory) <> "" Then
        If Dir(imagenFileUrl & "\*.png") <> "" Then
            'mostrar imagenes
        Else
            GoTo noExiste
        End If
    Else
        imagenFileUrl = imagePath & strCodigoProducto & "*.png"
        If Dir(imagenFileUrl) <> "" Then
            'mostrar imagenes
        Else
            GoTo noExiste
        End If
    End If
    
    ReDim listaImagenes(0) As String
    archivosEnCarpeta = Dir(strFileUrl, vbNormal)
      
    Do While Len(archivosEnCarpeta)
      
        If listaImagenes(0) = "" Then
            listaImagenes(0) = archivosEnCarpeta
        Else
            lCtr = UBound(listaImagenes) + 1
            ReDim Preserve listaImagenes(lCtr) As String
            listaImagenes(lCtr) = archivosEnCarpeta
        End If
        archivosEnCarpeta = Dir
    Loop
              
    cantidadImagenes = UBound(listaImagenes)
    Debug.Print "Cantidad de imagnes: " & cantidadImagenes
     
    If cantidadImagenes > 0 Then
        cmdAnterior.Enabled = False
        cmdSiguiente.Enabled = True
             urlImagenProducto = strPath & listaImagenes(0)
        Call cargarImagenFrame(urlImagenProducto)
        indiceImgActual = 0
       
    ElseIf Not listaImagenes(0) = "" Then
        urlImagenProducto = strPath & listaImagenes(0)
        Call cargarImagenFrame(urlImagenProducto)
        
    ElseIf listaImagenes(0) = "" Then
        'Mostrar mensaje que no hay imagenes disponibles
        MsgBox "No se tienen imagenes para el producto seleccionado."
    End If
noExiste:
    MousePointer = vbDefault
    MsgBox "No se encontraron imagenes para el producto seleccionado" & vbCrLf & _
                   imagenFileUrl, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            If cmdSiguiente.Enabled Then cmdSiguiente_Click
        Case vbKeyLeft
            If cmdAnterior.Enabled Then cmdAnterior_Click
        Case vbKeyEscape
            strCodigoProducto = ""
            Unload Me
       End Select
End Sub

Private Sub cargarImagenFrame(ByVal rutaImagen As String)
'    pctrImagenProducto.Picture = LoadPicture(rutaImagen)
'    pctrImagenProducto.ScaleMode = 3
'    pctrImagenProducto.AutoRedraw = True
'    pctrImagenProducto.PaintPicture pctrImagenProducto.Picture, 0, 0, pctrImagenProducto.ScaleWidth, pctrImagenProducto.ScaleHeight, 0, 0, pctrImagenProducto.Picture.Width / 26.46, pctrImagenProducto.Picture.Height / 26.46
    Dim cRenderer As New stdPicEx2
    Dim tPic As StdPicture
    Dim lDPI As Long, cx As Long, cy As Long
    Dim X As Long, Y As Long
    
    pctrImagenProducto.ScaleMode = 3
    pctrImagenProducto.AutoRedraw = True
    
    With pctrImagenProducto
        .Picture = Nothing
        cx = ScaleX(pctrImagenProducto.Picture.Width, vbHimetric, .ScaleMode)
        cy = ScaleY(pctrImagenProducto.Picture.Height, vbHimetric, .ScaleMode)
        X = (.ScaleWidth - cx) ' \ 2
        Y = (.ScaleHeight - cy) ' \ 2
    End With
    Set tPic = cRenderer.LoadPictureEx(rutaImagen, mgtAutoSelect)

    pctrImagenProducto.Picture = cRenderer.CopyStdPicture(tPic, X, Y)
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    cMsg = cMsg + 1
    Dim html As String
    'si url tiene about, no concatenar
    If InStr(1, URL, "about:") > 0 Then
        html = Replace(URL, "about:", "") 'remover about para setear el html directo, sin peticion
    Else
        html = "<html><body scroll=no><img width='100%' src ='" & URL & "'></img></body></html>"
    End If
    strResponse = WebBrowser1.Document.documentElement.innerHTML
    Debug.Print URL
    Debug.Print strResponse
'    MsgBox strResponse
    If InStr(1, strResponse, "HTTP 404") > 0 Or InStr(1, strResponse, "errorPageStrings.js") > 0 Then
        If cMsg = 1 Then Call cargaFrame
    Else
        pctrImagenProducto.Visible = False
        WebBrowser1.Visible = True
        WebBrowser1.Stop
        WebBrowser1.Document.Write html
    End If
End Sub

