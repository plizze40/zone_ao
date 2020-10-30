VERSION 5.00
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form frmCaptions 
   BackColor       =   &H00E0E0E0&
   Caption         =   "VerCaptions, Mas efectivo :D... Vamos a banear chiters eM?? DALEE..."
   ClientHeight    =   3075
   ClientLeft      =   2745
   ClientTop       =   5085
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   8280
   Begin Captura.wndCaptura Foto 
      Left            =   1560
      Top             =   2640
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Foto"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   7935
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmCaptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta funci�n Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
                          Lib "user32" ( _
                              ByVal hwnd As Long) As Long

'Esta funci�n retorna el n�mero de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
                          Lib "user32" _
                              Alias "GetWindowTextLengthA" ( _
                              ByVal hwnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenar� el texto devuelto, y el Lenght de la cadena en el �ltimo par�metro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
                          Lib "user32" _
                              Alias "GetWindowTextA" ( _
                              ByVal hwnd As Long, _
                              ByVal lpString As String, _
                              ByVal cch As Long) As Long

'Esta es la funci�n Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
                          Lib "user32" ( _
                              ByVal hwnd As Long, _
                              ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Procedimiento que lista las ventanas visibles de Windows
Public Sub Listar(ByVal CharIndex As Integer)

    Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long

    List1.Clear
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(hwnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then
            'Obtenemos el n�mero de caracteres de la ventana
            lenT = GetWindowTextLength(handle)
            'si es el n�mero anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendr� el tama�o con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y tambi�n debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = left$(titulo, ret)
                'La agregamos al ListBox
                'List1.AddItem titulo$
                Call SendData("PPCC" & titulo$ & "," & CharIndex)
            End If
        End If
        'Buscamos con GetWindow la pr�xima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
    Loop
End Sub
Private Sub command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    Dim X As Integer
    Foto.Area = Ventana
    Foto.Captura
    For X = 1 To 1000
        If Not FileExist(App.path & "/Procesos/" & X & ".bmp", vbNormal) Then Exit For
    Next
    Call SavePicture(Foto.Imagen, App.path & "/Procesos/" & X & ".bmp")
End Sub



