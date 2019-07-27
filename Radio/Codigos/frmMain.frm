VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radio LwK"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3480
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop ||"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Comprobando estado de la Radio..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "www.RadioXtreme.com.ar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3135
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6588
      _cy             =   5530
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "Opciones"
      Begin VB.Menu MnuMinimizar 
         Caption         =   "Minimizar"
      End
      Begin VB.Menu Mnusalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -- Api SetForegroundWindow Para traer la ventana al frente
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
' -- Api para desplegar el cuadro de diálogo Acerca de ...
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

' -- Estructura NOTIFYICONDATA
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' -- Constantes para las acciones
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

' -- Constantes para los botones y le mouse (mensajes)
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

' -- Función Api Shell_NotifyIcon
Private Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

' -- variables para la estructura NOTIFYICONDATA
Dim systray As NOTIFYICONDATA







Private Sub Form_MouseMove( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, Y As Single)

Dim msg As Long

    If (Me.ScaleMode = vbPixels) Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If

    Select Case msg
        Case WM_LBUTTONDBLCLK
            ' -- Si hacemos doble click con el botón izquierdo restauramos el form
            Me.WindowState = vbNormal
            Call SetForegroundWindow(Me.hWnd)
            Me.Show

        Case WM_RBUTTONUP
            Call SetForegroundWindow(Me.hWnd)
            ' -- Si hacemos Click con el boton derecho mostramos el popup Menu
            Me.PopupMenu Me.MnuPopUp

        Case WM_LBUTTONUP
    End Select
End Sub

Private Sub RemoverSystray()
    Shell_NotifyIcon NIM_DELETE, systray
End Sub

Private Sub Form_Resize()
    If (Me.WindowState = vbMinimized) Then
        Me.Hide
        Call PonerSystray
    Else
        Call RemoverSystray
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' -- cuando descargamos el form removemos el Icono del systray
    RemoverSystray
End Sub

Private Sub mnuAcerca_Click()
    ShellAbout Me.hWnd, App.Title, "Programa para ...", ByVal 0&
End Sub

Private Sub MnuMinimizar_Click()
Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIcon(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False
End Sub

Private Sub mnuMostrar_Click()
    Call RemoverSystray
    Me.Show
End Sub

Private Sub mnuSalir_Click()
    If MsgBox("¿Seguro que deseas salir?", vbOKCancel) = vbOK Then Unload Me
End Sub

Private Sub PonerSystray()
    
    With systray
        ' -- Tamaño de la estructura systray
        .cbSize = Len(systray)
        ' -- Establecemos el Hwnd, en este caso del formulario
        .hWnd = Me.hWnd

        .uId = vbNull
        ' -- Flags
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        ' -- Establecemos el mensaje callback
        .ucallbackMessage = WM_MOUSEMOVE
        ' -- establecemos el icono, en este caso el que tiene el form, puede ser otro
        .hIcon = Me.Icon
        ' -- Establecemos el tooltiptext
        .szTip = Me.Caption & vbNullChar
        ' -- Ponemos el icono en el systray
        Shell_NotifyIcon NIM_ADD, systray
    End With

End Sub

Private Sub Command1_Click()
WindowsMediaPlayer1.URL = "http://84.122.32.37:8000/"
Label1.Caption = "RadioXtreme..::..La Mejor Música..::..Play"
End Sub
Private Sub Command2_Click()
WindowsMediaPlayer1.Close
Label1.Caption = "RadioXtreme..::..La Mejor Música..::..Stop"
End Sub

Private Sub Form_Load()
Label1.Caption = "RadioXtreme..::..La Mejor Música..::..Stop"
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If
Winsock1.Connect "84.122.32.37", 8000 ' <= direccion y puerto del server
End Sub
Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uId = ID
    nidTemp.uFlags = flags
    nidTemp.ucallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Private Sub Winsock1_Connect()
Label3.ForeColor = vbGreen
Label3.Caption = "Radio Online"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label3.ForeColor = vbRed
Label3.Caption = "Radio Offline"
End Sub
