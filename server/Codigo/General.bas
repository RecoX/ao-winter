Attribute VB_Name = "General"
Option Explicit

Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'***************************************************
Dim CuerpoDesnudo As Integer
Select Case UserList(UserIndex).genero
    Case eGenero.Hombre
        Select Case UserList(UserIndex).raza
            Case eRaza.Humano
                CuerpoDesnudo = 21
            Case eRaza.Drow
                CuerpoDesnudo = 32
            Case eRaza.Elfo
                CuerpoDesnudo = 210
            Case eRaza.Gnomo
                CuerpoDesnudo = 222
            Case eRaza.Enano
                CuerpoDesnudo = 53
            Case eRaza.orco
                CuerpoDesnudo = 499
        End Select
    Case eGenero.Mujer
        Select Case UserList(UserIndex).raza
            Case eRaza.Humano
                CuerpoDesnudo = 39
            Case eRaza.Drow
                CuerpoDesnudo = 40
            Case eRaza.Elfo
                CuerpoDesnudo = 259
            Case eRaza.Gnomo
                CuerpoDesnudo = 260
            Case eRaza.Enano
                CuerpoDesnudo = 60
            Case eRaza.orco
                CuerpoDesnudo = 501
        End Select
End Select

If Mimetizado Then
    UserList(UserIndex).CharMimetizado.body = CuerpoDesnudo
Else
    UserList(UserIndex).Char.body = CuerpoDesnudo
End If

UserList(UserIndex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal B As Boolean)
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s

If toMap Then
    Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, B))
Else
    Call WriteBlockPosition(sndIndex, X, Y, B)
End If

End Sub


Function HayAgua(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If ((MapData(map, X, Y).Graphic(1) >= 1505 And MapData(map, X, Y).Graphic(1) <= 1520) Or _
    (MapData(map, X, Y).Graphic(1) >= 5665 And MapData(map, X, Y).Graphic(1) <= 5680) Or _
    (MapData(map, X, Y).Graphic(1) >= 13547 And MapData(map, X, Y).Graphic(1) <= 13562)) And _
       MapData(map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Private Function HayLava(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
If map > 0 And map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(map, X, Y).Graphic(1) >= 5837 And MapData(map, X, Y).Graphic(1) <= 5852 Then
        HayLava = True
    Else
        HayLava = False
    End If
Else
  HayLava = False
End If

End Function

Sub LimpiarMundo()
On Error GoTo Errhandler
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando Mundo.", FontTypeNames.FONTTYPE_SERVER))
Dim MapaActual As Long
Dim Y As Long
Dim X As Long

For MapaActual = 1 To NumMaps
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(MapaActual, X, Y).ObjInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then
                If ItemNoEsDeMapa(MapData(MapaActual, X, Y).ObjInfo.ObjIndex) Then Call EraseObj(10000, MapaActual, X, Y)
            End If
        Next X
    Next Y
Next MapaActual
 
LimpiezaTimerMinutos = TimerCleanWorld
 
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo realizada.", FontTypeNames.FONTTYPE_SERVER))

Errhandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.description)
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Long
Dim npcNames() As String

ReDim npcNames(1 To UBound(SpawnList)) As String

For k = 1 To UBound(SpawnList)
    npcNames(k) = SpawnList(k).NpcName
Next k

Call WriteSpawnList(UserIndex, npcNames())

End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub

Sub Main()
On Error Resume Next
Dim f As Date

ChDir App.Path
ChDrive App.Path

Call LoadMotd
Call BanIpCargar

Prision.map = 56
Libertad.map = 56

Prision.X = 44
Prision.Y = 52
Libertad.X = 43
Libertad.Y = 57
Denuncias = True

LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"



LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100


ListaRazas(eRaza.Humano) = "Humano"
ListaRazas(eRaza.Elfo) = "Elfo"
ListaRazas(eRaza.Drow) = "Drow"
ListaRazas(eRaza.Gnomo) = "Gnomo"
ListaRazas(eRaza.Enano) = "Enano"
ListaRazas(eRaza.orco) = "Orco"

ListaClases(eClass.Mage) = "Mago"
ListaClases(eClass.Cleric) = "Clerigo"
ListaClases(eClass.Warrior) = "Guerrero"
ListaClases(eClass.Assasin) = "Asesino"
ListaClases(eClass.Thief) = "Ladron"
ListaClases(eClass.Bard) = "Bardo"
ListaClases(eClass.Druid) = "Druida"
ListaClases(eClass.Bandit) = "Bandido"
ListaClases(eClass.Paladin) = "Paladin"
ListaClases(eClass.Hunter) = "Cazador"
ListaClases(eClass.trabajador) = "Trabajador"
ListaClases(eClass.Pirat) = "Pirata"

SkillsNames(eSkill.Equitacion) = "Equitacion"
SkillsNames(eSkill.Magia) = "Magia"
SkillsNames(eSkill.Robar) = "Robar"
SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
SkillsNames(eSkill.Armas) = "Combate con armas"
SkillsNames(eSkill.Meditar) = "Meditar"
SkillsNames(eSkill.Apu�alar) = "Apu�alar"
SkillsNames(eSkill.Ocultarse) = "Ocultarse"
SkillsNames(eSkill.Supervivencia) = "Supervivencia"
SkillsNames(eSkill.Talar) = "Talar arboles"
SkillsNames(eSkill.Comerciar) = "Comercio"
SkillsNames(eSkill.Defensa) = "Defensa con escudos"
SkillsNames(eSkill.Pesca) = "Pesca"
SkillsNames(eSkill.Mineria) = "Mineria"
SkillsNames(eSkill.Carpinteria) = "Carpinteria"
SkillsNames(eSkill.Herreria) = "Herreria"
SkillsNames(eSkill.Liderazgo) = "Liderazgo"
SkillsNames(eSkill.Domar) = "Domar animales"
SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
SkillsNames(eSkill.Wrestling) = "Wrestling"
SkillsNames(eSkill.Navegacion) = "Navegacion"

ListaAtributos(eAtributos.Fuerza) = "Fuerza"
ListaAtributos(eAtributos.Agilidad) = "Agilidad"
ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
ListaAtributos(eAtributos.Carisma) = "Carisma"
ListaAtributos(eAtributos.Constitucion) = "Constitucion"


frmCargando.Show

frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB

Call LoadQuests

Call CargarSpawnList
Call CargarForbidenWords
'�?�?�?�?�?�?�?� CARGAMOS DATOS DESDE ARCHIVOS �??�?�?�?�?�?�?�
frmCargando.Label1(2).Caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call CargaApuestas

'*************************************************
frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
Call CargaNpcsDat
'*************************************************

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
'Call LoadOBJData
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
frmCargando.Label1(2).Caption = "Cargando Objetos de Herrer�a"
Call LoadArmasHerreria
Call LoadArmadurasHerreria

frmCargando.Label1(2).Caption = "Cargando Objetos de Carpinter�a"
Call LoadObjCarpintero

frmCargando.Label1(2).Caption = "Cargando Objetos de Canjes"
Call LoadCanjes 'Lorwik

frmCargando.Label1(2).Caption = "Cargando Clima"
Call SortearClima 'Lorwik

frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
Call LoadBalance    '4/01/08 Pablo ToxicWaste

If BootDelBackUp Then
    
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

With frmMain
    .AutoSave.Enabled = True
    .tPiqueteC.Enabled = True
    .GameTimer.Enabled = True
    .Auditoria.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
    .EventoHora.Enabled = True
    .lblLimpieza = "Limpieza del mundo en: " & LimpiezaTimerMinutos & " minutos."
End With

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�




Unload frmCargando


'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #N

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF

End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = LenB(dir$(File, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************

    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function MapaValido(ByVal map As Integer) As Boolean
'***************************************************
MapaValido = map >= 1 And map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = NumUsers

End Sub


Public Sub LogCriticEvent(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

Errhandler:

End Sub


Public Sub LogIndex(ByVal index As Integer, ByVal desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

Errhandler:

End Sub


Public Sub LogError(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogStatic(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogTarea(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & desc
Close #nfile

Exit Sub

Errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub

Public Sub LogGM(Nombre As String, texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogGlobal(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Global.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogAsesinato(texto As String)
On Error GoTo Errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

Errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

Errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

Errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

Errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If LenB(Arg) = 0 Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Long
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

'Initialize statistics!!
Call Statistics.Initialize

For LoopC = 1 To UBound(UserList())
    Set UserList(LoopC).incomingData = Nothing
    Set UserList(LoopC).outgoingData = Nothing
Next LoopC

ReDim UserList(1 To MaxUsers) As User

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

LastUser = 0
NumUsers = 0

Call FreeNPCs
Call FreeCharIndexes

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " servidor reiniciado."
Close #N

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    
    If MapInfo(UserList(UserIndex).Pos.map).Zona <> "DUNGEON" Then
        If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And _
           MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And _
           MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
    
    Dim modifi As Integer
    
    If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
        UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
    Else
        If MapInfo(UserList(UserIndex).Pos.map).Terreno = Nieve Then
            Call WriteConsoleMsg(UserIndex, "��Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
            
            If UserList(UserIndex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(UserIndex, "��Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)
            End If
            
            Call WriteUpdateHP(UserIndex)
        Else
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
            Call QuitarSta(UserIndex, modifi)
            Call WriteUpdateSta(UserIndex)
        End If
        
        UserList(UserIndex).Counters.Frio = 0
    End If
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'If user is standing on lava, take health points from him
'***************************************************
    If UserList(UserIndex).Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
        UserList(UserIndex).Counters.Lava = UserList(UserIndex).Counters.Lava + 1
    Else
        If HayLava(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
            Call WriteConsoleMsg(UserIndex, "��Quitate de la lava, te est�s quemando!!.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
            
            If UserList(UserIndex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(UserIndex, "��Has muerto quemado!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)
            End If
            
            Call WriteUpdateHP(UserIndex)
        End If
        
        UserList(UserIndex).Counters.Lava = 0
    End If
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
'******************************************************
'Author: Unknown
'Last Update: 04/11/2008 (NicoNZ)
'
'******************************************************
    Dim Barco As ObjData
    
    With UserList(UserIndex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    If .Faccion.ArmadaReal = 1 Then
                        .Char.body = iFragataReal
                    ElseIf .Faccion.FuerzasCaos = 1 Then
                        .Char.body = iFragataCaos
                    Else
                        Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
                        If criminal(UserIndex) Then
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaPk
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraPk
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonPk
                        Else
                            If Barco.Ropaje = iBarca Then .Char.body = iBarcaCiuda
                            If Barco.Ropaje = iGalera Then .Char.body = iGaleraCiuda
                            If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCiuda
                        End If
                    End If
                Else
                    .Char.body = iFragataFantasmal
                End If
                
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            Else
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Char.Aura = .CharMimetizado.Aura
            End If
            
            With .Char
                Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .Aura)
            End With
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
Dim TiempoTranscurrido As Long
If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
     TiempoTranscurrido = (UserList(UserIndex).Counters.Invisibilidad * frmMain.GameTimer.Interval)
     
    If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
        If TiempoTranscurrido = 40 Then
             Call WriteTimeInvi(UserIndex, ((IntervaloInvisible * frmMain.GameTimer.Interval) / 1000))
         Else
             Call WriteTimeInvi(UserIndex, (((IntervaloInvisible * frmMain.GameTimer.Interval) / 1000) - (TiempoTranscurrido / 1000)))
        End If
    End If
Else
    UserList(UserIndex).Counters.Invisibilidad = 0
    UserList(UserIndex).flags.invisible = 0
    Call WriteTimeInvi(UserIndex, 0)
    If UserList(UserIndex).flags.Oculto = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
End If

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Ceguera > 0 Then
    UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
Else
    If UserList(UserIndex).flags.Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 0
        Call WriteBlindNoMore(UserIndex)
    End If
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
Dim TiempoTranscurrido As Long

If UserList(UserIndex).Counters.Paralisis > 0 Then
    UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
    TiempoTranscurrido = (IntervaloParalizado * frmMain.GameTimer.Interval) - (UserList(UserIndex).Counters.Paralisis * frmMain.GameTimer.Interval)
   
    If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
        If TiempoTranscurrido = 40 Then
            Call WriteTimeParal(UserIndex, ((IntervaloParalizado * frmMain.GameTimer.Interval) / 1000))
        Else
            Call WriteTimeParal(UserIndex, ((IntervaloParalizado * frmMain.GameTimer.Interval) / 1000) - (TiempoTranscurrido / 1000))
        End If
    End If
Else
    UserList(UserIndex).flags.Paralizado = 0
    UserList(UserIndex).flags.Inmovilizado = 0
    Call WriteTimeParal(UserIndex, 0)
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call WriteParalizeOK(UserIndex)
End If

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub


Dim massta As Integer
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
    If UserList(UserIndex).Counters.STACounter < Intervalo Then
        UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
    Else
        EnviarStats = True
        UserList(UserIndex).Counters.STACounter = 0
        If UserList(UserIndex).flags.Desnudo Or _
        UserList(UserIndex).flags.Makro <> 0 Then Exit Sub 'Desnudo y trabajando no sube energia. (ToxicWaste)
       
        massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta
        If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
Dim N As Integer

If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
  UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
Else
  Call WriteConsoleMsg(UserIndex, "Est�s envenenado, si no te curas morir�s.", FontTypeNames.FONTTYPE_VENENO)
  UserList(UserIndex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  Call WriteUpdateHP(UserIndex)
End If

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

'Controla la duracion de las pociones
If UserList(UserIndex).flags.DuracionEfecto > 0 Then
   UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1
   If UserList(UserIndex).flags.DuracionEfecto = 0 Then
        UserList(UserIndex).flags.TomoPocion = False
        UserList(UserIndex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
        Next
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
   End If
End If

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)

If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Exit Sub

'Sed
If UserList(UserIndex).Stats.MinAGU > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
        UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
    Else
        UserList(UserIndex).Counters.AGUACounter = 0
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
        If UserList(UserIndex).Stats.MinAGU <= 0 Then
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).flags.Sed = 1
        End If
        
        fenviarAyS = True
    End If
End If

'hambre
If UserList(UserIndex).Stats.MinHam > 0 Then
   If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
        UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
   Else
        UserList(UserIndex).Counters.COMCounter = 0
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10
        If UserList(UserIndex).Stats.MinHam <= 0 Then
               UserList(UserIndex).Stats.MinHam = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
    If UserList(UserIndex).Counters.HPCounter < Intervalo Then
        UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
    Else
        mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        
        UserList(UserIndex).Counters.HPCounter = 0
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
        EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Call LeerNPCs.Initialize(npcfile)
End Sub

Sub PasarSegundo()
'***************************************************
'Author: Unknown
'Last Modification: 25/09/2011
'Lorwik> Le agrege una cuenta regresiva al salir.
'***************************************************

On Error GoTo Errhandler
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                Call WriteConsoleMsg(i, "Cerrando en " & UserList(i).Counters.Salir - 1, FontTypeNames.FONTTYPE_INFO)
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                If UserList(i).Counters.Salir <= 0 Then
                    Call WriteConsoleMsg(i, "Gracias por Winter-AO", FontTypeNames.FONTTYPE_INFO)
                    Call WriteDisconnect(i)
                    Call FlushBuffer(i)
                    
                    Call CloseSocket(i)
                End If
            End If
        End If
    Next i
Exit Sub

Errhandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False
End Sub

Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub
Function ZonaCura(ByVal UserIndex As Integer) As Boolean
Dim X As Integer, Y As Integer
For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
        For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
       
            If MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex).NPCtype = 1 Then
                    If Distancia(UserList(UserIndex).Pos, Npclist(MapData(UserList(UserIndex).Pos.map, X, Y).NpcIndex).Pos) < 10 Then
                        ZonaCura = True
                        Exit Function
                    End If
                End If
            End If
           
        Next X
Next Y
ZonaCura = False
End Function

Public Function Tilde(data As String) As String
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "�", "A"), "�", "E"), "�", "I"), "�", "O"), "�", "U")
End Function

Public Sub Canjeos(ByVal UserIndex As Integer, ByVal index As Integer)

    If UserList(UserIndex).Stats.Puntos < ObjData(ObjCanjes(index)).Puntos Then
        Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos!.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserIndex).Stats.Puntos = UserList(UserIndex).Stats.Puntos - ObjData(ObjCanjes(index)).Puntos
    Call WriteConsoleMsg(UserIndex, "Se te han descontado " & ObjData(ObjCanjes(index)).Puntos & " Pts. de Torneo y ahora tienes " & ObjData(ObjCanjes(index)).Name & "!.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(UserIndex, "Tus puntos actuales son " & UserList(UserIndex).Stats.Puntos, FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.amount = 1
    MiObj.ObjIndex = ObjCanjes(index)
    
    Call MeterItemEnInventario(UserIndex, MiObj)
    Call WriteCanjesInit(UserIndex)
    
End Sub

'******************************************************************************
'Asigna la experiencia ganada al usuario
'Cre� esta funci�n para optimizar el sistema de eventos EXP * 2
'[MaxTus]
'******************************************************************************
Public Sub AsignarExperiencia(ByVal UserIndex As Integer, ByVal Experiencia As Long)

    If HappyHourAC = True Then
        Experiencia = Experiencia * 2
    End If
    
    With UserList(UserIndex)
        .Stats.Exp = .Stats.Exp + Experiencia
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        Call CheckUserLevel(UserIndex)
        Call WriteUpdateExp(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has ganado " & Experiencia & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)
    End With
 
End Sub

Public Function DentroDelRango(ByVal UserIndex As Integer, ByVal NpcNumber As Integer) As Boolean
'******************************************************************************
'Determina si el usuario est� en el rango de visi�n de X npc
'[MaxTus]
'******************************************************************************
On Error GoTo ErrorHandler

Dim X As Integer, Y As Integer

    With UserList(UserIndex)
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If MapData(.Pos.map, X, Y).NpcIndex > 0 Then
                    If Npclist(MapData(.Pos.map, X, Y).NpcIndex).Numero = NpcNumber Then
                        If Distancia(.Pos, Npclist(MapData(.Pos.map, X, Y).NpcIndex).Pos) < 10 Then
                            DentroDelRango = True
                            'La unica forma que se me ocurri� para almacenar el NPCIndex, sino
                            'iba a tener que devolver 2 par�metros :S
                            .flags.TargetNPC = MapData(.Pos.map, X, Y).NpcIndex
                            Exit Function
                        End If
                    End If
                End If
            Next X
        Next Y
        DentroDelRango = False
    End With
    
Exit Function
ErrorHandler:
    Call LogError("Error en DentroDelRango: " & Err.description)
End Function

Public Sub PostMensaje(ByVal Mensaje As String)
    frmMain.ListadoM.AddItem Date & " " & Time & " " & Mensaje
    If frmMain.ListadoM.ListCount > 9 Then frmMain.ListadoM.RemoveItem 0
    frmMain.ListadoM.ListIndex = frmMain.ListadoM.ListCount - 1
End Sub
