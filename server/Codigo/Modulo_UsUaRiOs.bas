Attribute VB_Name = "UsUaRiOs"
Option Explicit

'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo Usuarios
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Rutinas de los usuarios
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)
    Dim DaExp As Integer
    Dim EraCriminal As Boolean
    
    DaExp = CInt(UserList(VictimIndex).Stats.ELV) * 2
    
    With UserList(attackerIndex)
        
        Call AsignarExperiencia(attackerIndex, DaExp)
        Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimIndex, "�" & .Name & " te ha matado!", FontTypeNames.FONTTYPE_INFO)
        
        If TriggerZonaPelea(VictimIndex, attackerIndex) <> TRIGGER6_PERMITE Then
            EraCriminal = criminal(attackerIndex)
            
            With .Reputacion
                If Not criminal(VictimIndex) Then
                    .AsesinoRep = .AsesinoRep + vlASESINO * 2
                    If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                    .BurguesRep = 0
                    .NobleRep = 0
                    .PlebeRep = 0
                Else
                    .NobleRep = .NobleRep + vlNoble
                    If .NobleRep > MAXREP Then .NobleRep = MAXREP
                End If
            End With
            
            If criminal(attackerIndex) Then
                If Not EraCriminal Then Call RefreshCharStatus(attackerIndex)
            Else
                If EraCriminal Then Call RefreshCharStatus(attackerIndex)
            End If
        End If
        
        'Call UserDie(VictimIndex)
        
        Call FlushBuffer(VictimIndex)
        
        'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(VictimIndex).Name)
    End With
End Sub

Public Sub RevivirUsuario(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .flags.Muerto = 0
        .Stats.MinHP = .Stats.UserAtributos(eAtributos.Constitucion)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(UserList(UserIndex).Char.CharIndex, 0, 0))
        
        If .Stats.MinHP > .Stats.MaxHP Then
            .Stats.MinHP = .Stats.MaxHP
        End If
        
        If .flags.Navegando = 1 Then
            Dim Barco As ObjData
            Barco = ObjData(.Invent.BarcoObjIndex)
            .Char.Head = 0
            
            If .Faccion.ArmadaReal = 1 Then
                .Char.body = iFragataReal
            ElseIf .Faccion.FuerzasCaos = 1 Then
                .Char.body = iFragataCaos
            Else
                If criminal(UserIndex) Then
                    Select Case Barco.Ropaje
                        Case iBarca
                            .Char.body = iBarcaPk
                        
                        Case iGalera
                            .Char.body = iGaleraPk
                        
                        Case iGaleon
                            .Char.body = iGaleonPk
                    End Select
                Else
                    Select Case Barco.Ropaje
                        Case iBarca
                            .Char.body = iBarcaCiuda
                        
                        Case iGalera
                            .Char.body = iGaleraCiuda
                        
                        Case iGaleon
                            .Char.body = iGaleonCiuda
                    End Select
                End If
            End If
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.Aura = 0
        Else
            Call DarCuerpoDesnudo(UserIndex)
            
            .Char.Head = .OrigChar.Head
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.Aura)
        Call WriteUpdateUserStats(UserIndex)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 3, 1))
        .Char.Particle = 0
        .Char.loops = 0
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(20, .Pos.X, .Pos.Y))
        
        Call WriteConsoleMsg(UserIndex, "�Has resucitado!", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(250, .Pos.X, .Pos.Y))
    End With
End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer, ByVal Aura As Integer)

    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        .Aura = Aura
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .FX, .loops, casco, Aura))
    End With
End Sub
Public Function GetWeaponAnim(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/29/10
'
'***************************************************
    Dim Tmp As Integer

    With UserList(UserIndex)
        Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
        If Tmp > 0 Then
            If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
                GetWeaponAnim = Tmp
                Exit Function
            End If
        End If
        
        GetWeaponAnim = ObjData(ObjIndex).WeaponAnim
    End With
End Function

Sub EnviarFama(ByVal UserIndex As Integer)
    Dim L As Long
    
    With UserList(UserIndex).Reputacion
        L = (-.AsesinoRep) + _
            (-.BandidoRep) + _
            .BurguesRep + _
            (-.LadronesRep) + _
            .NobleRep + _
            .PlebeRep
        L = Round(L / 6)
        
        .Promedio = L
    End With
    
    Call WriteFame(UserIndex)
End Sub

Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)
'*************************************************
'Author: Unknown
'Last modified: 08/01/2009
'08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
'*************************************************

On Error GoTo ErrorHandler
    
    With UserList(UserIndex)
        CharList(.Char.CharIndex) = 0
        
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If
        
        ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
        If IsAdminInvisible Then
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            'Le mandamos el mensaje para que borre el personaje a los clientes que est�n cerca
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        End If
        
        Call QuitarUser(UserIndex, .Pos.map)
        
        MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description)
End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/07/2009
'Refreshes the status and tag of UserIndex.
'04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si Est�s muerto.
'*************************************************
    Dim klan As String
    Dim Barco As ObjData
    Dim esCriminal As Boolean
    
    With UserList(UserIndex)
        If .GuildIndex > 0 Then
            klan = modGuilds.GuildName(.GuildIndex)
            klan = " <" & klan & ">"
        End If
        
        esCriminal = criminal(UserIndex)
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, esCriminal, .Name & klan))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, esCriminal, vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = iFragataFantasmal
            Else
                Barco = ObjData(.Invent.Object(.Invent.BarcoSlot).ObjIndex)
                
                If .Faccion.ArmadaReal = 1 Then
                    .Char.body = iFragataReal
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    .Char.body = iFragataCaos
                Else
                    If esCriminal Then
                        Select Case Barco.Ropaje
                            Case iBarca
                                .Char.body = iBarcaPk
                            
                            Case iGalera
                                .Char.body = iGaleraPk
                            
                            Case iGaleon
                                .Char.body = iGaleonPk
                        End Select
                    Else
                        Select Case Barco.Ropaje
                            Case iBarca
                                .Char.body = iBarcaCiuda
                            
                            Case iGalera
                                .Char.body = iGaleraCiuda
                            
                            Case iGaleon
                                .Char.body = iGaleonCiuda
                        End Select
                    End If
                End If
            End If
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.Aura)
        End If
    End With
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/07/2009
'
'23/07/2009: Budi - Ahora se env�a el nick
'*************************************************

On Error GoTo hayerror
    Dim CharIndex As Integer
    
    With UserList(UserIndex)
    
        If InMapBounds(map, X, Y) Then
            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = UserIndex
            End If
            
            'Place character on map if needed
            If toMap Then MapData(map, X, Y).UserIndex = UserIndex
            
            'Send make character command to clients
            Dim klan As String
            If .GuildIndex > 0 Then
                klan = modGuilds.GuildName(.GuildIndex)
            End If
            
            Dim bCr As Byte
            Dim bNick As String
            Dim bPriv As Byte
            
            bCr = criminal(UserIndex)
            bPriv = .flags.Privilegios
            'Preparo el nick
            If .showName Then
                If UserList(sndIndex).flags.Privilegios And PlayerType.User Then
                    If LenB(klan) <> 0 Then
                        bNick = .Name & " <" & klan & ">"
                    Else
                        bNick = .Name
                    End If
'                    bPriv = .flags.Privilegios
                Else
                    If .flags.invisible Or .flags.Oculto Then
                        bNick = .Name & " " & TAG_USER_INVISIBLE
                    Else
                        If LenB(klan) <> 0 Then
                            bNick = .Name & " <" & klan & ">"
                        Else
                            bNick = .Name
                        End If
                    End If
'                    bPriv = .flags.Privilegios
                End If
            Else
                bNick = vbNullString
'                bPriv = PlayerType.User
            End If
            
            If Not toMap Then
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, _
                            .Char.CharIndex, X, Y, _
                            .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, _
                            bNick, bCr, bPriv, .Char.Aura)
            Else
                'Hide the name and clan - set privs as normal user
                 Call AgregarUser(UserIndex, .Pos.map)
            End If
            
        End If
    End With
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call CloseSocket(UserIndex)
End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 11/19/2009
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'07/08/2006 Integer - Modificacion de los valores
'01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
'13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constituci�n.
'09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consituci�n se controla desde Balance.dat
'12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y est� en un clan, se lo expulsa para no sumar antifacci�n
'02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
'11/19/2009 Pato - Modifico la nueva f�rmula de man� ganada para el bandido y se la limito a 499
'02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
'*************************************************
    Dim Pts As Integer
    Dim AumentoHIT As Integer
    Dim AumentoMANA As Integer
    Dim AumentoSTA As Integer
    Dim AumentoHP As Integer
    Dim WasNewbie As Boolean
    Dim Promedio As Double
    Dim aux As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI As Integer 'Guild Index
    
On Error GoTo Errhandler
    
    WasNewbie = EsNewbie(UserIndex)
    
    With UserList(UserIndex)
        Do While .Stats.Exp >= .Stats.ELU
            
            'Checkea si alcanz� el m�ximo nivel
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 1
                Exit Sub
            End If
            
            'Store it!
            Call Statistics.UserLevelUp(UserIndex)
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
            Call WriteConsoleMsg(UserIndex, "�Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
            Call WriteChatOverHeadHit(UserIndex, "Level Up!", UserList(UserIndex).Char.CharIndex, 0, 255, 0)
            
            If .Stats.ELV = 1 Then
                Pts = 10
            Else
                'For multiple levels being rised at once
                Pts = Pts + 5
            End If
            
            .Stats.ELV = .Stats.ELV + 1
            
            .Stats.Exp = .Stats.Exp - .Stats.ELU
            
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
            If .Stats.ELV < 15 Then
                .Stats.ELU = .Stats.ELU * 1.4
            ElseIf .Stats.ELV < 21 Then
                .Stats.ELU = .Stats.ELU * 1.35
            ElseIf .Stats.ELV < 33 Then
                .Stats.ELU = .Stats.ELU * 1.3
            ElseIf .Stats.ELV < 41 Then
                .Stats.ELU = .Stats.ELU * 1.225
            ElseIf .Stats.ELV = 50 Then
                .Stats.ELU = 1
                .Stats.Exp = 0
            Else
                .Stats.ELU = .Stats.ELU * 1.25
            End If
            
            'Calculo subida de vida
            Promedio = ModVida(.clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
            aux = RandomNumber(0, 100)
            
        
            If Promedio - Int(Promedio) = 0.5 Then
                'Es promedio semientero
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 1.5
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio - 0.5
                Else
                    AumentoHP = Promedio - 1.5
                End If
            Else
                'Es promedio entero
                
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
                DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 2
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 1
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio
                ElseIf aux <= DistVida(4) Then
                    AumentoHP = Promedio - 1
                Else
                    AumentoHP = Promedio - 2
                End If
                
            End If
            If WasNewbie Then
                .Stats.GLD = .Stats.GLD + 500
                Call WriteConsoleMsg(UserIndex, "Has ganado 500 monedas de oro.", FontTypeNames.FONTTYPE_GLD)
            End If
        
            Select Case .clase
                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Pirat
                    AumentoHIT = 3
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Thief
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mage
                    AumentoHIT = 1
                    AumentoMANA = 2.8 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                
                Case eClass.trabajador
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTTrabajador
                
                Case eClass.Cleric
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druid
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bard
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                    
                Case eClass.Bandit
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3 * 2
                    AumentoSTA = AumentoStBandido
                
                Case Else
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHP = .Stats.MaxHP + AumentoHP
            If .Stats.MaxHP > STAT_MAXHP Then .Stats.MaxHP = STAT_MAXHP
            
            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
            
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            

            'Actualizamos Golpe M�ximo
            .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MaxHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
                    .Stats.MaxHIT = STAT_MAXHIT_OVER36
            End If
            
            'Actualizamos Golpe M�nimo
            .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MinHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
                    .Stats.MinHIT = STAT_MAXHIT_OVER36
            End If
            
            'Notificamos al user
            If AumentoHP > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoSTA > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de energ�a.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoMANA > 0 Then
                Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de man�.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoHIT > 0 Then
                Call WriteConsoleMsg(UserIndex, "Tu golpe m�ximo aument� en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Tu golpe m�nimo aument� en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
            
            .Stats.MinHP = .Stats.MaxHP

                'If user is in a party, we modify the variable p_sumaniveleselevados
                Call mdParty.ActualizarSumaNivelesElevados(UserIndex)
                    'If user reaches lvl 25 and he is in a guild, we check the guild's alignment and expulses the user if guild has factionary alignment
        
            If .Stats.ELV = 25 Then
                GI = .GuildIndex
                If GI > 0 Then
                    If modGuilds.GuildAlignment(GI) = "Legi�n Oscura" Or modGuilds.GuildAlignment(GI) = "Caballeros del Temple" Then
                        'We get here, so guild has factionary alignment, we have to expulse the user
                        Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                        Call WriteConsoleMsg(UserIndex, "�Ya tienes la madurez suficiente como para decidir bajo que estandarte pelear�s! Por esta raz�n, hasta tanto no te enlistes en la facci�n bajo la cual tu clan est� alineado, estar�s exclu�do del mismo.", FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If
            End If

        Loop
        
        'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
        If Not EsNewbie(UserIndex) And WasNewbie Then
            Call QuitarNewbieObj(UserIndex)
            If UCase$(MapInfo(.Pos.map).Restringir) = "NEWBIE" Then
                Call WarpUserChar(UserIndex, 1, 50, 50, True)
                Call WriteConsoleMsg(UserIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'Send all gained skill points at once (if any)
        If Pts > 0 Then
            Call WriteLevelUp(UserIndex, Pts)
            
            .Stats.SkillPts = .Stats.SkillPts + Pts
            
            Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
    
    Call WriteUpdateUserStats(UserIndex)
Exit Sub

Errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)

End Sub
Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 _
                    Or UserList(UserIndex).flags.Vuela = 1
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)
'*************************************************
'Author: Unknown
'Last modified: 13/07/2009
'Moves the char, sending the message to everyone in range.
'30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
'28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
'13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
'13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
'*************************************************
    Dim nPos As WorldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim isAdminInvi As Boolean
    
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
        
    isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)
    
    If MoveToLegalPos(UserList(UserIndex).Pos.map, nPos.X, nPos.Y, sailing, Not sailing) Then
        'si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.map).NumUsers > 1 Then
               
            CasperIndex = MapData(UserList(UserIndex).Pos.map, nPos.X, nPos.Y).UserIndex
            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then
                ' Los admins invisibles no pueden patear caspers
                If Not isAdminInvi Then
    
                    With UserList(CasperIndex)
                        CasperHeading = InvertHeading(nHeading)
                        Call HeadtoPos(CasperHeading, .Pos)
                    
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not .flags.AdminInvisible = 1 Then _
                            Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y))
                        
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                            
                        'Update map and char
                        .Char.heading = CasperHeading
                        MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = CasperIndex
                    End With
                
                    'Actualizamos las �reas de ser necesario
                    Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                End If
            End If
            
            ' Si es un admin invisible, no se avisa a los demas clientes
            If Not isAdminInvi Then _
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.X, nPos.Y))
            
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If Not (isAdminInvi And (CasperIndex <> 0)) Then
            Dim oldUserIndex As Integer
            
            With UserList(UserIndex)
                oldUserIndex = MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex
                
                ' Si no hay intercambio de pos con nadie
                If oldUserIndex = UserIndex Then
                    MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
                End If
                
                .Pos = nPos
                .Char.heading = nHeading
                MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                
                'Si estoy cerca del sacerdote... [MaxTus]
                If DentroDelRango(UserIndex, 5) = True Then
                    AutoCuraUser UserIndex, .flags.TargetNPC
                End If
                
                Call DoTileEvents(UserIndex, .Pos.map, .Pos.X, .Pos.Y)
            End With
            
            'Actualizamos las �reas de ser necesario
            Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
        Else
            Call WritePosUpdate(UserIndex)
        End If
    Else
        Call WritePosUpdate(UserIndex)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
End Sub

Public Sub AutoCuraUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

If UserList(UserIndex).flags.Muerto = 1 Then
    NpcLanzaSpellSobreUser NpcIndex, UserIndex, 11
End If
 
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP And UserList(UserIndex).flags.Muerto <> 1 Then
    NpcLanzaSpellSobreUser NpcIndex, UserIndex, 48
End If
 
If UserList(UserIndex).flags.Envenenado = 1 Then UserList(UserIndex).flags.Envenenado = 0
 
End Sub
Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    Dim GuildI As Integer
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHP & "/" & .Stats.MaxHP & "  Mana: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Vitalidad: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        GuildI = .GuildIndex
        If GuildI > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)
            End If
            'guildpts no tienen objeto
        End If
        
#If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.GLD & "  Posicion: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Atributos: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Puntos: " & .Stats.Puntos, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribuci�n de par�metros.
'*************************************************
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Caballeros del Temple Desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingres� en Nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " Ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejercito Oscuro Desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingres� en Nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Caballeros del Temple", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Legionario", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        
        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is offline.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribuci�n de par�metros.
'*************************************************
    Dim CharFile As String
    Dim Ban As String
    Dim BanDetailPath As String
    
    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Caballeros del Temple Desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingres� en Nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " Ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejercito Oscuro Desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingres� en Nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Caballeros del Temple", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue Legionario", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que Ingres�: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        End If

        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next

    Dim j As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            If .Invent.Object(j).ObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).amount, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    End With
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
    Dim j As Integer
    
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

    If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "��" & UserList(UserIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 06/28/2008
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran m�s al lado de �l sin hacer Ninguno.
'**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then
        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
        End If
    End If
    
    If EsMascotaCiudadano(NpcIndex, UserIndex) Then
        Call VolverCriminal(UserIndex)
        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1
    Else
        EraCriminal = criminal(UserIndex)
        
        'Reputacion
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
           If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                Call VolverCriminal(UserIndex)
           Else
                If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                    Call VolverCriminal(UserIndex)
                End If
           End If
        
        ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
           UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2
           If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
            UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        End If
        
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            'hacemos que el npc se defienda
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
        End If
        
        If EraCriminal And Not criminal(UserIndex) Then
            Call VolverCiudadano(UserIndex)
        End If
    End If
End Sub
Public Function PuedeApu�alar(ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apu�ala = 1 Then
            PuedeApu�alar = UserList(UserIndex).Stats.UserSkills(eSkill.Apu�alar) >= MIN_APU�ALAR _
                        Or UserList(UserIndex).clase = eClass.Assasin
        End If
    End If
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

    With UserList(UserIndex)
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            
            If .Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
  
            If RandomNumber(1, 15) = 7 Then
                .Stats.UserSkills(Skill) = .Stats.UserSkills(Skill) + 1
                Call WriteConsoleMsg(UserIndex, "�Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & .Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                Call AsignarExperiencia(UserIndex, 50)
            End If
        End If
    End With
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 21/07/2009
'04/15/2008: NicoNZ - Ahora se resetea el counter del invi
'13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
'27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
'21/07/2009: Marco - Al morir se desactiva el comercio seguro.
'************************************************
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    
    With UserList(UserIndex)
        'Sonido
        If .genero = eGenero.Mujer Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_MUJER)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_HOMBRE)
        End If
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHP = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1
        .flags.Equitando = 0
        
        ' No se activa en arenas
        If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
            Call WriteResuscitationSafeOff(UserIndex)
        End If
        
        aN = .flags.AtacadoPorNpc
        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString
        End If
        
        aN = .flags.NPCAtacado
        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString
            End If
        End If
        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(UserIndex)
        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(UserIndex)
        End If
        
        '<<<< Invisible >>>>
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
        End If
        
        If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Or MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
            ' << Si es newbie no pierde el inventario >>
            If Not EsNewbie(UserIndex) Or criminal(UserIndex) Then
                Call TirarTodo(UserIndex)
            Else
                Call TirarTodosLosItemsNoNewbies(UserIndex)
            End If
        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0
        End If
        
        If .flags.AutoTorneo = True Then
            Call Rondas_UsuarioMuere(UserIndex)
        End If
        
        ' << Restauramos el mimetismo
        If .flags.Mimetizado = 1 Then
            .Char.body = .CharMimetizado.body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then
            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
        End If
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.Aura = 0
        Else
            .Char.body = iFragataFantasmal
            .Char.Aura = 0
        End If
        
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                Call MuereNpc(.MascotasIndex(i), 0)
            ' Si estan en agua o zona segura
            Else
                .MascotasType(i) = 0
            End If
        Next i
        
        .NroMascotas = 0
        
        If Titan Then
            For i = 1 To 3
                Dim tmpN As Integer
                If Invocadores(i) > 0 Then
                    If Invocadores(i) = UserIndex Then
                        If UserList(Invocadores(i)).flags.Muerto = 1 Then
                            'murieron todos, desaparece el npc
                            tmpN = Npclist(TitanIndex).Name
                            QuitarNPC TitanIndex
                            Invocadores(i) = 0
                            Titan = False
                        End If
                    End If
                End If
            Next i
        End If
        
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco, .Char.Aura)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
        '<<Castigos por party>>
        If .PartyIndex > 0 Then
            Call mdParty.ObtenerExito(UserIndex, .Stats.ELV * -10 * mdParty.CantMiembros(UserIndex), .Pos.map, .Pos.X, .Pos.Y)
        End If
        
        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(UserIndex)
    End With
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripci�n: " & Err.description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    With UserList(Atacante)
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
        If criminal(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name
                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                    .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
            End If
            
            If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
                .Faccion.Reenlistadas = 200  'jaja que trucho
                
                'con esto evitamos que se vuelva a reenlistar
            End If
        Else
            If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                .flags.LastCiudMatado = UserList(Muerto).Name
                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                    .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
            End If
        End If
        
        If .Stats.UsuariosMatados < MAXUSERMATADOS Then _
            .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1
    End With
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    Dim hayobj As Boolean
    
    hayobj = False
    nPos.map = Pos.map
    nPos.X = 0
    nPos.Y = 0
    
    Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
                
                If LegalPos(nPos.map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.map, tX, tY).ObjInfo.amount + Obj.amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        
                        'break both fors
                        tX = Pos.X + LoopC
                        tY = Pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
    Loop
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal FX As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 15/07/2009
'15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
'**************************************************************
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    With UserList(UserIndex)
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        Call WriteRemoveAllDialogs(UserIndex)
        
        OldMap = .Pos.map
        OldX = .Pos.X
        OldY = .Pos.Y

        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)
        
        If OldMap <> map Then
            Call WriteChangeMap(UserIndex, map, MapInfo(map).Zona)
            Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(map).Music, 45)))
            
            'Update new Map Users
            MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
        End If
        
        .Pos.X = X
        .Pos.Y = Y
        .Pos.map = map
        
        Call MakeUserChar(True, map, UserIndex, map, X, Y)
        Call WriteUserCharIndexInServer(UserIndex)
        
        Call DoTileEvents(UserIndex, map, X, Y)
        
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(UserIndex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            Call SetInvisible(UserIndex, .Char.CharIndex, True)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(.Char.CharIndex, 24, 0))
        End If
        
        If .NroMascotas Then Call WarpMascotas(UserIndex)
        
        
        ' Automatic toogle navigate
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
            If HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then
                If .flags.Navegando = 0 Then
                    .flags.Navegando = 1
                        
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)
                End If
            Else
                If .flags.Navegando = 1 Then
                    .flags.Navegando = 0
                            
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)
                End If
            End If
        End If
      
    End With
End Sub

Private Sub WarpMascotas(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 11/05/2009
'13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
'13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
'11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
'************************************************
    Dim i As Integer
    Dim petType As Integer
    Dim PetRespawn As Boolean
    Dim PetTiempoDeVida As Integer
    Dim NroPets As Integer
    Dim InvocadosMatados As Integer
    Dim canWarp As Boolean
    Dim index As Integer
    Dim iMinHP As Integer
    
    NroPets = UserList(UserIndex).NroMascotas
    canWarp = (MapInfo(UserList(UserIndex).Pos.map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        index = UserList(UserIndex).MascotasIndex(i)
        
        If index > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(UserIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = Npclist(index).flags.Respawn = 0
                petType = UserList(UserIndex).MascotasType(i)
                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = Npclist(index).Stats.MinHP
                
                Call QuitarNPC(index)
                
                ' Restauramos el valor de la variable
                UserList(UserIndex).MascotasType(i) = petType

            End If
        ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0
        End If
        
        If petType > 0 And canWarp Then
            index = SpawnNpc(petType, UserList(UserIndex).Pos, False, PetRespawn)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            ' Exception: Pets don't spawn in water if they can't swim
            If index = 0 Then
                Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Else
                UserList(UserIndex).MascotasIndex(i) = index

                ' Nos aseguramos de que conserve el hp, si estaba da�ado
                Npclist(index).Stats.MinHP = IIf(iMinHP = 0, Npclist(index).Stats.MinHP, iMinHP)
            
                Npclist(index).MaestroUser = UserIndex
                Npclist(index).Movement = TipoAI.SigueAmo
                Npclist(index).Target = 0
                Npclist(index).TargetNPC = 0
                Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(index)
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If Not canWarp Then
        Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. �stas te esperar�n afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    UserList(UserIndex).NroMascotas = NroPets
End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/04/08 (NicoNZ)
'
'***************************************************
    Dim isNotVisible As Boolean
With UserList(UserIndex)
    If .flags.UserLogged And Not .Counters.Saliendo Then
        .Counters.Saliendo = True
        .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).Pk, IntervaloCerrarConexion, 0)
        
        isNotVisible = (.flags.Oculto Or .flags.invisible)
        If isNotVisible Then
            .flags.Oculto = 0
            .flags.invisible = 0
            Call WriteTimeInvi(UserIndex, 0)
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call SetInvisible(UserIndex, .Char.CharIndex, False)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
        End If
        
        If .flags.Metamorfosis = 1 Then
            .Char.Head = .OrigChar.Head
            
            If .Invent.ArmourEqpObjIndex > 0 Then
                .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If
            
            If .Invent.EscudoEqpObjIndex > 0 Then _
                .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
            If .Invent.WeaponEqpObjIndex > 0 Then _
                .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
            If .Invent.CascoEqpObjIndex > 0 Then _
                .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                
            .Stats.Da�oExtra = 0
            .flags.Metamorfosis = 0
            .Counters.Metamorfosis = 0
         
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.Aura)
        End If
        
        Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrar� el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
    End If
End With
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecut� la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
    Dim ViejoNick As String
    Dim ViejoCharBackup As String
    
    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name
    
    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
    End If
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)
    If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Vitalidad: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
#If ConUpTime Then
        Dim TempSecs As Long
        Dim TempStr As String
        TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If
    
    End If
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
    Dim CharFile As String
    
On Error Resume Next
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente
'**************************************************************
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            .Reputacion.BurguesRep = 0
            .Reputacion.NobleRep = 0
            .Reputacion.PlebeRep = 0
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO
            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
        End If
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        .Reputacion.LadronesRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO
        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
    End With
    
    Call RefreshCharStatus(UserIndex)
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, ByVal userCharIndex As Integer, ByVal invisible As Boolean)
Dim sndNick As String
Dim klan As String
Call SendData(SendTarget.ToUsersAreaButGMs, UserIndex, PrepareMessageSetInvisible(userCharIndex, invisible))

If invisible Then
    sndNick = UserList(UserIndex).Name & " " & TAG_USER_INVISIBLE
Else
    sndNick = UserList(UserIndex).Name
    If UserList(UserIndex).GuildIndex > 0 Then
        sndNick = sndNick & " <" & modGuilds.GuildName(UserList(UserIndex).GuildIndex) & ">"
    End If
End If

Call SendData(SendTarget.ToGMsArea, UserIndex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))
End Sub

Public Sub ResucitandoUsuario(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        If .flags.Resucitando = 1 Then Exit Sub
        If .Counters.Saliendo = True Then Exit Sub
        
            If .flags.Muerto = 1 Then
                .flags.Resucitando = 1
                Call WriteConsoleMsg(UserIndex, "Tu cuerpo comienza a tomar forma... En pocos segundos resucitar�s.", FontTypeNames.FONTTYPE_INFO)
            End If
    End With
                
End Sub
