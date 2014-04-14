Imports DotNetWikiBot
Imports VpNet
Imports VpNet.Core
Imports VpNet.Core.Structs
Imports VpNet.NativeApi
Imports System.Runtime.InteropServices
Imports System
Imports System.Globalization
Imports System.IO
Imports System.Threading

'Implement a simple "undo" by having the last 5 objects a user built stored in the user array
Module modMain
    Dim ObjectDatPath As String = Path.Combine(Environment.CurrentDirectory, "vpstats_objects.dat")
    Dim StatsDatPath As String = Path.Combine(Environment.CurrentDirectory, "vpstats.dat")
    Dim ChatLogPath As String = Path.Combine(Environment.CurrentDirectory, "chat.txt")
    Dim ConfigPath As String = Path.Combine(Environment.CurrentDirectory, "Config.ini")

    Public objWriter As System.IO.TextWriter = New System.IO.StreamWriter(ObjectDatPath, True) With {.AutoFlush = True}
    Public objWriterChat As System.IO.TextWriter = New System.IO.StreamWriter(ChatLogPath, True) With {.AutoFlush = True}

    Dim State As connectionState = connectionState.Disconnected

    ' Dim LastMarker As Short
    Dim Exiting As Boolean
    Dim SpinnyObj As VpObject
    Dim SpinnyD As Boolean
    
    Sub Main()
        Try
            LoadConfig()
            SetupBot()

            'Main loop
            While Not Exiting

                ' If DateTime.UtcNow.Subtract(Wiki.CitListLastUpdate).TotalDays >= 4 And Options.EnableWikiUpdates = True Then Wiki.CitListLastUpdate = DateTime.UtcNow : UpdateWikiCitizenList() 'Update citizen list automatically
                ' Bot updates
                Select Case State
                    Case connectionState.Disconnected
                        ' State reset
                        Wiki.CitListPolling = False
                        Queries.Clear()
                        Users.Clear()

                        LoginBot()
                        CheckMarkers() ' Only needs to be called after connection
                        info("Bot is now active. Use commands in-world to control.")

                    Case connectionState.Connected
                        Bot.Instance.Wait(25)
                End Select

                ' Update statistics
                If Options.EnableStatisticsLogging Then
                    If DateTime.UtcNow.Hour > VPStats.LastHour.Hour Then VPStats.LastHour = DateTime.UtcNow : UpdateStatisticsLog()
                    If DateTime.UtcNow.Day > VPStats.LastSave.Day Then VPStats.LastSave = DateTime.UtcNow : SaveStatisticsLog()
                End If

                ' Update cit list
                If Options.EnableWikiUpdates Then
                    PollCitList()
                End If
            End While

            EndProgram() 'End the program

        Catch ex As Exception
            info("FATAL EXCEPTION: " & ex.Message)
        End Try
    End Sub

    Sub EndProgram()
        If Options.EnableStatisticsLogging Then SaveStatisticsLog()

        info("Saving logs...")
        objWriter.Close()
        objWriterChat.Close()

        'Delete all the marker objects
        If Options.EnableMapUpdates Then
            info("Clearing markers...")
            ' Changed For to For Each
            For Each User In Users
                If User.MarkerObjectID = -1 Then Continue For

                Dim markerObject As New VpNet.Core.Structs.VpObject
                markerObject.Id = User.MarkerObjectID
                Bot.Instance.DeleteObject(markerObject)
                Bot.Instance.Wait(100)
            Next
        End If

        Bot.Instance.Dispose()
    End Sub

    Sub LoadConfig()
        'Load configuration from file
        ConfigINI.Load(ConfigPath)

        Bot.LoginName = ConfigINI.GetKeyValue("Bot", "Name")
        Bot.UniHost = ConfigINI.GetKeyValue("Bot", "UniHost")
        Bot.UniPort = Convert.ToInt32(ConfigINI.GetKeyValue("Bot", "UniPort"))
        Bot.CitName = ConfigINI.GetKeyValue("Bot", "CitName")
        Bot.CitPass = ConfigINI.GetKeyValue("Bot", "CitPass")
        Bot.WorldName = ConfigINI.GetKeyValue("Bot", "World")

        Options.EnableMapUpdates = Val2Bool(ConfigINI.GetKeyValue("Options", "EnableMapUpdates"))
        Options.EnableObjectLogging = Val2Bool(ConfigINI.GetKeyValue("Options", "EnableObjectLogging"))
        Options.EnableChatLogging = Val2Bool(ConfigINI.GetKeyValue("Options", "EnableChatLogging"))
        Options.EnableStatisticsLogging = Val2Bool(ConfigINI.GetKeyValue("Options", "EnableStatisticsLogging"))
        Options.EnableWikiUpdates = Val2Bool(ConfigINI.GetKeyValue("Options", "EnableWikiUpdates"))

        Wiki.CitListLastUpdate = DateTime.Parse(ConfigINI.GetKeyValue("Wiki", "CitListLastUpdate"), New CultureInfo("en-GB"))
        Wiki.Username = ConfigINI.GetKeyValue("Wiki", "Username")
        Wiki.Password = ConfigINI.GetKeyValue("Wiki", "Password")
        VPStats.LastSave = DateTime.Parse(ConfigINI.GetKeyValue("Stats", "LastSave"), New CultureInfo("en-GB"))

        ' VPStats.LastSave = DateTime.UtcNow
        info("EnableWikiUpdates = " & Options.EnableWikiUpdates)
        info("EnableMapUpdates = " & Options.EnableMapUpdates)
        info("EnableObjectLogging = " & Options.EnableObjectLogging)
        info("EnableStatisticsLogging = " & Options.EnableStatisticsLogging)
        info("EnableChatLogging = " & Options.EnableChatLogging)
    End Sub

    Sub SetupBot()
        Bot.Instance = New VpNet.Core.Instance

        AddHandler Bot.Instance.EventAvatarAdd, AddressOf vpnet_EventAvatarAdd
        AddHandler Bot.Instance.EventAvatarChange, AddressOf vpnet_EventAvatarChange
        AddHandler Bot.Instance.EventAvatarDelete, AddressOf vpnet_EventAvatarDelete

        AddHandler Bot.Instance.EventObjectCreate, AddressOf vpnet_EventObjectCreate
        AddHandler Bot.Instance.EventObjectChange, AddressOf vpnet_EventObjectChange
        AddHandler Bot.Instance.EventObjectDelete, AddressOf vpnet_EventObjectDelete
        AddHandler Bot.Instance.EventObjectClick, AddressOf vpnet_EventObjectClick

        AddHandler Bot.Instance.EventQueryCellResult, AddressOf vpnet_EventQueryCellResult
        AddHandler Bot.Instance.EventQueryCellEnd, AddressOf vpnet_EventQueryCellEnd
        AddHandler Bot.Instance.CallbackObjectAdd, AddressOf vpnet_CallbackObjectAdd

        AddHandler Bot.Instance.EventUniverseDisconnect, AddressOf vpnet_EventUniverseDisconnect
        AddHandler Bot.Instance.EventWorldDisconnect, AddressOf vpnet_EventWorldDisconnect
        AddHandler Bot.Instance.EventChat, AddressOf vpnet_EventAvatarChat
        AddHandler Bot.Instance.EventUserAttributes, AddressOf vpnet_EventUserAttributes
    End Sub

    Sub LoginBot()
        Dim attempt = 1
        Dim delay = 1000
        Dim nextDelay As Integer

        Do Until State = connectionState.Connected
            Try
                info("Logging into universe...")
                Bot.Instance.Connect(Bot.UniHost, Bot.UniPort)
                Bot.Instance.Login(Bot.CitName, Bot.CitPass, Bot.LoginName)

                info("Entering " & Bot.WorldName & "...")
                Bot.Instance.Enter(Bot.WorldName)
                Bot.Instance.UpdateAvatar(0, -100, 0, 0, 0)

                State = connectionState.Connected
                info("Bot is now connected and in-world")

            Catch ex As Exception
                State = connectionState.Disconnected
                nextDelay = Math.Min(delay * attempt, 60000)

                info(String.Format("Attempt {0} to connect failed. Next attempt in {1}ms", attempt, nextDelay))
                info(ex.Message)

                attempt += 1
                Thread.Sleep(nextDelay)

            End Try
        Loop
    End Sub

    Sub CheckMarkers()
        'Query cells
        For CellXi As Integer = 240 To 260
            For CellZi As Integer = -260 To -240
                Bot.Instance.QueryCell(CellXi, CellZi)
            Next
        Next
    End Sub

    Private Sub vpnet_EventAvatarChat(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.EventData.Chat)
        Dim ChatMessage As String = LTrim(eventData.Message)

        If eventData.ChatType = 1 Then 'Log console messages 
            If eventData.Username = "" Then
                ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "]  " & ChatMessage)
            Else
                ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "]  " & eventData.Username & ":" & vbTab & ChatMessage)
            End If
        Else                            'Log chat messages
            If eventData.Username = "" Then
                ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "] " & ChatMessage)
            Else
                ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "] " & eventData.Username & ":" & vbTab & ChatMessage)
            End If
        End If

        If eventData.ChatType = 1 Then Exit Sub 'No console messages

        Dim User = FindUser(eventData.Session) 'Find user ID
        If User Is Nothing Then Return
        User.statsActiveInLastHour = True

        '  User(un).LastActive = DateTime.Now

        If eventData.Message.Length < 2 Then Exit Sub
        '  Console.WriteLine(eventData.Message)
        'Process commands
        If eventData.Message.ToLower = Bot.LoginName.ToLower & " version" Then vpSay("VP-Utilities Bot by Chris Daxon.", User.Session)

        If eventData.Message.Substring(0, 2).ToLower = "u:" Then
            Dim SelectMsg As String = eventData.Message.Substring(2, eventData.Message.Length - 2).ToLower
            If eventData.Message.Substring(2, eventData.Message.Length - 2).Contains(" ") = True Then
                SelectMsg = SelectMsg.Substring(0, SelectMsg.IndexOf(" "))
            End If

            Select Case SelectMsg
                Case "version"
                    vpSay("VP-Utilities Bot by Chris Daxon.", User.Session)
                Case "coords"
                    vpSay("Coordinates: " & User.X & " " & User.Y & " " & User.Z, User.Session)
                Case "exit"
                    If User.Session = Bot.Owner Then Exiting = True
                Case "freeze"
                    If User.Session = Bot.Owner And eventData.Message.Length > 10 Then
                        Dim CmdSplit() As String
                        CmdSplit = Split(eventData.Message.Substring(2, eventData.Message.Length - 2), " ", 2) 'Easiest way
                        Dim Target = FindUserByName(CmdSplit(1))
                        If Target Is Nothing Then vpSay("No such user.", User.Session) : Exit Sub
                        Target.AvatarFrozen = True
                        vpSay(Target.Name & " has been frozen.", User.Session)
                    End If
                Case "unfreeze"
                    If User.Session = Bot.Owner And eventData.Message.Length > 10 Then
                        Dim CmdSplit() As String
                        CmdSplit = Split(eventData.Message.Substring(2, eventData.Message.Length - 2), " ", 2) 'Easiest way
                        Dim Target = FindUserByName(CmdSplit(1))
                        If Target Is Nothing Then vpSay("No such user.", User.Session) : Exit Sub
                        Target.AvatarFrozen = False
                        vpSay(Target.Name & " has been defrosted.", User.Session)
                    End If
                Case "updatecitlist"
                    If User.Session <> Bot.Owner Then Exit Select
                    UpdateWikiCitizenList()

                Case "add mirror point"
                    vpSay("Command not implemented.", User.Session) : Exit Select
#If False Then ' Dead code
                    'we need directions and for individual users
                    Dim newObject As New VpNet.Core.Structs.VpObject
                    newObject.Position = New VpNet.Core.Structs.Vector3(User.X, User.Y, User.Z)
                    newObject.Rotation = New VpNet.Core.Structs.Vector3(0, 90, 0)
                    newObject.Angle = Single.PositiveInfinity
                    newObject.Description = "Mirror Point" & vbCrLf & "Owner: " & User.Name
                    newObject.Action = "create color green, name mirrorpoint"
                    newObject.Model = "w1pan_1000e"
                    newObject.ReferenceNumber = (1000 + un)
                    Try
                        Bot.Instance.AddObject(newObject)
                    Catch ex As Exception
                        info(ex.Message)
                    End Try
#End If
                Case Else
                    vpSay("Command not recognised.", User.Session)
            End Select
        End If
    End Sub

    Private Sub vpnet_CallbackObjectAdd(ByVal sender As VpNet.Core.Instance, ByVal objectData As VpNet.Core.Structs.VpObject)
        'Update mirror object
        'If objectData.ReferenceNumber > 1000 Then User(objectData.ReferenceNumber - 1000).MirrorObjectData = objectData : Exit Sub

        'Update users markernumber with their objectId
        Dim User = FindUser(objectData.ReferenceNumber)

        If User IsNot Nothing Then
            User.MarkerObjectID = objectData.Id
            User.MarkerPending = False
            info("Marker successfully created for " & User.Name)
        End If
        '   Console.WriteLine(User(objectData.ReferenceNumber).Name & ": " & objectData.ReferenceNumber)
    End Sub

    Private Sub vpnet_EventQueryCellResult(ByVal sender As VpNet.Core.Instance, ByVal obj As VpNet.Core.Structs.VpObject)
        'Delete old markers
        If obj.Action.Contains("name avmarker") Then
            Dim markerUser = FindUserByName(obj.Description)

            If markerUser Is Nothing OrElse markerUser.MarkerObjectID <> obj.Id Then
                info("Queued delete of old marker for " & obj.Description)
                Bot.Instance.DeleteObject(obj)

                Return
            End If
        End If

        'Save object data
        obj.ReferenceNumber = -1
        Queries.Add(obj)

        If obj.Action.Contains("name spinny2") Then SpinnyObj = obj
    End Sub

    Private Sub vpnet_EventQueryCellEnd(ByVal sender As VpNet.Core.Instance, ByVal CellX As Integer, ByVal CellZ As Integer)
    End Sub

    Sub CreateMarker(ByRef User As objUser)
        If Not Options.EnableMapUpdates Then Return

        Randomize()

        ' Skip users that already have markers or are waiting for them
        If User.MarkerObjectID <> -1 OrElse User.MarkerPending Then Return

        ' No markers for bots, except for [Cat]
        If User.Name.Substring(0, 1) = "[" And User.Name <> "[Cat]" Then Return

        User.MarkerObjectAction = "create solid no,color " & Hex(Int(Rnd() * 255)).PadRight(2, "0") & Hex(Int(Rnd() * 255)).PadRight(2, "0") & Hex(Int(Rnd() * 255)).PadRight(2, "0") & ",move {x} 0 {z} time={t} wait=9e9,rotate {r} time={rt} wait=9e9 nosync,name avmarker"

        Dim markerObject As New VpNet.Core.Structs.VpObject
        markerObject.Position = New VpNet.Core.Structs.Vector3(250, 0.014, -250)
        markerObject.Rotation = New VpNet.Core.Structs.Vector3(0, 0, 0)
        markerObject.Angle = Single.PositiveInfinity
        markerObject.Description = User.Name
        markerObject.Action = User.MarkerObjectAction.Replace("{x}", 0).Replace("{z}", 0).Replace("{t}", 0).Replace("{r}", 0).Replace("{rt}", 0)
        markerObject.Model = "cyfigure.rwx"
        markerObject.ReferenceNumber = User.Session

        Bot.Instance.AddObject(markerObject)
        User.MarkerPending = True
        info("Queued add marker for " & User.Name)
    End Sub

    Sub UpdateMarker(ByRef User As objUser) 'Updates the map marker for an avatar
        If Not Options.EnableMapUpdates Then Return

        'If nessecary, update map marker location
        If Not User.MovedSinceLastMarkerUpdate OrElse User.Session = 0 OrElse User.MarkerObjectID = -1 Then Return

        User.MovedSinceLastMarkerUpdate = False

        Dim movX As Single = (User.X - User.oldX) / 15 'Controls the distance and direction
        Dim movZ As Single = (User.Z - User.oldZ) / 15 'TODO: make this more precise, could even use the yaw of the user
        Dim movTime As Single = 0.3 'Move must finish by this time; therefore controls the speed.

        Dim markerYAW As Single = User.Yaw '-(User(i).YAW) + 180
        Dim markeroldYAW As Single = User.oldYAW '-(User(i).oldYAW) + 180
        Dim rotYAW As Integer = markerYAW - markeroldYAW '(markerYAW Mod 360) - (markeroldYAW Mod 360)
        '  If rotYAW < 0 Then rotYAW += 360

        'Dim rotY As Single = ((Math.Abs(rotYAW) / 360) * 120)
        Dim rotY As Single = ((Math.Abs(rotYAW) / 360) * 120) 'RPM, 500ms. 60 = 1 rotation in a second
        Dim rotTime As Single = 0.3

        If markerYAW < markeroldYAW Then rotY = -rotY 'Reverse rotation
        'Dim rotTime As Single = (rotYAW / 360)
        rotY = -rotY 'Invert because nessecary

        If markerYAW = markeroldYAW Then rotY = 0 : rotTime = 0 'Don't rotate if user isn't, may be redundant.
        Dim markerObject As New VpNet.Core.Structs.VpObject
        markerObject.Position = New VpNet.Core.Structs.Vector3(250 + (User.oldX / 100), 0.014, -250 + (User.oldZ / 100))
        'markerObject.Position = New VpNet.Core.Structs.Vector3(250 + (User(i).X / 100), 0.014, -250 + (User(i).Z / 100))
        markerObject.Rotation = New VpNet.Core.Structs.Vector3(0, -(User.Yaw) + 180, 0)
        markerObject.Angle = Single.PositiveInfinity
        markerObject.Description = User.Name
        markerObject.Action = User.MarkerObjectAction.Replace("{x}", movX).Replace("{z}", movZ).Replace("{t}", movTime).Replace("{r}", rotY).Replace("{rt}", rotTime)
        markerObject.ReferenceNumber = User.Session
        markerObject.Model = "cyfigure.rwx"
        markerObject.Id = User.MarkerObjectID

        Bot.Instance.ChangeObject(markerObject)

        'Store old position
        User.oldX = User.X
        User.oldY = User.Y
        User.oldZ = User.Z
        User.oldYAW = User.Yaw
    End Sub

    Private Sub vpnet_EventAvatarAdd(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.Structs.Avatar)
        Dim newUser As objUser = New objUser

        newUser.Name = eventData.Name
        newUser.Session = eventData.Session
        newUser.AvatarType = eventData.AvatarType
        newUser.X = eventData.X
        newUser.Y = eventData.Y
        newUser.Z = eventData.Z
        newUser.oldX = eventData.X
        newUser.oldY = eventData.Y
        newUser.oldZ = eventData.Z
        newUser.MarkerObjectID = -1
        newUser.Id = eventData.Id

        Users.Add(newUser)

        'Set bot owner
        'TODO: Use code from chatlink for multiple owners
        If newUser.Id = 104 OrElse newUser.Name = "Chris D" Then
            Bot.Owner = newUser.Session
            info("Bot owner detected. Session: " & Bot.Owner)
        End If

        ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "] ENTERS: " & eventData.Name & ", " & eventData.Id & ", " & eventData.Session)
        CreateMarker(newUser)
    End Sub

    Private Sub vpnet_EventAvatarChange(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.Structs.Avatar)
        Dim User = FindUser(eventData.Session)
        If User Is Nothing Then Return

        If User.X <> eventData.X OrElse User.Y <> eventData.Y OrElse User.Z <> eventData.Z OrElse User.Yaw <> eventData.Yaw Then

            If User.X = eventData.X And User.Y = eventData.Y And User.Z = eventData.Z Then
                'Ignore small changes in YAW
                Dim YAWdiff As Single = Math.Abs(eventData.Yaw - User.Yaw)
                If YAWdiff < 0.3 Then GoTo UpdateUserArray 'Ignore small changes in YAW
            End If

            User.MovedSinceLastMarkerUpdate = True
            If User.X <> 0 And User.Z <> 0 Then User.statsActiveInLastHour = True 'If they've moved, set them as active in last hour
            If User.AvatarFrozen = True Then
                Bot.Instance.TeleportAvatar(User.Session, "", User.X, User.Y, User.Z, eventData.Yaw, eventData.Pitch)
                Return
            End If
        End If
UpdateUserArray:

        User.X = eventData.X
        User.Y = eventData.Y
        User.Z = eventData.Z
        User.Yaw = eventData.Yaw
        User.Pitch = eventData.Pitch
        User.AvatarType = eventData.AvatarType

        'Update map location
        UpdateMarker(User)
        '  If User(i).oldX = User(i).X And User(i).oldZ = User(i).Z And User(i).Name = "Chris D" Then info(User(i).YAW)
    End Sub

    Private Sub vpnet_EventAvatarDelete(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.Structs.Avatar)
        Dim User = FindUser(eventData.Session)
        If User Is Nothing Then Return

        ' Check if they have marker
        If User.MarkerObjectID = -1 Then Return

        info("Removing marker for " & User.Name)
        'Delete marker object
        Dim markerObject As New VpNet.Core.Structs.VpObject
        markerObject.Id = User.MarkerObjectID
        Bot.Instance.DeleteObject(markerObject)

        Users.Remove(User)
        ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "] EXITS: " & User.Name & ", " & User.Id & ", " & User.Session)
    End Sub

    Private Sub vpnet_EventObjectClick(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal objectId As Integer, ByVal clickHitX As Single, ByVal clickHitY As Single, ByVal clickHitZ As Single)
        'TODO: neaten out

        ' Check if av figure was clicked
        For Each User In Users
            If User.MarkerObjectID = objectId Then
                vpSay("User: " & User.Name & " (" & User.Id & ")", sessionId)
                Return
            End If
        Next

        If Queries Is Nothing Then Return

        ' Check if MapTile was clicked
        Dim MapTile As VpObject = Nothing
        For Each Q In Queries
            If Q.Id = objectId Then
                MapTile = Q
                Exit For
            End If
        Next

        If MapTile Is Nothing Then Return

        Dim Clicker = FindUser(sessionId)

        'User clicked map
        If MapTile.Action.Contains("name mapobject") Then
            'vpSay((clickHitX - 250) * 100 & " 2 " & (clickHitZ + 250) * 100, sessionId)
            If DateTime.Now.Subtract(Clicker.MapClickTime).TotalSeconds <= 2 Then
                vpSay("Teleporting to " & (clickHitX - 250) * 100 & " 2 " & (clickHitZ + 250) * 100, sessionId)
                Bot.Instance.TeleportAvatar(sessionId, "", (clickHitX - 250) * 100, 2, (clickHitZ + 250) * 100, 0, 0)
                Clicker.MapClickTime = Nothing
            Else
                Clicker.MapClickTime = DateTime.Now 'Wait for another click within 2 seconds to classify as a "double click"
            End If
        End If
    End Sub

    Private Sub vpnet_EventObjectCreate(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal vpObject As VpNet.Core.Structs.VpObject)
        If vpObject.Action.Contains("name avmarker") Then Exit Sub 'Ignore bots own changes

        Dim ObjectLine As String

        'PREFIXED TO THE VPPTSV1 FORMAT:
        'Type (0 = create, 1 = change, 2 = delete)
        'Session
        'User ID
        'Object ID

        If Options.EnableObjectLogging = True Then
            Dim userId = 0
            Dim user = FindUser(sessionId)
            If user IsNot Nothing Then userId = user.Id

            ObjectLine = "0" & vbTab & sessionId & vbTab & userId & vbTab & vpObject.Id & vbTab & vpObject.Owner & vbTab & vpObject.Time.ToString & vbTab & vpObject.Position.X & vbTab & vpObject.Position.Y & vbTab & vpObject.Position.Z & vbTab & vpObject.Rotation.X & vbTab & vpObject.Rotation.Y & vbTab & vpObject.Rotation.Z & vbTab & vpObject.Angle & vbTab & vpObject.ObjectType & vbTab & Escape(vpObject.Model) & vbTab & Escape(vpObject.Description) & vbTab & Escape(vpObject.Action)

            objWriter.WriteLine(ObjectLine)
        End If

        Exit Sub 'TODO: Mirror code (for making builds that mirror each side) is below
        For Each User In Users
            If User.MirrorObjectData.Id > 0 Then
                Mirrors.Add(vpObject)

                'Now create the mirror
                'either add or subtract 180 (half of a full circle, 360 degrees),
                'and use abs()


                Exit Sub 'TODO: remove
            End If
        Next
    End Sub

    Sub ChangeMirrorObject(ByRef mirrObj As VpObject, ByRef User As objUser)
#If False Then ' Figure out what this does and re-enable later
        If mirrObj.Position.X < User.MirrorObjectData.Position.X Then Exit Sub
        'X> mirror only

        Dim offsetX As Integer = mirrObj.Position.X - User.MirrorObjectData.Position.X

        Dim newObject As New VpNet.Core.Structs.VpObject
        newObject = Mirrors(mirrObj)

        'newObject.Position.X = newObject.Position.X - offsetX 'TODO: Error	3	Expression is a value and therefore cannot be the target of an assignment.	E:\Documents and Settings\Chris\My Documents\Visual Studio 2010\Projects\VPUtilities\ConsoleApplication1\ConsoleApplication1\modMain.vb	540	9	ConsoleApplication1
        newObject.Rotation = New VpNet.Core.Structs.Vector3(0, 90, 0)
        Try
            Bot.Instance.AddObject(newObject)
        Catch ex As Exception
            info(ex.Message)
        End Try
#End If
    End Sub

    Private Sub vpnet_EventObjectChange(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal vpObject As VpNet.Core.Structs.VpObject)
        If vpObject.Action.Contains("name avmarker") Then Exit Sub 'Ignore bots own changes
        Dim ObjectLine As String

        'PREFIXED TO THE VPPTSV1 FORMAT:
        'Type (0 = create, 1 = change, 2 = delete)
        'Session
        'User ID
        'Object ID
        If Options.EnableObjectLogging = True Then
            Dim userId = 0
            Dim user = FindUser(sessionId)
            If user IsNot Nothing Then userId = user.Id

            ObjectLine = "1" & vbTab & sessionId & vbTab & userId & vbTab & vpObject.Id & vbTab & vpObject.Owner & vbTab & vpObject.Time.ToString & vbTab & vpObject.Position.X & vbTab & vpObject.Position.Y & vbTab & vpObject.Position.Z & vbTab & vpObject.Rotation.X & vbTab & vpObject.Rotation.Y & vbTab & vpObject.Rotation.Z & vbTab & vpObject.Angle & vbTab & vpObject.ObjectType & vbTab & Escape(vpObject.Model) & vbTab & Escape(vpObject.Description) & vbTab & Escape(vpObject.Action)

            objWriter.WriteLine(ObjectLine)
        End If

        'TODO: Mirror object code
#If False Then
        'Update mirror object location 
        If vpObject.Action.Contains("name mirrorpoint") Then
            For Each User In Users
                If User.Id = vpObject.Owner Then User.MirrorObjectData = vpObject
            Next
        End If

        'Mirror objects
        For Each User In Users
            If User.MirrorObjectData.Id = 0 Then Continue For

            For Each Mirror In Mirrors
                If Mirror.Id = vpObject.Id Then
                    ChangeMirrorObject(Mirror, User)
                    '<X mirror
                End If
            Next
        Next
#End If

    End Sub

    Private Sub vpnet_EventObjectDelete(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal objectId As Integer)
        For Each User In Users 'Ignore bots own changes
            If User.MarkerObjectID = objectId Then Exit Sub
        Next

        Dim ObjectLine As String

        'PREFIXED TO THE VPPTSV1 FORMAT:
        'Type (0 = create, 1 = change, 2 = delete)
        'Session
        'User ID
        'Object ID
        If Options.EnableObjectLogging = True Then
            Dim userId = 0
            Dim user = FindUser(sessionId)
            If user IsNot Nothing Then userId = user.Id

            ObjectLine = "2" & vbTab & sessionId & vbTab & userId & vbTab & objectId & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab

            objWriter.WriteLine(ObjectLine)
        End If
    End Sub

    Private Sub vpnet_EventWorldDisconnect(ByVal sender As VpNet.Core.Instance)
        State = connectionState.Disconnected
        info("World connection lost")
    End Sub

    Private Sub vpnet_EventUniverseDisconnect(ByVal sender As VpNet.Core.Instance)
        State = connectionState.Disconnected
        info("Universe connection lost")
    End Sub

    Private Sub vpnet_EventUserAttributes(ByVal sender As VpNet.Core.Instance, ByVal userAttributes As VpNet.Core.Structs.UserAttributes)
        Citizens.Add(userAttributes)
    End Sub

    Sub PollCitList()
        If Not Wiki.CitListPolling Then Return

        ' All citizens have been collected toward the current target, begin next batch
        If Citizens.Count = Wiki.CitListPollTarget Then
            For count As Integer = 1 To 25
                ' Conviniently, this skips the 0 user on first batch
                Bot.Instance.UserAttributesById(count + Wiki.CitListPollTarget)
            Next

            Wiki.CitListLastUpdate = DateTime.UtcNow
            Wiki.CitListPollTarget += 25
            info(String.Format("Citizen poll: {0} found, 25 more to go", Citizens.Count))
            Return
        End If

        ' Assume after a 30 second timeout of no further updates that all citizens have
        ' been collected. Begins upload to wiki
        If DateTime.Now.Subtract(Wiki.CitListLastUpdate).TotalSeconds >= 30 Then
            info("Citizen poll: Timed out, uploading gathered data to wiki")
            Wiki.CitListPolling = False
            UploadWikiCitizenList()
        End If
    End Sub

    Sub UpdateWikiCitizenList()
        If Wiki.CitListPolling Then
            info("Citizen list update is already in progress")
            Return
        End If

        Citizens.Clear()
        Wiki.CitListPolling = True
        Wiki.CitListPollTarget = 0
        info("Begun polling of citizen list")
    End Sub

    Sub UploadWikiCitizenList()
        If Citizens.Count = 0 Then
            info("Cannot upload to wiki as Citizen list is empty (did polling go wrong?)")
            Return
        End If

        'Save the last wiki update time
        ConfigINI.SetKeyValue("Wiki", "CitListLastUpdate", Wiki.CitListLastUpdate.ToString(New CultureInfo("en-GB")))
        ConfigINI.Save(ConfigPath)

        Dim outText As String = Replace(My.Resources.WikiText1a, "{date}", Wiki.CitListLastUpdate.ToString(New CultureInfo("en-GB")))

        For Each Citizen In Citizens
            If Citizen.Name = "" Then Continue For
            'Generate online time string
            Dim nTimeInSeconds As Long, nHours As Long, nMinutes As Long, nSeconds As Long, nDays As Long
            nTimeInSeconds = Citizen.OnlineTime
            nDays = nTimeInSeconds \ 86400
            nHours = (nTimeInSeconds \ 3600) - nDays * 24
            nTimeInSeconds = nTimeInSeconds Mod 3600
            nMinutes = nTimeInSeconds \ 60
            nSeconds = nTimeInSeconds Mod 60

            'Add column to output text
            If DateTime.Now.Subtract(UnixTimestampToDateTime(Citizen.RegistrationTime)).TotalDays < 30 Then
                outText = outText & vbNewLine & "|-" & vbNewLine & "| " & Citizen.Id & " || [[User:" & Citizen.Name.Replace(" ", "_") & "|" & Citizen.Name & "]] || <span style=""color:red"">'''" & UnixTimestampToDateTime(Citizen.RegistrationTime).ToString(New CultureInfo("en-GB")) & "'''</span> || " & nDays & " day(s), " & nHours & " hour(s), " & nMinutes & " minute(s), " & nSeconds & " second(s)"
            ElseIf Citizen.Id = 104 Then 'TODO: Look into having a method so public users can link to their wiki pages if different (like mine)
                outText = outText & vbNewLine & "|-" & vbNewLine & "| " & Citizen.Id & " || [[User:Chris|" & Citizen.Name & "]] || " & UnixTimestampToDateTime(Citizen.RegistrationTime).ToString(New CultureInfo("en-GB")) & " || " & nDays & " day(s), " & nHours & " hour(s), " & nMinutes & " minute(s), " & nSeconds & " second(s)"
            Else
                outText = outText & vbNewLine & "|-" & vbNewLine & "| " & Citizen.Id & " || [[User:" & Citizen.Name.Replace(" ", "_") & "|" & Citizen.Name & "]] || " & UnixTimestampToDateTime(Citizen.RegistrationTime).ToString(New CultureInfo("en-GB")) & " || " & nDays & " day(s), " & nHours & " hour(s), " & nMinutes & " minute(s), " & nSeconds & " second(s)"
            End If
        Next

        outText = outText & vbNewLine & My.Resources.WikiText1b

        '  Dim objWriter2 As System.IO.TextWriter = New System.IO.StreamWriter("e:/tmp.txt", False)
        '  objWriter2.Write(outText)
        '  objWriter2.Close()
        '  Exit Sub

        'Inherits Bot
        ' Firstly make Site object, specifying site's URL and your bot account
        Dim vpWiki As New DotNetWikiBot.Site("http://wiki.virtualparadise.org", Wiki.Username, Wiki.Password)
        ' Then make Page object, specifying site and page title in constructor
        Dim p As New DotNetWikiBot.Page(vpWiki, "List_of_citizens")
        ' Load actual page text from live wiki
        p.Load()
        ' Save "Art" article's text back to live wiki with specified comment
        p.Save(outText, "Auto-update of citizen list", True)
        info("List of Citizens updated on wiki")
        Citizens.Clear()
    End Sub

    Sub UpdateStatisticsLog()
        Try
            'CONSIDER THE VALUE OF INDIVIDUAL USER DATA, WHAT BENEFIT IS IT ON A GRAPH? 
            'Well, you can have user pie charts/bar graphs showing user activity - you can see online patterns, etc
            info("Saving stats for the last hour " & (DateTime.UtcNow.Hour - 1))
            'Go through all users with a session number, get current active and inactive users by ID, add to a list for active and a list for inactive (combine for total, count )

            Dim LastHour As Integer
            If DateTime.UtcNow.Hour = 0 Then LastHour = 23 Else LastHour = (DateTime.UtcNow.Hour - 1)
            'Inactive users
            VPStats.UserActivity += "," & (LastHour).ToString & "|"
            For Each User In Users
                If User.Session <> 0 And User.Name.Length > 2 Then
                    If User.Name.Substring(0, 1) <> "[" And User.statsActiveInLastHour = False Then
                        VPStats.UserActivity += ":" & User.Id
                    End If
                End If
            Next
            'Active users
            VPStats.UserActivity += "|"
            For Each User In Users
                If User.Session <> 0 And User.Name.Length > 2 Then
                    If User.Name.Substring(0, 1) <> "[" And User.statsActiveInLastHour = True Then
                        VPStats.UserActivity += ":" & User.Id : User.statsActiveInLastHour = False
                    End If
                End If
            Next
            'This will still be tainted with duplicate logins, but that can be delt with when processing the stats.dat file for the actual stats.
        Catch ex As Exception 'TODO: This can be removed once any bugs here are fully fixed
            info("EXCEPTION: " & ex.Message)
            Exit Sub
        End Try
    End Sub

    Sub SaveStatisticsLog()
        Try
            If VPStats.UserActivity = "" Then Return

            VPStats.LastHour = DateTime.UtcNow : UpdateStatisticsLog() 'Log for 23:00 - The previous update won't notice because the hour wasn't more than the previous hour - so we need to make a log for the last hour, as midnight is still

            'Save the last stats update time
            ConfigINI.SetKeyValue("Stats", "LastSave", VPStats.LastSave.ToString(New CultureInfo("en-GB")))
            ConfigINI.Save(ConfigPath)

            'Save user stats to log file
            'The "Using" block auto-disposes of the StreamWriter at the end
            Using StatsDatWriter = New System.IO.StreamWriter(StatsDatPath, True)
                Dim outText As String = Date.UtcNow.AddDays(-1).ToString(New CultureInfo("en-GB")) & VPStats.UserActivity

                StatsDatWriter.WriteLine(outText & vbNewLine)
                StatsDatWriter.Close()
            End Using

            VPStats.UserActivity = ""
            For Each User In Users
                User.statsActiveInLastHour = False
            Next

            info("Saved today's statistics logs.")
        Catch ex As Exception
            info("EXCEPTION: " & ex.Message)
        End Try
    End Sub


End Module
