Imports DotNetWikiBot
Imports VpNet
Imports VpNet.Core
Imports VpNet.Core.Structs
'Imports VpNet.Core.EventData
Imports VpNet.NativeApi
Imports System.Runtime.InteropServices
Imports System
Imports System.Globalization
'Implement a simple "undo" by having the last 5 objects a user built stored in the user array

Module modMain
    Dim objWriter As System.IO.TextWriter = New System.IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory & "vpstats_objects.dat", True)
    Public objWriterChat As System.IO.TextWriter = New System.IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory & "chat.txt", True)
    'Inherits Bot
    Public vp As VpNet.Core.Instance
    ' Dim LastMarker As Short
    Dim ProgramIsClosing As Boolean
    Dim Timer1 As New System.Timers.Timer
    Dim Timer2 As New System.Timers.Timer
    Dim Timer3 As New System.Timers.Timer
    Dim ReconnectTimer As New System.Timers.Timer
    Dim SpinnyID As Integer
    Dim SpinnyD As Boolean
    Dim VpConnected As Boolean 'TODO: Attempt to fix crashing bug
    Sub Main()
        Try
            'Initialise timers
            Timer1.AutoReset = True
            Timer1.Interval = 80
            AddHandler Timer1.Elapsed, AddressOf Timer1_Tick

            Timer2.AutoReset = True
            Timer2.Interval = 300
            AddHandler Timer2.Elapsed, AddressOf Timer2_Tick

            ReconnectTimer.AutoReset = True
            ReconnectTimer.Interval = 25000
            AddHandler ReconnectTimer.Elapsed, AddressOf ReconnectTimer_Tick

            Timer3.AutoReset = True
            Timer3.Interval = 10000
            Timer3.Enabled = False
            AddHandler Timer3.Elapsed, AddressOf Timer3_Tick

            'Load configuration from file
            ConfigINI.Load(System.AppDomain.CurrentDomain.BaseDirectory & "Config.ini")

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
            Dim RetryCount As Byte = 0
            ' VPStats.LastSave = DateTime.UtcNow
            info("EnableWikiUpdates = " & Options.EnableWikiUpdates)
            info("EnableMapUpdates = " & Options.EnableMapUpdates)
            info("EnableObjectLogging = " & Options.EnableObjectLogging)
            info("EnableStatisticsLogging = " & Options.EnableStatisticsLogging)
            info("EnableChatLogging = " & Options.EnableChatLogging)

RetryLogin:
            If LoginBot() = False Then 'Attempt login. Retry a few times. End program if fails.
                If RetryCount > 2 Then End
                RetryCount = RetryCount + 1
                info("Login failed. Retrying...")
                Wait(100)
                GoTo RetryLogin
            End If

            'Main loop
            Do
                If DateTime.UtcNow.Subtract(Wiki.CitListLastUpdate).TotalDays >= 4 And Options.EnableWikiUpdates = True Then Wiki.CitListLastUpdate = DateTime.UtcNow : UpdateWikiCitizenList() 'Update citizen list automatically

                If Options.EnableStatisticsLogging = True Then
                    If DateTime.UtcNow.Hour > VPStats.LastHour.Hour Then VPStats.LastHour = DateTime.UtcNow : UpdateStatisticsLog()
                    If DateTime.UtcNow.Day > VPStats.LastSave.Day Then VPStats.LastSave = DateTime.UtcNow : SaveStatisticsLog()
                End If

                'Input console commands
                If True = False Then 'Quick way to COMMENT OUT everything here
                    Dim input As String = ""
                    ' Console.Write(vbBack & "> " & Chr(Console.Read))
                    Console.Write("> ")
                    Dim curTime As DateTime = DateTime.UtcNow
                    input = Console.ReadLine
                    Select Case input.ToLower
                        Case "exit"
                            Exit Do
                        Case "updatecitlist"
                            Wiki.CitListLastUpdate = DateTime.UtcNow : UpdateWikiCitizenList() 'Update citizen list manually
                        Case "arraylength"
                            Console.WriteLine("User = " & User.Length)
                            Console.WriteLine("QueryData = " & QueryData.Length)
                            Console.WriteLine("Marker = " & Marker.Length)
                        Case "showoptions"
                            Console.WriteLine("EnableWikiUpdates = " & Options.EnableWikiUpdates)
                            Console.WriteLine("EnableMapUpdates = " & Options.EnableMapUpdates)
                            Console.WriteLine("EnableObjectLogging = " & Options.EnableObjectLogging)
                        Case Else
                            Console.WriteLine("Command not recognised.")
                    End Select
                End If
                'TODO: Add "sleep commands" here to stop 100% cpu usage by Mono
            Loop

            If Options.EnableStatisticsLogging = True Then SaveStatisticsLog()
            EndProgram() 'End the program

        Catch ex2 As Exception
            End
        End Try


    End Sub

    Sub EndProgram()
        ProgramIsClosing = True

        info("Saving logs...")
        Timer3.Stop()
        objWriter.Close()
        objWriterChat.Close()




        'Delete all the marker objects

        If Options.EnableMapUpdates = True Then
            info("Clearing markers...")
            For i As Integer = 1 To User.GetUpperBound(0)
                If User(i).MarkerObjectID > 0 Then
                    Dim markerObject As New VpNet.Core.Structs.VpObject
                    markerObject.Id = User(i).MarkerObjectID
                    vp.DeleteObject(markerObject)
                    vp.Wait(100)
                End If
            Next
        End If

        End
    End Sub


    Function LoginBot()

        ReDim User(1)
        ReDim QueryData(0)
        ReDim UserAttribute(0)
        ' LastMarker = 0

        Try
            vp = New VpNet.Core.Instance()
            Timer1.Start()

            vp.Wait(1)

            AddHandler vp.EventAvatarAdd, AddressOf vpnet_EventAvatarAdd
            AddHandler vp.EventAvatarChange, AddressOf vpnet_EventAvatarChange
            AddHandler vp.EventAvatarDelete, AddressOf vpnet_EventAvatarDelete
            AddHandler vp.EventObjectCreate, AddressOf vpnet_EventObjectCreate
            AddHandler vp.EventObjectChange, AddressOf vpnet_EventObjectChange
            AddHandler vp.EventObjectDelete, AddressOf vpnet_EventObjectDelete

            AddHandler vp.EventObjectClick, AddressOf vpnet_EventObjectClick
            AddHandler vp.EventQueryCellResult, AddressOf vpnet_EventQueryCellResult
            AddHandler vp.EventQueryCellEnd, AddressOf vpnet_EventQueryCellEnd
            AddHandler vp.CallbackObjectAdd, AddressOf vpnet_CallbackObjectAdd
            AddHandler vp.EventUniverseDisconnect, AddressOf vpnet_EventUniverseDisconnect
            AddHandler vp.EventWorldDisconnect, AddressOf vpnet_EventWorldDisconnect
            AddHandler vp.EventChat, AddressOf vpnet_EventAvatarChat
            AddHandler vp.EventUserAttributes, AddressOf vpnet_EventUserAttributes
RepeatLogin:
            info("Logging into universe...")
            vp.Connect(Bot.UniHost, Bot.UniPort)
            vp.Wait(10)

            Try
                vp.Login(Bot.CitName, Bot.CitPass, Bot.LoginName)
            Catch ex As Exception
                If ex.Message.Contains("17") Then 'Strange error, repeat login.
                    info(ex.Message)
                    GoTo RepeatLogin
                End If
            End Try
            vp.Wait(10)
            VpConnected = True
            ProgramIsClosing = False
            info("Entering " & Bot.WorldName & "...")
            vp.Enter(Bot.WorldName)
            vp.UpdateAvatar(0, -100, 0, 0, 0)

        Catch ex As Exception
            info(ex.Message)
            Timer1.Stop()
            Return False
        End Try

            Try
                'Query cells
                For CellXi As Integer = 240 To 260
                    For CellZi As Integer = -260 To -240
                        vp.QueryCell(CellXi, CellZi)
                    Next
                Next
                Wait(5000) 'Wait for query and user avatar adds

                'Scan for old markers
                For o = 1 To QueryData.GetUpperBound(0)
                    'Delete old markers
                    If QueryData(o).Action.Contains("name avmarker") Then
                    info("Deleted old marker for " & QueryData(o).Description)
                        'Delete old objects
                        vp.DeleteObject(QueryData(o))
                    End If
                Next
                If VpConnected = False Then Return False

            Catch ex As Exception
            info(ex.Message)
                Timer1.Stop()
                Return False
            End Try
            If Options.EnableMapUpdates = True Then Timer2.Start()
            Timer3.Start()


            If CreateMarkers() = False Then Timer2.Stop() : Timer3.Stop() : Timer1.Stop() : Return False 'Attempt to create marker objects

        info("Bot is now active. Use commands in-world to control.")

            Return True
    End Function

    Private Sub vpnet_EventAvatarChat(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.EventData.Chat)
        If ProgramIsClosing = True Then Exit Sub
        If VpConnected = False Then Exit Sub

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

        Dim un As Integer = FindUser(eventData.Session) 'Find user ID
        If un = 0 Then Exit Sub
        User(un).statsActiveInLastHour = True

        '  User(un).LastActive = DateTime.Now

        If eventData.Message.Length < 2 Then Exit Sub
        '  Console.WriteLine(eventData.Message)
        'Process commands
        If eventData.Message.ToLower = Bot.LoginName.ToLower & " version" Then vpSay("VP-Utilities Bot by Chris Daxon.", User(un).Session)

        If eventData.Message.Substring(0, 2).ToLower = "u:" Then
            Dim SelectMsg As String = eventData.Message.Substring(2, eventData.Message.Length - 2).ToLower
            If eventData.Message.Substring(2, eventData.Message.Length - 2).Contains(" ") = True Then
                SelectMsg = SelectMsg.Substring(0, SelectMsg.IndexOf(" "))
            End If

            Select Case SelectMsg
                Case "version"
                    vpSay("VP-Utilities Bot by Chris Daxon.", User(un).Session)
                Case "coords"
                    vpSay("Coordinates: " & User(un).X & " " & User(un).Y & " " & User(un).Z, User(un).Session)
                Case "exit"
                    If User(un).Session = Bot.Owner Then EndProgram()
                Case "freeze"
                    If User(un).Session = Bot.Owner And eventData.Message.Length > 10 Then
                        Dim CmdSplit() As String
                        CmdSplit = Split(eventData.Message.Substring(2, eventData.Message.Length - 2), " ", 2) 'Easiest way
                        Dim k As Integer = FindUserByName(CmdSplit(1))
                        If k = 0 Then vpSay("No such user.", User(un).Session) : Exit Sub
                        User(k).AvatarFrozen = True
                        vpSay(User(k).Name & " has been frozen.", User(un).Session)
                    End If
                Case "unfreeze"
                    If User(un).Session = Bot.Owner And eventData.Message.Length > 10 Then
                        Dim CmdSplit() As String
                        CmdSplit = Split(eventData.Message.Substring(2, eventData.Message.Length - 2), " ", 2) 'Easiest way
                        Dim k As Integer = FindUserByName(CmdSplit(1))
                        If k = 0 Then vpSay("No such user.", User(un).Session) : Exit Sub
                        User(k).AvatarFrozen = False
                        vpSay(User(k).Name & " has been defrosted.", User(un).Session)
                    End If
                Case "updatecitlist"
                    If User(un).Session <> Bot.Owner Then Exit Select
                    Wiki.CitListLastUpdate = DateTime.UtcNow : UpdateWikiCitizenList()

                Case "add mirror point"
                    vpSay("Command not implemented.", User(un).Session) : Exit Select

                    'we need directions and for individual users
                    Dim newObject As New VpNet.Core.Structs.VpObject
                    newObject.Position = New VpNet.Core.Structs.Vector3(User(un).X, User(un).Y, User(un).Z)
                    newObject.Rotation = New VpNet.Core.Structs.Vector3(0, 90, 0)
                    newObject.Angle = Single.PositiveInfinity
                    newObject.Description = "Mirror Point" & vbCrLf & "Owner: " & User(un).Name
                    newObject.Action = "create color green, name mirrorpoint"
                    newObject.Model = "w1pan_1000e"
                    newObject.ReferenceNumber = (1000 + un)
                    Try
                        vp.AddObject(newObject)
                    Catch ex As Exception
                        info(ex.Message)
                    End Try
                Case Else
                    vpSay("Command not recognised.", User(un).Session)
            End Select
        End If
    End Sub

    Private Sub vpnet_CallbackObjectAdd(ByVal sender As VpNet.Core.Instance, ByVal objectData As VpNet.Core.Structs.VpObject)
        'Update mirror object
        'If objectData.ReferenceNumber > 1000 Then User(objectData.ReferenceNumber - 1000).MirrorObjectData = objectData : Exit Sub

        'Update users markernumber with their objectId
        If objectData.ReferenceNumber > User.GetUpperBound(0) Or objectData.ReferenceNumber <= 0 Then Exit Sub
        User(objectData.ReferenceNumber).MarkerObjectID = objectData.Id
        '   Console.WriteLine(User(objectData.ReferenceNumber).Name & ": " & objectData.ReferenceNumber)
    End Sub

    Private Sub vpnet_EventQueryCellResult(ByVal sender As VpNet.Core.Instance, ByVal objectData As VpNet.Core.Structs.VpObject)
        'Save object data
        ReDim Preserve QueryData(QueryData.GetUpperBound(0) + 1)
        QueryData(QueryData.GetUpperBound(0)) = objectData
        QueryData(QueryData.GetUpperBound(0)).ReferenceNumber = -1

        'If objectData.Action.Contains("name spinny2") Then SpinnyID = QueryData.GetUpperBound(0) : Timer3.Start() : Console.WriteLine("started")
    End Sub

    Private Sub vpnet_EventQueryCellEnd(ByVal sender As VpNet.Core.Instance, ByVal CellX As Integer, ByVal CellZ As Integer)
    End Sub

    Function CreateMarkers() As Boolean

        If Timer2.Enabled = False Then Return True : Exit Function 'Wait until all avatars have entered first - to avoid duplicates
        If Options.EnableMapUpdates = False Then Return False : Exit Function

        Randomize()


        'Find avatars which have no marker created and create them
        For i = 1 To User.GetUpperBound(0)
            If User(i).Session <> 0 And User(i).Name <> "" And User(i).MarkerObjectID <= 0 Then
                If User(i).Name = "[Cat]" Then GoTo CatPrivilege
                If User(i).Name.Substring(0, 1) <> "[" Then 'Do not create markers for bots
CatPrivilege:       'Except cat bot (for testing)
                    'Create a marker for all of the users

                    User(i).MarkerObjectAction = "create solid no,color " & Hex(Int(Rnd() * 255)).PadRight(2, "0") & Hex(Int(Rnd() * 255)).PadRight(2, "0") & Hex(Int(Rnd() * 255)).PadRight(2, "0") & ",move {x} 0 {z} time={t} wait=9e9,rotate {r} time={rt} wait=9e9 nosync,name avmarker"

                    Dim markerObject As New VpNet.Core.Structs.VpObject
                    markerObject.Position = New VpNet.Core.Structs.Vector3(250, 0.014, -250)
                    markerObject.Rotation = New VpNet.Core.Structs.Vector3(0, 0, 0)
                    markerObject.Angle = Single.PositiveInfinity
                    markerObject.Description = User(i).Name
                    markerObject.Action = User(i).MarkerObjectAction.Replace("{x}", 0).Replace("{z}", 0).Replace("{t}", 0).Replace("{r}", 0).Replace("{rt}", 0)
                    markerObject.Model = "cyfigure.rwx"
                    markerObject.ReferenceNumber = i
                    'Console.WriteLine(User(i).Name & ": " & i)
                    Try
                        vp.AddObject(markerObject)
                        info("Added marker for " & User(i).Name)
                    Catch ex As Exception
                        info(ex.Message)
                        Return False
                    End Try
                End If
            End If
        Next
        'All markers created
        Return True

    End Function

    Private Sub vpnet_EventAvatarAdd(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.Structs.Avatar)
        'If VpConnected = False Then Exit Sub
        Dim i As Integer
        'Find space for user

        For i = 1 To User.GetUpperBound(0)
            If User(i).Online = False Then GoTo FoundSpace
        Next

        'No space found - Make a slot in array for the additional user
        ReDim Preserve User(User.GetUpperBound(0) + 1)

        i = User.GetUpperBound(0)
FoundSpace:
        User(i).Name = eventData.Name
        User(i).Session = eventData.Session
        User(i).AvatarType = eventData.AvatarType
        User(i).X = eventData.X
        User(i).Y = eventData.Y
        User(i).Z = eventData.Z
        User(i).oldX = eventData.X
        User(i).oldY = eventData.Y
        User(i).oldZ = eventData.Z
        User(i).MarkerObjectID = -1
        User(i).Id = eventData.Id
        User(i).Online = True

        If User(i).Id = 104 Or User(i).Name = "Chris D" Then Bot.Owner = User(i).Session : info("Bot owner detected. Session: " & Bot.Owner) 'Set bot owner TODO: Use code from chatlink for multiple owners

        ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "] ENTERS: " & eventData.Name & ", " & eventData.Id & ", " & eventData.Session)

        CreateMarkers()
    End Sub

    Private Sub vpnet_EventAvatarChange(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.Structs.Avatar)
        If VpConnected = False Then Exit Sub
        If ProgramIsClosing = True Then Exit Sub
        Dim i As Integer = FindUser(eventData.Session)
        'If i = 0 Or eventData.Session = 0 Or User(i).Online = False Or User(i).Name.Substring(0, 1) = "[" Then Exit Sub
        If i = 0 Or eventData.Session = 0 Or User(i).Online = False Then Exit Sub

        If User(i).X <> eventData.X Or User(i).Y <> eventData.Y Or User(i).Z <> eventData.Z Or User(i).YAW <> eventData.Yaw Then
            User(i).MovedSinceLastMarkerUpdate = True
            If User(i).X <> 0 And User(i).Z <> 0 Then User(i).statsActiveInLastHour = True 'If they've moved, set them as active in last hour
            If User(i).AvatarFrozen = True Then
                vp.TeleportAvatar(User(i).Session, "", User(i).X, User(i).Y, User(i).Z, eventData.Yaw, eventData.Pitch)
                Exit Sub
            End If
        End If


        User(i).X = eventData.X
        User(i).Y = eventData.Y
        User(i).Z = eventData.Z
        User(i).YAW = eventData.Yaw
        User(i).Pitch = eventData.Pitch
        User(i).AvatarType = eventData.AvatarType

        'Update map location
        UpdateMarker(i)

    End Sub

    Private Sub vpnet_EventAvatarDelete(ByVal sender As VpNet.Core.Instance, ByVal eventData As VpNet.Core.Structs.Avatar)
        ' If VpConnected = False Then Exit Sub
        Dim i As Integer
        i = FindUser(eventData.Session)
        If i = 0 Then Exit Sub 'User not found

        User(i).Online = False 'Stop updating of their av figure

        If User(i).MarkerObjectID > 0 Then
            info("Removing marker for " & User(i).Name)
            'Delete marker object
            Dim markerObject As New VpNet.Core.Structs.VpObject
            markerObject.Id = User(i).MarkerObjectID
            User(i).MarkerObjectID = -1
            vp.DeleteObject(markerObject)
            vp.Wait(100)
        End If
        ChatLogAppend("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "] EXITS: " & User(i).Name & ", " & User(i).Id & ", " & User(i).Session)
    End Sub



    Sub ClearUserData(ByVal i As Short)
        'Clear user data but keep their slot in array
        User(i).Session = 0
        User(i).Name = ""
        User(i).AvatarType = 0
        User(i).X = 0
        User(i).Y = 0
        User(i).Z = 0
        User(i).YAW = 0
        User(i).Pitch = 0
        User(i).Id = 0
        User(i).ListIndex = -1
        User(i).MarkerObjectID = -1
        User(i).MovedSinceLastMarkerUpdate = False
        User(i).MapClickTime = Nothing
        User(i).statsActiveInLastHour = False
        User(i).Online = False
    End Sub

    Private Sub vpnet_EventObjectClick(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal objectId As Integer, ByVal clickHitX As Single, ByVal clickHitY As Single, ByVal clickHitZ As Single)
        'TODO: neaten out

        For f = 1 To User.GetUpperBound(0)
            'User clicked an av figure
            If User(f).MarkerObjectID = objectId Then
                vpSay("User: " & User(f).Name & " (" & User(f).Id & ")", sessionId)
                Exit Sub
            End If
        Next



        Dim n As Integer
        If QueryData Is Nothing Then Exit Sub
        For n = 1 To QueryData.GetUpperBound(0)
            If QueryData(n).Id = objectId Then GoTo FoundID
        Next
        Exit Sub
FoundID:
        Dim d As Integer = FindUser(sessionId)

        'User clicked map
        If QueryData(n).Action.Contains("name mapobject") Then
            'vpSay((clickHitX - 250) * 100 & " 2 " & (clickHitZ + 250) * 100, sessionId)
            If DateTime.Now.Subtract(User(d).MapClickTime).TotalSeconds <= 2 Then
                vpSay("Teleporting to " & (clickHitX - 250) * 100 & " 2 " & (clickHitZ + 250) * 100, sessionId)
                vp.TeleportAvatar(sessionId, "", (clickHitX - 250) * 100, 2, (clickHitZ + 250) * 100, 0, 0)
                User(d).MapClickTime = Nothing
            Else
                User(d).MapClickTime = DateTime.Now 'Wait for another click within 2 seconds to classify as a "double click"
            End If
        End If
    End Sub

    Private Sub vpnet_EventObjectCreate(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal vpObject As VpNet.Core.Structs.VpObject)
        If vpObject.Action.Contains("name avmarker") Then Exit Sub 'Ignore bots own changes
        If ProgramIsClosing = True Then Exit Sub 'Program is closing
        Dim ObjectLine As String

        'PREFIXED TO THE VPPTSV1 FORMAT:
        'Type (0 = create, 1 = change, 2 = delete)
        'Session
        'User ID
        'Object ID

        If Options.EnableObjectLogging = True Then
            Dim r As Integer
            Dim userId As Integer
            r = FindUser(sessionId)
            If r = 0 Then userId = 0 Else userId = User(r).Id

            ObjectLine = "0" & vbTab & sessionId & vbTab & userId & vbTab & vpObject.Id & vbTab & vpObject.Owner & vbTab & vpObject.Time.ToString & vbTab & vpObject.Position.X & vbTab & vpObject.Position.Y & vbTab & vpObject.Position.Z & vbTab & vpObject.Rotation.X & vbTab & vpObject.Rotation.Y & vbTab & vpObject.Rotation.Z & vbTab & vpObject.Angle & vbTab & vpObject.ObjectType & vbTab & Escape(vpObject.Model) & vbTab & Escape(vpObject.Description) & vbTab & Escape(vpObject.Action)

            objWriter.WriteLine(ObjectLine)
        End If

        Exit Sub 'TODO: Mirror code (for making builds that mirror each side) is below
        For n = 1 To User.GetUpperBound(0)
            If User(n).MirrorObjectData.Id > 0 Then
                ReDim Preserve MirrorDataOriginal(MirrorDataOriginal.GetUpperBound(0) + 1)
                MirrorDataOriginal(MirrorDataOriginal.GetUpperBound(0)) = vpObject.Id

                'Now create the mirror
                'either add or subtract 180 (half of a full circle, 360 degrees),
                'and use abs()


                Exit Sub 'TODO: remove
            End If
        Next
    End Sub

    Sub ChangeMirrorObject(ByVal o As Integer, ByVal n As Integer)
        If MirrorData(o).Position.X < User(n).MirrorObjectData.Position.X Then Exit Sub
        'X> mirror only

        Dim offsetX As Integer = MirrorData(o).Position.X - User(n).MirrorObjectData.Position.X

        Dim newObject As New VpNet.Core.Structs.VpObject
        newObject = MirrorData(o)

        'newObject.Position.X = newObject.Position.X - offsetX 'TODO: Error	3	Expression is a value and therefore cannot be the target of an assignment.	E:\Documents and Settings\Chris\My Documents\Visual Studio 2010\Projects\VPUtilities\ConsoleApplication1\ConsoleApplication1\modMain.vb	540	9	ConsoleApplication1
        newObject.Rotation = New VpNet.Core.Structs.Vector3(0, 90, 0)
        Try
            vp.AddObject(newObject)
        Catch ex As Exception
            info(ex.Message)
        End Try
    End Sub
    Private Sub vpnet_EventObjectChange(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal vpObject As VpNet.Core.Structs.VpObject)
        If vpObject.Action.Contains("name avmarker") Then Exit Sub 'Ignore bots own changes
        If ProgramIsClosing = True Then Exit Sub 'Program is closing
        Dim ObjectLine As String

        'PREFIXED TO THE VPPTSV1 FORMAT:
        'Type (0 = create, 1 = change, 2 = delete)
        'Session
        'User ID
        'Object ID
        If Options.EnableObjectLogging = True Then
            Dim r As Integer
            Dim userId As Integer
            r = FindUser(sessionId)
            If r = 0 Then userId = 0 Else userId = User(r).Id

            ObjectLine = "1" & vbTab & sessionId & vbTab & userId & vbTab & vpObject.Id & vbTab & vpObject.Owner & vbTab & vpObject.Time.ToString & vbTab & vpObject.Position.X & vbTab & vpObject.Position.Y & vbTab & vpObject.Position.Z & vbTab & vpObject.Rotation.X & vbTab & vpObject.Rotation.Y & vbTab & vpObject.Rotation.Z & vbTab & vpObject.Angle & vbTab & vpObject.ObjectType & vbTab & Escape(vpObject.Model) & vbTab & Escape(vpObject.Description) & vbTab & Escape(vpObject.Action)

            objWriter.WriteLine(ObjectLine)
        End If

        Exit Sub 'TODO: Mirror object code

        'Update mirror object location 
        If vpObject.Action.Contains("name mirrorpoint") Then
            For n = 0 To User.GetUpperBound(0)
                If User(n).Id = vpObject.Owner Then User(n).MirrorObjectData = vpObject
            Next
        End If

        'Mirror objects
        For n = 0 To UBound(User)
            If User(n).MirrorObjectData.Id <> 0 Then

                For o = 1 To MirrorDataOriginal.GetUpperBound(0)
                    If MirrorDataOriginal(o) = vpObject.Id Then
                        ChangeMirrorObject(o, n)
                        '<X mirror
                    End If

                Next
            End If
        Next

    End Sub

    Private Sub vpnet_EventObjectDelete(ByVal sender As VpNet.Core.Instance, ByVal sessionId As Integer, ByVal objectId As Integer)
        For n = 1 To User.GetUpperBound(0) 'Ignore bots own changes
            If User(n).MarkerObjectID = objectId Then Exit Sub
        Next
        If ProgramIsClosing = True Then Exit Sub 'Program is closing


        Dim ObjectLine As String

        'PREFIXED TO THE VPPTSV1 FORMAT:
        'Type (0 = create, 1 = change, 2 = delete)
        'Session
        'User ID
        'Object ID
        If Options.EnableObjectLogging = True Then
            Dim r As Integer
            Dim userId As Integer
            r = FindUser(sessionId)
            If r = 0 Then userId = 0 Else userId = User(r).Id

            ObjectLine = "2" & vbTab & sessionId & vbTab & userId & vbTab & objectId & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab

            objWriter.WriteLine(ObjectLine)
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        On Error Resume Next
        '  Try
        vp.Wait(2)
        'Catch ex As Exception
        '    vp.Wait(2)
        'TODO: Figure out why this causes some kind of error occasionally
        'Console.Write("Error: vp_wait, " & ex.Message)
        'End Try

    End Sub

    Private Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        If ProgramIsClosing = True Then Exit Sub
        If VpConnected = False Then Exit Sub
        'Update marker locations
        For i As Integer = 1 To User.GetUpperBound(0)
            If User(i).MovedSinceLastMarkerUpdate = True And User(i).Session <> 0 And User(i).MarkerObjectID > 0 Then
                User(i).MovedSinceLastMarkerUpdate = False
                'If QueryData(Marker(User(i).MarkerObjectID)).Action.Substring(QueryData(Marker(User(i).MarkerObjectID)).Action.Length - 2, 2) = "no" Then QueryData(Marker(User(i).MarkerObjectID)).Action = QueryData(Marker(User(i).MarkerObjectID)).Action.Substring(0, QueryData(Marker(User(i).MarkerObjectID)).Action.Length - 2) & "yes"
                'Console.WriteLine(User(i).MarkerObjectID & " " & User(i).X & " " & User(i).Y & " " & User(i).Z)
                'Console.WriteLine(User(i).YAW)

                'TODO: What if they have no oldX/Z?


                Dim movX As Single = (User(i).X - User(i).oldX) / 15 'Controls the distance and direction
                Dim movZ As Single = (User(i).Z - User(i).oldZ) / 15 'TODO: make this more precise, could even use the yaw of the user
                Dim movTime As Single = 0.3 'Move must finish by this time; therefore controls the speed.


                Dim markerYAW As Single = User(i).YAW '-(User(i).YAW) + 180
                Dim markeroldYAW As Single = User(i).oldYAW '-(User(i).oldYAW) + 180
                Dim rotYAW As Integer = markerYAW - markeroldYAW '(markerYAW Mod 360) - (markeroldYAW Mod 360)
                '  If rotYAW < 0 Then rotYAW += 360

                Dim rotY As Single = ((Math.Abs(rotYAW) / 360) * 120) 'RPM, 500ms. 60 = 1 rotation in a second
                Dim rotTime As Single = 0.5

                If markerYAW < markeroldYAW Then rotY = -rotY 'Reverse rotation
                'Dim rotTime As Single = (rotYAW / 360)
                rotY = -rotY 'Invert because nessecary

                If markerYAW = markeroldYAW Then rotY = 0 : rotTime = 0 'Don't rotate if user isn't, may be redundant.
                Dim markerObject As New VpNet.Core.Structs.VpObject
                markerObject.Position = New VpNet.Core.Structs.Vector3(250 + (User(i).oldX / 100), 0.014, -250 + (User(i).oldZ / 100))
                'markerObject.Position = New VpNet.Core.Structs.Vector3(250 + (User(i).X / 100), 0.014, -250 + (User(i).Z / 100))
                markerObject.Rotation = New VpNet.Core.Structs.Vector3(0, -(User(i).YAW) + 180, 0)
                markerObject.Angle = Single.PositiveInfinity
                markerObject.Description = User(i).Name
                markerObject.Action = User(i).MarkerObjectAction.Replace("{x}", movX).Replace("{z}", movZ).Replace("{t}", movTime).Replace("{r}", rotY).Replace("{rt}", rotTime)
                markerObject.ReferenceNumber = -1
                markerObject.Model = "cyfigure.rwx"
                markerObject.Id = User(i).MarkerObjectID
                Try
                    vp.ChangeObject(markerObject)
                Catch ex As Exception
                    'This error could indicate connection loss
                    If Timer2.Enabled = False Then Exit Sub 'May have received a disconnect due to failed login attempt, this will be handled by the login procedure
                    If VpConnected = False Then Exit Sub
                    info("Object change error: " & ex.Message)
                    info("Connection could have been lost... reconnecting.")
                    VpConnected = False

                    Try
                        vp.Leave()
                        vp.Dispose()
                    Catch ex2 As Exception
                    End Try
                    For ib As Integer = 1 To User.GetUpperBound(0)
                        ClearUserData(ib)
                    Next

                    Timer3.Enabled = False
                    Timer2.Enabled = False
                    Timer1.Enabled = False
                    ReconnectTimer.Enabled = True
                End Try
                'Store old position
                User(i).oldX = User(i).X
                User(i).oldY = User(i).Y
                User(i).oldZ = User(i).Z
                User(i).oldYAW = User(i).YAW
nextUser:

            End If
        Next
    End Sub

    Sub UpdateMarker(ByVal i As Integer) 'Updates the map marker for an avatar

        'If nessecary, update map marker location
        If User(i).MovedSinceLastMarkerUpdate = True And User(i).Session <> 0 And User(i).MarkerObjectID > 0 Then
            User(i).MovedSinceLastMarkerUpdate = False

            Dim movX As Single = (User(i).X - User(i).oldX) / 15 'Controls the distance and direction
            Dim movZ As Single = (User(i).Z - User(i).oldZ) / 15 'TODO: make this more precise, could even use the yaw of the user
            Dim movTime As Single = 0.3 'Move must finish by this time; therefore controls the speed.


            Dim markerYAW As Single = User(i).YAW '-(User(i).YAW) + 180
            Dim markeroldYAW As Single = User(i).oldYAW '-(User(i).oldYAW) + 180
            Dim rotYAW As Integer = markerYAW - markeroldYAW '(markerYAW Mod 360) - (markeroldYAW Mod 360)
            '  If rotYAW < 0 Then rotYAW += 360

            Dim rotY As Single = ((Math.Abs(rotYAW) / 360) * 120) 'RPM, 500ms. 60 = 1 rotation in a second
            Dim rotTime As Single = 0.5

            If markerYAW < markeroldYAW Then rotY = -rotY 'Reverse rotation
            'Dim rotTime As Single = (rotYAW / 360)
            rotY = -rotY 'Invert because nessecary

            If markerYAW = markeroldYAW Then rotY = 0 : rotTime = 0 'Don't rotate if user isn't, may be redundant.
            Dim markerObject As New VpNet.Core.Structs.VpObject
            markerObject.Position = New VpNet.Core.Structs.Vector3(250 + (User(i).oldX / 100), 0.014, -250 + (User(i).oldZ / 100))
            'markerObject.Position = New VpNet.Core.Structs.Vector3(250 + (User(i).X / 100), 0.014, -250 + (User(i).Z / 100))
            markerObject.Rotation = New VpNet.Core.Structs.Vector3(0, -(User(i).YAW) + 180, 0)
            markerObject.Angle = Single.PositiveInfinity
            markerObject.Description = User(i).Name
            markerObject.Action = User(i).MarkerObjectAction.Replace("{x}", movX).Replace("{z}", movZ).Replace("{t}", movTime).Replace("{r}", rotY).Replace("{rt}", rotTime)
            markerObject.ReferenceNumber = -1
            markerObject.Model = "cyfigure.rwx"
            markerObject.Id = User(i).MarkerObjectID
            Try
                vp.ChangeObject(markerObject)
            Catch ex As Exception
                'This error could indicate connection loss
                If Timer2.Enabled = False Then Exit Sub 'May have received a disconnect due to failed login attempt, this will be handled by the login procedure
                If VpConnected = False Then Exit Sub
                info("Object change error: " & ex.Message)
                info("Connection could have been lost... reconnecting.")
                VpConnected = False

                Try
                    vp.Leave()
                    vp.Dispose()
                Catch ex2 As Exception
                End Try
                For ib As Integer = 1 To User.GetUpperBound(0)
                    ClearUserData(ib)
                Next

                Timer3.Enabled = False
                Timer2.Enabled = False
                Timer1.Enabled = False
                ReconnectTimer.Enabled = True
            End Try
            'Store old position
            User(i).oldX = User(i).X
            User(i).oldY = User(i).Y
            User(i).oldZ = User(i).Z
            User(i).oldYAW = User(i).YAW

        End If
    End Sub

    Private Sub vpnet_EventWorldDisconnect(ByVal sender As VpNet.Core.Instance)
        VpConnected = False
        If Timer2.Enabled = False Then Exit Sub 'May have received a disconnect due to failed login attempt, this will be handled by the login procedure
        info("World connection lost. Attempting to reconnect...")
        On Error Resume Next

        vp.Dispose()

        For i As Integer = 1 To User.GetUpperBound(0)
            ClearUserData(i)
        Next

        Timer3.Enabled = False
        Timer2.Enabled = False
        Timer1.Enabled = False
        ReconnectTimer.Enabled = True
    End Sub

    Private Sub vpnet_EventUniverseDisconnect(ByVal sender As VpNet.Core.Instance)
        VpConnected = False
        If Timer2.Enabled = False Then Exit Sub 'May have received a disconnect due to failed login attempt, this will be handled by the login procedure
        info("Universe connection lost. Attempting to reconnect...")
        Try
            vp.Dispose()
        Catch ex As Exception
        End Try
        For i As Integer = 1 To User.GetUpperBound(0)
            ClearUserData(i)
        Next
        Timer3.Enabled = False
        Timer2.Enabled = False
        Timer1.Enabled = False
        ReconnectTimer.Enabled = True
    End Sub
    Private Sub ReconnectTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If LoginBot() = True Then
            'Login success
            ReconnectTimer.Stop()
        Else
            Try
                'Clear up the login to avoid errors when trying again
                vp.Dispose()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        objWriter.Flush()
        If Options.EnableChatLogging = True Then objWriterChat.Flush()
    End Sub


    Private Sub vpnet_EventUserAttributes(ByVal sender As VpNet.Core.Instance, ByVal userAttributes As VpNet.Core.Structs.UserAttributes)
        ReDim Preserve UserAttribute(UserAttribute.GetUpperBound(0) + 1)
        UserAttribute(UserAttribute.GetUpperBound(0)) = userAttributes
    End Sub


    Sub UpdateWikiCitizenList()
        Try
            info("Updating wiki citizen list...")

            'Save the last wiki update time
            ConfigINI.SetKeyValue("Wiki", "CitListLastUpdate", Wiki.CitListLastUpdate.ToString(New CultureInfo("en-GB")))
            ConfigINI.Save(System.AppDomain.CurrentDomain.BaseDirectory & "Config.ini")

            Erase UserAttribute
            ReDim UserAttribute(0)

            'TODO: This could be redone in a neater way

            Dim LastNumbers1 As Integer = 0
            Dim LastNumbers2 As Integer = 100
            Dim lGrace As Byte = 3
RequestCitizens:
            'Get the citizen list
            For l As Integer = LastNumbers1 To LastNumbers2
                If l = 0 And LastNumbers1 = 0 Then GoTo NextAttribute 'Avoid querying for citizen 0
                vp.UserAttributesById(l)
                Dim curtime As DateTime = DateTime.Now
                Do
                    vp.Wait(1)
                    If DateTime.Now.Subtract(curtime).TotalSeconds >= 5 Then ReDim Preserve UserAttribute(UserAttribute.GetUpperBound(0) + 1) : lGrace = lGrace - 1 : GoTo NextAttribute
                Loop Until UserAttribute.GetUpperBound(0) = l
                'Console.WriteLine(UserAttribute(UserAttribute.GetUpperBound(0)).Name)

NextAttribute:
                If lGrace <= 0 Then GoTo UploadCitizenList 'Too many unresponsive citizens
                If l = LastNumbers2 Then Exit For 'TODO: quick fix for weird bug that makes it go to 101?
            Next

            LastNumbers1 += UserAttribute.GetUpperBound(0) + 1
            LastNumbers2 += (UserAttribute.GetUpperBound(0) + 101)
            GoTo RequestCitizens

UploadCitizenList:

            Dim outText As String = Replace(My.Resources.WikiText1a, "{date}", Wiki.CitListLastUpdate.ToString(New CultureInfo("en-GB")))

            For n = 1 To UserAttribute.GetUpperBound(0)
                If UserAttribute(n).Name = "" Then GoTo Skip3
                'Generate online time string
                Dim nTimeInSeconds As Long, nHours As Long, nMinutes As Long, nSeconds As Long, nDays As Long
                nTimeInSeconds = UserAttribute(n).OnlineTime
                nDays = nTimeInSeconds \ 86400
                nHours = (nTimeInSeconds \ 3600) - nDays * 24
                nTimeInSeconds = nTimeInSeconds Mod 3600
                nMinutes = nTimeInSeconds \ 60
                nSeconds = nTimeInSeconds Mod 60

                'Add column to output text
                If DateTime.Now.Subtract(UnixTimestampToDateTime(UserAttribute(n).RegistrationTime)).TotalDays < 30 Then
                    outText = outText & vbNewLine & "|-" & vbNewLine & "| " & UserAttribute(n).Id & " || [[User:" & UserAttribute(n).Name.Replace(" ", "_") & "|" & UserAttribute(n).Name & "]] || <span style=""color:red"">'''" & UnixTimestampToDateTime(UserAttribute(n).RegistrationTime).ToString(New CultureInfo("en-GB")) & "'''</span> || " & nDays & " day(s), " & nHours & " hour(s), " & nMinutes & " minute(s), " & nSeconds & " second(s)"
                ElseIf UserAttribute(n).Id = 104 Then 'TODO: Look into having a method so public users can link to their wiki pages if different (like mine)
                    outText = outText & vbNewLine & "|-" & vbNewLine & "| " & UserAttribute(n).Id & " || [[User:Chris|" & UserAttribute(n).Name & "]] || " & UnixTimestampToDateTime(UserAttribute(n).RegistrationTime).ToString(New CultureInfo("en-GB")) & " || " & nDays & " day(s), " & nHours & " hour(s), " & nMinutes & " minute(s), " & nSeconds & " second(s)"
                Else
                    outText = outText & vbNewLine & "|-" & vbNewLine & "| " & UserAttribute(n).Id & " || [[User:" & UserAttribute(n).Name.Replace(" ", "_") & "|" & UserAttribute(n).Name & "]] || " & UnixTimestampToDateTime(UserAttribute(n).RegistrationTime).ToString(New CultureInfo("en-GB")) & " || " & nDays & " day(s), " & nHours & " hour(s), " & nMinutes & " minute(s), " & nSeconds & " second(s)"
                End If
Skip3:
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
        Catch ex As Exception
            info("EXCEPTION: " & ex.Message)
            Exit Sub
        End Try
        info("Citizen list updated. " & Timer1.Enabled & Timer2.Enabled & Timer3.Enabled)
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
            For i As Integer = 1 To User.GetUpperBound(0)
                If User(i).Session <> 0 And User(i).Name.Length > 2 Then
                    If User(i).Name.Substring(0, 1) <> "[" And User(i).statsActiveInLastHour = False Then
                        VPStats.UserActivity += ":" & User(i).Id
                    End If
                End If
            Next
            'Active users
            VPStats.UserActivity += "|"
            For i As Integer = 1 To User.GetUpperBound(0)
                If User(i).Session <> 0 And User(i).Name.Length > 2 Then
                    If User(i).Name.Substring(0, 1) <> "[" And User(i).statsActiveInLastHour = True Then
                        VPStats.UserActivity += ":" & User(i).Id : User(i).statsActiveInLastHour = False
                    End If
                End If

            Next
            'This will still be tainted with duplicate logins, but that can be delt with when processing the stats.dat file for the actual stats.

            'Clear out the array of all the offline users each hour
            For i As Integer = 1 To User.GetUpperBound(0)
                If User(i).Online = False Then ClearUserData(i)
            Next
        Catch ex As Exception 'TODO: This can be removed once any bugs here are fully fixed
            info("EXCEPTION: " & ex.Message)
            Exit Sub
        End Try
    End Sub
    Sub SaveStatisticsLog()
        Try
            If VPStats.UserActivity = "" Then Exit Sub


            VPStats.LastHour = DateTime.UtcNow : UpdateStatisticsLog() 'Log for 23:00 - The previous update won't notice because the hour wasn't more than the previous hour - so we need to make a log for the last hour, as midnight is still


            'Save the last stats update time
            ConfigINI.SetKeyValue("Stats", "LastSave", VPStats.LastSave.ToString(New CultureInfo("en-GB")))
            ConfigINI.Save(System.AppDomain.CurrentDomain.BaseDirectory & "Config.ini")

            'Save user stats to log file
            Dim outText As String = Date.UtcNow.AddDays(-1).ToString(New CultureInfo("en-GB")) & VPStats.UserActivity
            Dim objWriter3 As System.IO.TextWriter = New System.IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory & "vpstats.dat", True)
            objWriter3.WriteLine(outText & vbNewLine)
            objWriter3.Close()

            VPStats.UserActivity = ""
            For i = 1 To User.GetUpperBound(0)
                User(i).statsActiveInLastHour = False
            Next

            info("Saved today's statistics logs.")
        Catch ex As Exception
            info("EXCEPTION: " & ex.Message)
        End Try
    End Sub


End Module
