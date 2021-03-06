Imports System.Globalization
Imports System.Collections.Generic
Imports VpNet.Core.Structs

Module modStuff
    Enum connectionState
        Disconnected
        Connected
    End Enum
    Class objBot
        Public Instance As VpNet.Core.Instance
        Public LoginName As String
        Public UniHost As String
        Public UniPort As Integer
        Public CitName As String
        Public CitPass As String
        Public WorldName As String
        ' Public ChatPath As String
        Public GreetUsers As Boolean
        Public MyX As Single
        Public MyY As Single
        Public MyZ As Single
        Public MyYAW As Single
        Public lowDistance
        Public EnterMessage As String
        Public ExitMessage As String
        Public EnterTime As DateTime
        Public Owner As Integer
        ' Public Connected As Boolean
    End Class
    Class objUser
        Public Session As Integer
        Public Name As String
        Public Id As Integer
        Public AvatarType As Integer
        Public X As Single
        Public Y As Single
        Public Z As Single
        Public Yaw As Single
        Public oldX As Single
        Public oldY As Single
        Public oldZ As Single
        Public oldYAW As Single
        Public Pitch As Single
        Public ListIndex As Integer 'User Listbox Index
        Public MarkerPending As Boolean
        Public MarkerObjectID As Integer
        Public MarkerObjectAction As String
        Public MirrorObjectData As VpNet.Core.Structs.VpObject
        Public MovedSinceLastMarkerUpdate As Boolean
        Public MapClickTime As DateTime
        Public LastActive As DateTime
        Public AvatarFrozen As Boolean
        Public statsActiveInLastHour As Boolean
        'format: Datetime,hour | activenumber | totalnumber,hour | activenumber | totalnumber,hour | activenumber | totalnumber, [..repeats until next day, POSSIBLE to have each date on a newline]
        'csv, with | separation (allows extra values to be added later on, for each hour
        'example: 23/04/2013,1|3|5,2|4|5,3|1|4,4|2|3,5|2|2,5|2|0,6|2|0 [new line for each day]
    End Class
    Structure structOptions
        Dim EnableStatisticsLogging As Boolean
        Dim EnableObjectLogging As Boolean
        Dim EnableChatLogging As Boolean
        Dim EnableMapUpdates As Boolean
        Dim EnableWikiUpdates As Boolean
    End Structure
    Structure structWorld
        Dim Name As String
        Dim UserCount As Short
        Dim State As Short
    End Structure
    Structure structWorldAttribs
        Dim Key As String
        Dim Value As String
    End Structure
    Structure structStatistics
        Dim LastSave As DateTime
        Dim LastHour As DateTime
        Dim UserActivity As String
        '  Dim ObjectStatsBuffer As String
    End Structure
    Class stateWiki
        Public CitListLastUpdate As DateTime
        Public CitListPolling As Boolean
        Public CitListPollTarget As Integer
        Public Username As String
        Public Password As String
    End Class
    Public ConfigINI As New IniFile

    Public Bot As objBot = New objBot
    Public Wiki As stateWiki = New stateWiki
    Public Users As List(Of objUser) = New List(Of objUser)
    Public Citizens As List(Of UserAttributes) = New List(Of UserAttributes)
    Public Queries As List(Of VpObject) = New List(Of VpObject)
    Public Groups As List(Of VpObject) = New List(Of VpObject)
    Public Mirrors As List(Of VpObject) = New List(Of VpObject)

    Dim LastWriteLine As String
    Public Marker(31) As Integer 'Contains the querydata array index of each marker
    Public Options As structOptions
    Public VPStats As structStatistics
    'Public World() As structWorld
    'Public WorldAttributes() As structWorldAttribs

    Function Val2Bool(ByVal Valu As Byte) As Boolean
        If Valu = 1 Then Return True
        If Valu = 0 Then Return False
        Return False
    End Function

    Function Bool2Val(ByVal Boole As Boolean) As Byte
        If Boole = True Then Return 1
        If Boole = False Then Return 0
        Return 0
    End Function

    Function FindUser(ByVal UserSession As Long) As objUser
        For Each User In Users
            If User.Session = UserSession Then Return User
        Next

        Return Nothing
    End Function

    Function FindUserByName(ByVal UserName As String) As objUser
        For Each User In Users
            If User.Name.ToLower = UserName.ToLower Then Return User
        Next

        Return Nothing
    End Function


    Sub vpSay(ByVal Message As String, Optional ByVal Session As Integer = 0)
        Bot.Instance.ConsoleMessage(Session, "", "� " & Message, VpNet.Core.EventData.VPTextEffect.TextEffectItalic, 255, 0, 0)
    End Sub

    Public Function DateTimeToUnixTimeStamp(ByVal currDate As DateTime) As Long
        'create Timespan by subtracting the value provided from the Unix Epoch
        Dim span As TimeSpan = (currDate - New DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime())
        'return the total seconds (which is a UNIX timestamp)
        Return span.TotalSeconds
    End Function

    Public Function UnixTimestampToDateTime(ByVal UnixTimeStamp As Long) As DateTime
        Return (New DateTime(1970, 1, 1, 0, 0, 0)).AddSeconds(UnixTimeStamp)
    End Function

    Public Function Escape(ByVal InputText As String) As String
        'This performs a minimal escape for newlines and characters that text editors may not display properly
        'Reference: http://msdn.microsoft.com/en-us/library/6aw8xdf2.aspx
        InputText = Replace(InputText, Chr(0), "\0")
        InputText = Replace(InputText, "\", "\\")
        InputText = Replace(InputText, Chr(9), "\t")
        InputText = Replace(InputText, Chr(13), "\r")
        InputText = Replace(InputText, Chr(10), "\n")
        InputText = Replace(InputText, Chr(11), "\v")
        InputText = Replace(InputText, Chr(8), "\b") '
        InputText = Replace(InputText, Chr(13), "\r")
        InputText = Replace(InputText, Chr(12), "\f")
        Return InputText
    End Function

    Public Function Unescape(ByVal InputText As String) As String
        'Reverse of above
        Try
            InputText = System.Text.RegularExpressions.Regex.Unescape(InputText)
        Catch ex As Exception
            Return ""
        End Try
        Return InputText
    End Function

    Public Sub ChatLogAppend(ByVal Text As String)
        If Options.EnableChatLogging = True Then modMain.objWriterChat.WriteLine(Text)
    End Sub

    Public Sub info(ByVal Text As String)
        Console.WriteLine("[" & DateTime.UtcNow.ToString(New CultureInfo("en-GB")) & "] " & Text)
    End Sub

    Public Function D2R(ByVal Angle As Single) As Single
        D2R = Angle / 180 * Math.PI
    End Function

    Public Function R2D(ByVal Angle As Single) As Single
        R2D = Angle * 180 / Math.PI
    End Function
End Module
