
#Region " Import "

Imports System.IO
Imports System.IO.Compression

#End Region


Namespace Core
    Public Class ZipCore

        Class Progress(Of T)
            Implements IProgress(Of T)

            Private ReadOnly _handler As Action(Of T)

            Public Sub New(handler As Action(Of T))
                _handler = handler
            End Sub

            Private Sub Report(value As T) Implements IProgress(Of T).Report
                _handler(value)
            End Sub
        End Class

        Public Class StreamWithProgress
            Inherits Stream

            Private ReadOnly _stream As Stream
            Private ReadOnly _readProgress As IProgress(Of Integer)
            Private ReadOnly _writeProgress As IProgress(Of Integer)

            Public Sub New(Stream As Stream, readProgress As IProgress(Of Integer), writeProgress As IProgress(Of Integer))
                _stream = Stream
                _readProgress = readProgress
                _writeProgress = writeProgress
            End Sub

            Public Overrides ReadOnly Property CanRead As Boolean
                Get
                    Return _stream.CanRead
                End Get
            End Property

            Public Overrides ReadOnly Property CanSeek As Boolean
                Get
                    Return _stream.CanSeek
                End Get
            End Property

            Public Overrides ReadOnly Property CanWrite As Boolean
                Get
                    Return _stream.CanWrite
                End Get
            End Property

            Public Overrides ReadOnly Property Length As Long
                Get
                    Return _stream.Length
                End Get
            End Property

            Public Overrides Property Position As Long
                Get
                    Return _stream.Position
                End Get
                Set(value As Long)
                    _stream.Position = value
                End Set
            End Property

            Public Overrides Sub Flush()
                _stream.Flush()
            End Sub

            Public Overrides Sub SetLength(value As Long)
                _stream.SetLength(value)
            End Sub

            Public Overrides Function Seek(offset As Long, origin As SeekOrigin) As Long
                Return _stream.Seek(offset, origin)
            End Function

            Public Overrides Sub Write(buffer() As Byte, offset As Integer, count As Integer)
                _stream.Write(buffer, offset, count)
                _writeProgress.Report(count)
            End Sub

            Public Overrides Function Read(buffer() As Byte, offset As Integer, count As Integer) As Integer
                Dim bytesRead As Integer = _stream.Read(buffer, offset, count)

                _readProgress.Report(bytesRead)
                Return bytesRead
            End Function

        End Class

        NotInheritable Class ZipFileWithProgress

            'Dim sourceDirectory As String = args(0),
            '        archive As String = args(1),
            '       archiveDirectory As String = Path.GetDirectoryName(Path.GetFullPath(archive)),
            '        unpackDirectoryName As String = Guid.NewGuid().ToString()
            '
            '   File.Delete(archive)
            '    ZipFileWithProgress.CreateFromDirectory(sourceDirectory, archive,
            '       New BasicProgress(Of Double)(
            '            Sub(p)
            '                Console.WriteLine($"{p:P2} archiving complete")
            '           End Sub))
            '
            '   ZipFileWithProgress.ExtractToDirectory(archive, unpackDirectoryName,
            '       New BasicProgress(Of Double)(
            '          Sub(p)
            '               Console.WriteLine($"{p:P0} extracting complete")
            '           End Sub))

            Public Shared Sub CreateFromDirectory(
                sourceDirectoryName As String,
                destinationArchiveFileName As String,
                progress As IProgress(Of Double))

                sourceDirectoryName = Path.GetFullPath(sourceDirectoryName)

                Dim sourceFiles As FileInfo() = New DirectoryInfo(sourceDirectoryName).GetFiles("*", SearchOption.AllDirectories)
                Dim totalBytes As Double = sourceFiles.Sum(Function(f) f.Length)
                Dim currentBytes As Long = 0

                Using archive As ZipArchive = ZipFile.Open(destinationArchiveFileName, ZipArchiveMode.Create)
                    For Each fileInfo As FileInfo In sourceFiles
                        ' NOTE: naive method To Get Sub-path from file name, relative to
                        ' input directory. Production code should be more robust than this.
                        ' Either use Path class Or similar to parse directory separators And
                        ' reconstruct output file name, Or change this entire method to be
                        ' recursive so that it can follow the sub-directories And include them
                        ' in the entry name as they are processed.
                        Dim entryName As String = fileInfo.FullName.Substring(sourceDirectoryName.Length + 1)
                        Dim entry As ZipArchiveEntry = archive.CreateEntry(entryName)

                        entry.LastWriteTime = fileInfo.LastWriteTime

                        Using inputStream As Stream = File.OpenRead(fileInfo.FullName)
                            Using outputStream As Stream = entry.Open()
                                Dim progressStream As Stream = New StreamWithProgress(inputStream,
                                    New Progress(Of Integer)(
                                        Sub(i)
                                            currentBytes += i
                                            progress.Report(currentBytes / totalBytes)
                                        End Sub), Nothing)

                                progressStream.CopyTo(outputStream)
                            End Using
                        End Using
                    Next
                End Using
            End Sub

            Public Shared Sub CreateFromFiles(
                Files As FileInfo(),
                Folders As DirectoryInfo(),
                destinationArchiveFileName As String,
                progress As IProgress(Of Double))

                Dim sourceFiles As FileInfo() = Files
                Dim sourceFolders As DirectoryInfo() = Folders
                Dim totalBytes As Double = sourceFiles.Sum(Function(f) f.Length)
                Dim currentBytes As Long = 0

                Using archive As ZipArchive = ZipFile.Open(destinationArchiveFileName, ZipArchiveMode.Create)
                    For Each fileInfo As FileInfo In sourceFiles

                        Dim entryName As String = IO.Path.GetFileName(fileInfo.FullName)

                        Dim entry As ZipArchiveEntry = archive.CreateEntry(entryName)

                        entry.LastWriteTime = fileInfo.LastWriteTime

                        Using inputStream As Stream = File.OpenRead(fileInfo.FullName)
                            Using outputStream As Stream = entry.Open()
                                Dim progressStream As Stream = New StreamWithProgress(inputStream,
                                    New Progress(Of Integer)(
                                        Sub(i)
                                            currentBytes += i
                                            progress.Report(currentBytes / totalBytes)
                                        End Sub), Nothing)

                                progressStream.CopyTo(outputStream)
                            End Using
                        End Using
                    Next

                    If sourceFolders IsNot Nothing Then
                        For Each folderInfo As DirectoryInfo In sourceFolders

                            Dim FilesInFolder As List(Of FileInfo) = FileDirSearcher.GetFiles(folderInfo.FullName, SearchOption.AllDirectories).ToList

                            For Each fileInfo As FileInfo In FilesInFolder

                                Dim entryNameEx As String = IO.Path.Combine(folderInfo.Name, IO.Path.GetFileName(fileInfo.FullName))

                                Dim entryEx As ZipArchiveEntry = archive.CreateEntry(entryNameEx)

                                entryEx.LastWriteTime = fileInfo.LastWriteTime

                                Using inputStream As Stream = File.OpenRead(fileInfo.FullName)
                                    Using outputStream As Stream = entryEx.Open()
                                        Dim progressStream As Stream = New StreamWithProgress(inputStream,
                                            New Progress(Of Integer)(
                                                Sub(i)
                                                    currentBytes += i
                                                    progress.Report(currentBytes / totalBytes)
                                                End Sub), Nothing)

                                        progressStream.CopyTo(outputStream)
                                    End Using
                                End Using

                            Next

                        Next
                    End If

                End Using
            End Sub

            Public Shared Sub ExtractToDirectory(
                sourceArchiveFileName As String,
                destinationDirectoryName As String,
                progress As IProgress(Of Double))

                Using archive As ZipArchive = ZipFile.OpenRead(sourceArchiveFileName)
                    Dim totalBytes As Double = archive.Entries.Sum(Function(e) e.Length)
                    Dim currentBytes As Long = 0

                    For Each entry As ZipArchiveEntry In archive.Entries
                        Dim fileName As String = Path.Combine(destinationDirectoryName, entry.FullName)

                        Directory.CreateDirectory(Path.GetDirectoryName(fileName))
                        Using inputStream As Stream = entry.Open()
                            Using outputStream As Stream = File.OpenWrite(fileName)
                                Dim progressStream As Stream = New StreamWithProgress(outputStream, Nothing,
                                    New Progress(Of Integer)(
                                        Sub(i)
                                            currentBytes += i
                                            progress.Report(currentBytes / totalBytes)
                                        End Sub))

                                inputStream.CopyTo(progressStream)
                            End Using
                        End Using

                        File.SetLastWriteTime(fileName, entry.LastWriteTime.LocalDateTime)
                    Next
                End Using
            End Sub
        End Class

    End Class
End Namespace

