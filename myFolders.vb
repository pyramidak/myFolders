#Region " Folders "

Class myFolder2

#Region " Get Files and Folders "

    Public Shared Function SubFolders(Folder As String, SystemFolder As Boolean) As String()
        Dim Dir As New IO.DirectoryInfo(Folder)
        If Dir.GetDirectories.Length = 0 Then
            Return Nothing
        Else
            Return Dir.GetDirectories.Select(Function(x) x.FullName).Where(Function(y) Exist(y, False, SystemFolder) = True).ToArray
        End If
    End Function

    Private Shared Sub LoadSubFolders(Folder As String)
        FindFiles(Folder)
        Dim Folders As String() = SubFolders(Folder, bSysFileDir)
        If bDoSubFolders = False Or Folders Is Nothing Then Exit Sub
        For Each oneSlozka As String In Folders
            LoadSubFolders(oneSlozka)
        Next
    End Sub

    Private Shared FoundFiles As New List(Of String)
    Private Shared sPatern As String
    Private Shared bDoSubFolders, bSysFileDir As Boolean

    Public Shared Function Files(Folder As String, Optional Patern As String = "*.*", Optional DoSubFolders As Boolean = False, Optional SysFileDir As Boolean = False) As String()
        FoundFiles.Clear() : sPatern = Patern : bDoSubFolders = DoSubFolders : bSysFileDir = SysFileDir
        If Exist(Folder, False, False) = False Then Return FoundFiles.ToArray
        LoadSubFolders(Folder)
        Return FoundFiles.ToArray
    End Function

    Private Shared Sub FindFiles(ByVal Folder As String)
        Dim allFiles() As String
        Dim Filters() As String = Split(sPatern, ";")
        For Each oneFilter In Filters
            allFiles = IO.Directory.GetFiles(Folder, oneFilter)
            For Each File In allFiles
                If myFile.Exist(File, bSysFileDir) Then
                    FoundFiles.Add(File)
                End If
            Next
        Next
    End Sub

#End Region

#Region " Volume serial number"

    Public Shared Function VolumeSerialNumber(ByVal Path As String) As String
        If Path = "" Then Return ""
        Try
            Dim mo As New System.Management.ManagementObject("Win32_LogicalDisk.DeviceID=""" & Path.Substring(0, 1) & ":" & """")
            Dim pd As System.Management.PropertyData = mo.Properties("VolumeSerialNumber")
            Return pd.Value.ToString()
        Catch
        End Try
        Return ""
    End Function

#End Region

#Region " Delete Empty Folders "

    Public Shared Function DeleteEmpty(ByVal Folder As String, Optional ByVal Subfolders As Boolean = False) As Boolean
        If Exist(Folder) = False Then Return False
        deleteEmptyDirs(Folder, Subfolders)
        Return True
    End Function

    Private Shared Sub deleteEmptyDirs(ByVal Folder As String, ByVal Subfolders As Boolean)
        If Subfolders Then
            For Each oneDir As String In System.IO.Directory.GetDirectories(Folder)
                If Exist(Folder) Then
                    deleteEmptyDirs(oneDir, Subfolders)
                End If
            Next
        End If
        Try
            If IO.Directory.GetDirectories(Folder).GetLength(0) = 0 And IO.Directory.GetFiles(Folder).GetLength(0) = 0 Then IO.Directory.Delete(Folder)
        Catch ex As Exception
        End Try
    End Sub

#End Region

#Region " Join "

    Private Shared Function FixPath(Folder As String) As String
        If Folder.StartsWith("\") Then Folder = Folder.Substring(1, Folder.Length - 1)
        Return Folder
    End Function

    Public Shared Function Join(Folder1 As String, Folder2 As String) As String
        Return System.IO.Path.Combine(FixPath(Folder1), FixPath(Folder2))
    End Function

    Public Shared Function Join(Folder1 As String, Folder2 As String, Folder3 As String) As String
        Return System.IO.Path.Combine(FixPath(Folder1), FixPath(Folder2), FixPath(Folder3))
    End Function

    Public Shared Function Join(Folder1 As String, Folder2 As String, Folder3 As String, Folder4 As String) As String
        Return System.IO.Path.Combine(FixPath(Folder1), FixPath(Folder2), FixPath(Folder3), FixPath(Folder4))
    End Function

#End Region

    Public Shared Function CheckAccess(Folder As String) As Boolean
        CheckAccess = False
        Dim sFile As String = Join(Folder, "access.test")
        If IO.Directory.Exists(Folder) Then
            Try
                Using fs As New IO.FileStream(sFile, IO.FileMode.CreateNew, IO.FileAccess.Write)
                    fs.WriteByte(&HFF)
                End Using
                If IO.File.Exists(sFile) Then
                    IO.File.Delete(sFile)
                    CheckAccess = True
                End If
            Catch generatedExceptionName As Exception
            End Try
        End If
    End Function

    Public Shared Function DiskType(ByVal Path As String) As Integer
        Return New System.IO.DriveInfo(Path.Substring(0, 1)).DriveType
    End Function

    Public Shared Function Name(ByVal Folder As String) As String
        If Folder = "" Then Return ""
        Folder = RemoveQuotationMarks(Folder)
        If Folder.Length < 4 Then Return Folder
        Dim Fols() As String = Split(Folder, "\".ToCharArray)
        Dim dName As String = Fols(UBound(Fols))
        Return dName
    End Function

    Public Shared Function Path(ByVal Folder As String) As String
        If Folder = "" Then Return ""
        Folder = RemoveQuotationMarks(Folder)
        If Folder.LastIndexOf("\") = -1 Or Folder.LastIndexOf("\") = 2 Then Return Folder
        Folder = Folder.Substring(0, Folder.LastIndexOf("\"))
        If Folder.Length = 2 Then Folder += "\"
        Return Folder
    End Function

    Private Shared Function RemoveQuotationMarks(ByVal Text As String) As String
        If Text = "" Then Return Text
        Dim Pos As Integer = Text.IndexOf(Chr(34))
        If Not Pos = -1 Then Text = Text.Substring(Pos + 1, Len(Text) - Pos - 1)
        Pos = Text.IndexOf(Chr(34))
        If Not Pos = -1 Then Text = Text.Substring(0, Pos)
        Pos = Text.IndexOf("//")
        If Pos = -1 Then
            Pos = Text.IndexOf("/")
            If Not Pos = -1 Then Text = Text.Substring(0, Pos - 1)
        End If
        Return Text
    End Function

    Public Shared Function Exist(ByVal Folder As String, Optional ByVal vytvorit As Boolean = False, Optional AllowSysDirs As Boolean = True) As Boolean
        If Folder = "" Then Return False
        Folder = RemoveQuotationMarks(Folder)
        Dim sName As String = Name(Folder).ToLower
        If sName.StartsWith("found.") Or sName = "perflogs" Or sName = "intel" Or sName = "system volume information" Or sName = "recycler" Or sName.Substring(0, 1) = "$" Or sName = "recycled" Or sName = "onedrivetemp" Or sName = "windows.old" Or sName = "system32" _
             Or sName = "programdata" Or sName = "msocache" Or sName.Substring(0, 1) = "." Or sName = "recovery" Or sName = "boot" Or sName = "appdata" Or sName = "intelgraphicsprofiles" Or sName = "inetpub" Then Return False
        If AllowSysDirs = False Then
            If sName = "users" Or sName = "windows" Or sName = "program files" Or sName = "program files (x86)" Or sName = "documents and settings" Then Return False
        End If
        Try
            Dim checkDir As New System.IO.DirectoryInfo(Folder)
            If checkDir.Exists = False Then
                If vytvorit Then
                    checkDir.Create()
                Else
                    Return False
                End If
            End If
            If checkDir.Root.ToString = checkDir.FullName Then Return True
            If checkDir.Attributes = 18 Or checkDir.Attributes = 19 Or checkDir.Attributes = 22 Then Return False
            checkDir.GetFiles("*.txt", IO.SearchOption.TopDirectoryOnly)
        Catch
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Create(ByVal Folder As String) As Boolean
        Folder = RemoveQuotationMarks(Folder)
        Return Exist(Folder, True)
    End Function

    Public Shared Function RemoveLastSlash(ByVal Path As String) As String
        Path = RemoveQuotationMarks(Path)
        If Path.EndsWith("\") Then
            Return Path.Substring(0, Path.Length - 1)
        Else
            Return Path
        End If
    End Function

    Public Shared Function Rename(ByVal source As String, ByVal destination As String, Optional ByVal SourceNotExistReturnTrue As Boolean = False) As Boolean
        source = RemoveLastSlash(source) : destination = RemoveLastSlash(destination)
        If Exist(source) = False Then Return SourceNotExistReturnTrue
        If source = destination Then Return False
        Dim myCopy As New clsCopyFolder(source, destination, True, Nothing)
        myCopy.Synch()
        If Exist(destination) Then
            myFile.Delete(source, True, False)
            Return True
        End If
        Return False
    End Function

    Public Shared Sub Copy(ByVal source As String, ByVal destination As String, ByVal AndSubfolders As Boolean, Optional myProgressBar As ProgressBar = Nothing)
        If Exist(source) = False Then Exit Sub
        If myProgressBar IsNot Nothing Then myProgressBar.Visibility = Visibility.Visible
        Dim myCopy As New clsCopyFolder(source, destination, AndSubfolders, myProgressBar)
        myCopy.Asynch()
    End Sub

End Class

#Region " Class CopyFolder "

Class clsCopyFolder

    Private SubFolders, wasError As Boolean
    Private CountFile, CountDir As Integer
    Private OldFolder, NewFolder As String
    Private PB As ProgressBar
    Public WithEvents thread As New System.ComponentModel.BackgroundWorker

    Sub New(ByVal SourceFolder As String, ByVal DestinationFolder As String, ByVal AndSubfolders As Boolean, ByVal myPB As ProgressBar)
        OldFolder = ClearPath(SourceFolder) : NewFolder = ClearPath(DestinationFolder) : SubFolders = AndSubfolders : wasError = False
        CountFile = 0 : CountDir = 0
        countDirs(OldFolder)
        If myPB IsNot Nothing Then
            PB = myPB
            myPB.Minimum = 0
            myPB.Value = 0
            myPB.Maximum = CountFile
        End If
        thread.WorkerReportsProgress = If(myPB Is Nothing, False, True)
    End Sub

    Public Sub Synch()
        If Not OldFolder = NewFolder And myFolder.Exist(OldFolder) Then
            If SubFolders Then
                copyDirs(OldFolder, NewFolder)
            Else
                copyFiles(OldFolder, NewFolder)
            End If
        End If
    End Sub

    Public Sub Asynch()
        thread.RunWorkerAsync()
    End Sub

    Private Sub thread_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles thread.DoWork
        Synch()
    End Sub

    Private Sub thread_ProgressChanged(sender As Object, e As ComponentModel.ProgressChangedEventArgs) Handles thread.ProgressChanged
        If PB.Value + e.ProgressPercentage <= PB.Maximum Then PB.Value += e.ProgressPercentage
    End Sub

    Private Sub thread_RunWorkerCompleted(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs) Handles thread.RunWorkerCompleted
        If thread.WorkerReportsProgress Then PB.Visibility = Visibility.Collapsed
    End Sub

    Private Sub copyDirs(ByVal oldFolder As String, ByVal newFolder As String)
        copyFiles(oldFolder, newFolder)
        For Each oneDir As String In System.IO.Directory.GetDirectories(oldFolder)
            If myFolder.Exist(oldFolder) Then
                copyDirs(oneDir, newFolder + "\" + myFolder.Name(oneDir))
            End If
        Next
    End Sub

    Private Sub copyFiles(ByVal oldFolder As String, ByVal newFolder As String)
        Dim bOnce As Boolean = False
        Dim bAlways As Boolean = False
        For Each file As IO.FileInfo In New IO.DirectoryInfo(oldFolder).GetFiles
            If thread.WorkerReportsProgress Then thread.ReportProgress(1)
            If myFile.Exist(newFolder & "\" & file.Name) Then
                If bAlways = False Then
                    Dim result As MessageBoxResult = MessageBox.Show("File already exists. Do you want to replace all (Yes), this one (No),                 none (Cancel)?" & newFolder & "\" & file.Name, Application.ProductName & " " & Application.Version, MessageBoxButton.YesNo, MessageBoxImage.Question)
                    bAlways = If(result = MessageBoxResult.Yes, True, False)
                    bOnce = If(result = MessageBoxResult.No, True, False)
                End If
                If bOnce Or bAlways Then
                    Try
                        myFile.Delete(newFolder & "\" & file.Name, False, True)
                        myFile.Copy(oldFolder & "\" & file.Name, newFolder & "\" & file.Name)
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "File access", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                    End Try
                End If
            Else
                Try
                    myFile.Copy(oldFolder & "\" & file.Name, newFolder & "\" & file.Name)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "File access", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                End Try
            End If
        Next
    End Sub

    Private Function ClearPath(ByVal FullNamePath As String) As String
        If FullNamePath.EndsWith("\") Then
            Return FullNamePath.Substring(0, FullNamePath.Length - 1)
        Else
            Return FullNamePath
        End If
    End Function

    Private Sub countDirs(ByVal oldFolder As String)
        countFiles(oldFolder)
        CountDir += System.IO.Directory.GetDirectories(oldFolder).Length
        For Each oneDir As String In System.IO.Directory.GetDirectories(oldFolder)
            countDirs(oneDir)
        Next
    End Sub
    Private Sub countFiles(ByVal oldFolder As String)
        CountFile += New IO.DirectoryInfo(oldFolder).GetFiles.Length
    End Sub
End Class

#End Region

#End Region