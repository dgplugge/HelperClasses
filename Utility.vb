' Name:     FCS File Class
' Author:   Donald G Plugge
' Date:     07/17/07
' Purpose:  
Imports System.IO
Imports System.Text.RegularExpressions

Public Class Utility

    Private Shared mMyAppPath

    Private Shared mAllDrives As ArrayList

    Public Delegate Sub CopyEventHandler(ByVal File As String)
    Public Shared Event CopyEvent As CopyEventHandler

    Public Delegate Sub CopyErrorEventHandler(ByVal File As String)
    Public Shared Event CopyErrorEvent As CopyErrorEventHandler

    ' dgp rev 5/11/09 Scan for Local Fixed drives
    Private Shared Sub ScanDrives()
        mDrives = New ArrayList
        mAllDrives = New ArrayList

        ' dgp rev 2/19/08 scan drives for valid locations
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim drv As DriveInfo

        Dim maxdrive As Long = 0

        For Each drv In allDrives

            If drv.IsReady Then
                If (Not drv.DriveType = DriveType.Removable And drv.DriveType = DriveType.Fixed) Then
                    mDrives.Add(drv.Name)
                    If drv.AvailableFreeSpace > maxdrive Then
                        maxdrive = drv.AvailableFreeSpace
                        mLargeDrive = drv.Name.ToUpper
                    End If
                End If
                If (Not drv.DriveType = DriveType.Removable And (drv.DriveType = DriveType.Fixed Or drv.DriveType = DriveType.Network)) Then
                    mAllDrives.Add(drv.Name)
                End If
            End If
        Next

    End Sub

    ' dgp rev 12/21/2010
    Public Shared ReadOnly Property AllDrives()
        Get
            If mAllDrives Is Nothing Then ScanDrives()
            Return mAllDrives
        End Get
    End Property

    Public Shared ReadOnly Property MyAppPath() As String
        Get
            If mMyAppPath Is Nothing Then
                If (System.Environment.GetEnvironmentVariable("APPDATA") Is Nothing) Then
                    If (System.Environment.GetEnvironmentVariable("USERPROFILE") Is Nothing) Then
                        mMyAppPath = My.Computer.FileSystem.CurrentDirectory
                    Else
                        mMyAppPath = System.IO.Path.Combine(System.Environment.GetEnvironmentVariable("USERPROFILE"), "Application Data")
                    End If
                Else
                    mMyAppPath = System.Environment.GetEnvironmentVariable("APPDATA")
                End If
                mMyAppPath = System.IO.Path.Combine(mMyAppPath, "Flow Control")
            End If
            Return mMyAppPath
        End Get
    End Property

    ' dgp rev 4/27/2010 recursive tree delete, replacement for system.io.directory.delete
    Public Shared Function DeleteTree(ByVal path As String) As Boolean

        If Not System.IO.Directory.Exists(path) Then Return True

        Dim file
        For Each file In System.IO.Directory.GetFiles(path)
            System.IO.File.Delete(file)
        Next
        Dim fld
        For Each fld In System.IO.Directory.GetDirectories(path)
            DeleteTree(fld)
        Next
        Try
            System.IO.Directory.Delete(path)
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function



    ' dgp rev 4/12/07 Create a Unique Name
    Public Shared Function Unique_Name() As String

        Return Format(Now(), "yyyyMMddhhmmss")

    End Function

    Private Shared mDrives As ArrayList
    ' dgp rev 7/20/07 
    Public Shared ReadOnly Property LocalDrives()
        Get
            If mDrives Is Nothing Then ScanDrives()
            Return mDrives
        End Get
    End Property

    Public Shared Sub DirectoryCopy( _
            ByVal sourceDirName As String, _
            ByVal destDirName As String)

        DirectoryCopy(sourceDirName, destDirName, True)

    End Sub

    Public Shared Sub DirectoryCopy( _
            ByVal sourceDirName As String, _
            ByVal destDirName As String, _
            ByVal copySubDirs As Boolean)

        Dim dir As DirectoryInfo = New DirectoryInfo(sourceDirName)
        Dim dirs As DirectoryInfo() = dir.GetDirectories()

        ' If the source directory does not exist, throw an exception.
        If Not dir.Exists Then
            Throw New DirectoryNotFoundException( _
                "Source directory does not exist or could not be found: " _
                + sourceDirName)
        End If

        ' If the destination directory does not exist, create it.
        If Not Directory.Exists(destDirName) Then
            Directory.CreateDirectory(destDirName)
        End If

        ' Get the file contents of the directory to copy.
        Dim files As FileInfo() = dir.GetFiles()

        Dim file
        For Each file In files

            ' Create the path to the new copy of the file.
            Dim temppath As String = Path.Combine(destDirName, file.Name)

            ' Copy the file.
            Try
                file.CopyTo(temppath, False)
                RaiseEvent CopyEvent(temppath)
            Catch ex As Exception
                RaiseEvent CopyErrorEvent(temppath)
            End Try
        Next file

        Dim subdir
        ' If copySubDirs is true, copy the subdirectories.
        If copySubDirs Then

            For Each subdir In dirs

                ' Create the subdirectory.
                Dim temppath As String = _
                    Path.Combine(destDirName, subdir.Name)

                ' Copy the subdirectories.
                DirectoryCopy(subdir.FullName, temppath, copySubDirs)
            Next subdir
        End If
    End Sub

    Private Shared mLargeDrive As String
    Public Shared ReadOnly Property LargestDrive() As String
        Get
            If mLargeDrive Is Nothing Then ScanDrives()
            Return mLargeDrive
        End Get
    End Property


    ' dgp rev 11/8/06 create a folder with error handling
    ' dgp rev 11/13/06 parent folder must exist
    Private Shared Function create_folder(ByVal strPath As String)

        create_folder = False ' assume failure
        If (Not System.IO.Directory.Exists(strPath)) Then
            On Error GoTo Create_Error
            If (System.IO.Directory.Exists(System.IO.Directory.GetParent(strPath).Parent.FullName)) Then
                System.IO.Directory.CreateDirectory(strPath)
            End If
        End If
        create_folder = True
        Exit Function

Create_Error:
        MsgBox("Unable to create folder -- " + strPath, vbExclamation, "Folder Creation Error")

    End Function

    ' dgp rev 11/30/2010 folder size
    Public Shared Function FolderSize(ByVal sPath As String, ByVal bRecursive As Boolean) As Long

        Dim Size As Long = 0
        Dim diDir As New DirectoryInfo(sPath)

        Try

            Dim fil As FileInfo
            For Each fil In diDir.GetFiles()
                Size += fil.Length
            Next fil
            If bRecursive = True Then

                Dim diSubDir As DirectoryInfo
                For Each diSubDir In diDir.GetDirectories()
                    Size += FolderSize(diSubDir.FullName, True)
                Next diSubDir
            End If

            Return Size
        Catch ex As System.IO.FileNotFoundException
            ' File not found. Take no action

        Catch exx As Exception
            ' Another error occurred

            Return 0
        End Try

    End Function


    ' dgp rev 11/27/06 once it is determined that a path is to be created, back up
    ' to the first subdirectory that exists.
    Public Shared Function Create_Tree(ByVal path_str As String) As Boolean

        Dim sep As Char = System.IO.Path.DirectorySeparatorChar
        If path_str.Length = 0 Then Return True

        If (System.IO.Directory.Exists(path_str)) Then Return True
        Dim split_arr() As String
        Dim offset = 0
        If (path_str.Substring(0, 2).Contains(sep + sep)) Then
            ' dgp rev 2/18/09 remote path
            split_arr = path_str.Split(sep)
            If (split_arr.Length < 4) Then Return False
            Dim root = sep + sep + split_arr(2) + sep + split_arr(3)
            If (Not System.IO.Directory.Exists(root)) Then Return False
            split_arr(3) = root
            offset = 3
        Else
            split_arr = path_str.Split(sep)
        End If

        Dim test_path As String

        Dim idx
        test_path = split_arr(offset) + sep
        If (Not System.IO.Directory.Exists(test_path)) Then Return False

        For idx = 1 + offset To split_arr.Length - 1
            test_path = test_path + split_arr(idx) + sep
            If (Not System.IO.Directory.Exists(test_path)) Then create_folder(test_path)
        Next

        If (System.IO.Directory.Exists(path_str)) Then Return True

    End Function

End Class
