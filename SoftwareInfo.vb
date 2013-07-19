' Name:     SoftwareInfo
' Author:   Donald G Plugge
' Date:     9/14/2012
' Purpose:  Utility to provide help information for various software versions
' 

Imports System.IO

Public Class SoftwareInfo

    Private Shared mPackageName = VerInfo.AssemblyName
    Public Shared ReadOnly Property PackageName As String
        Get
            Return System.IO.Path.GetFileNameWithoutExtension(mPackageName)
        End Get
    End Property

    Private Shared ReadOnly Property AppServerHelpPath As String
        Get
            Return System.IO.Path.Combine(InstallInfo.GlobalHelpServerPath, PackageName)
        End Get
    End Property

    Private Shared ReadOnly Property UserServerHelpPath As String
        Get
            Return System.IO.Path.Combine(InstallInfo.UserHelpServerPath, PackageName)
        End Get
    End Property

    ' dgp rev 9/10/2012
    Public Shared Function NewHelp() As String

        NewHelp = ""
        If TextArr.Count > 0 Then
            Dim idx
            For idx = TextArr.Count - 1 To 0 Step -1
                NewHelp += TextArr(idx)
            Next
            Return NewHelp
        End If

    End Function

    Private Shared TextArr As New ArrayList

    ' dgp rev 9/10/2012
    Public Shared Function CheckHelp() As Boolean

        If SoftwareInfo.AppHelpList.Count = 0 Then Return False
        Dim item
        For Each item In SoftwareInfo.AppHelpList
            If Not SoftwareInfo.UserHelpList.Contains(item) Then
                If SoftwareInfo.UpdateHelp(item) Then
                    TextArr.Add(vbCrLf + SoftwareInfo.HelpText(item))
                End If
            End If
        Next
        Return TextArr.Count > 0

    End Function


    ' dgp rev 9/10/2012
    Public Shared Function HelpText(name As String) As String

        HelpText = ""
        If Utility.Create_Tree(UserServerHelpPath) Then
            Dim target As String = System.IO.Path.Combine(UserServerHelpPath, name)
            If Not System.IO.File.Exists(target) Then Return True
            Try
                Dim sr As New StreamReader(target)
                Dim text As String = sr.ReadToEnd()
                sr.Close()
                Return text
            Catch ex As Exception
                Return ""
            End Try
        End If

    End Function

    ' dgp rev 9/10/2012
    Public Shared Function UpdateHelp(name As String) As Boolean

        UpdateHelp = False
        If Utility.Create_Tree(UserServerHelpPath) Then
            Dim source = System.IO.Path.Combine(AppServerHelpPath, name)
            If System.IO.File.Exists(source) Then
                Dim target = System.IO.Path.Combine(UserServerHelpPath, name)
                If System.IO.File.Exists(target) Then Return True
                Try
                    System.IO.File.Copy(source, target)
                    UpdateHelp = True
                Catch ex As Exception
                    Return False
                End Try
            End If
        End If

    End Function

    ' dgp rev 9/10/2012
    Private Shared mUserHelpList = Nothing
    Public Shared ReadOnly Property UserHelpList As ArrayList
        Get
            If mUserHelpList Is Nothing Then
                mUserHelpList = New ArrayList
                If System.IO.Directory.Exists(UserServerHelpPath) Then
                    Dim fil
                    For Each fil In System.IO.Directory.GetFiles(UserServerHelpPath)
                        mUserHelpList.add(System.IO.Path.GetFileName(fil))
                    Next
                End If
            End If
            Return mUserHelpList
        End Get
    End Property

    ' dgp rev 9/10/2012
    Private Shared mAppHelpList = Nothing
    Public Shared ReadOnly Property AppHelpList As ArrayList
        Get
            If mAppHelpList Is Nothing Then
                mAppHelpList = New ArrayList
                If System.IO.Directory.Exists(AppServerHelpPath) Then
                    Dim fil
                    For Each fil In System.IO.Directory.GetFiles(AppServerHelpPath)
                        mAppHelpList.add(System.IO.Path.GetFileName(fil))
                    Next
                End If
            End If
            Return mAppHelpList
        End Get
    End Property

End Class
