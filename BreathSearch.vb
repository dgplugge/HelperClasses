' Name:     BreathSearch
' Author:   Donald G Plugge
' Date:     2/13/2011
' Purpose:  Breath first search for data
Imports HelperClasses

Public Class BreathSearch

    Private Shared mSearchFolders As ArrayList
    Private Shared mSearchSubs As ArrayList
    Private Shared mMatches As ArrayList

    ' dgp rev 2/13/2011 Matches from search
    Public Shared ReadOnly Property Matches As ArrayList
        Get
            Return mMatches
        End Get
    End Property

    ' dgp rev 2/13/2011 Matches from search
    Private Shared Function RecursiveSearch() As Boolean

        mSearchFolders = mSearchSubs
        mCurDepth = mCurDepth + 1
        If mCurDepth >= mMaxDepth Then Return False
        mSearchSubs = New ArrayList
        Dim filematch
        Dim path
        Dim item
        For Each path In mSearchFolders
            Try
                For Each filematch In System.IO.Directory.GetFiles(path, mWild)
                    mMatches.Add(filematch)
                Next
            Catch ex As Exception

            End Try
            Try
                If System.IO.Directory.GetDirectories(path).Length > 0 Then
                    For Each item In System.IO.Directory.GetDirectories(path)
                        mSearchSubs.Add(item)
                    Next
                End If
            Catch ex As Exception
            End Try
        Next

        Return RecursiveSearch()

    End Function

    Private Shared mWild As String
    Private Shared mMaxDepth As Integer = 4
    Private Shared mCurDepth As Integer = 1

    ' dgp rev 2/13/2011 Search Engine
    Public Shared Function Search(ByVal wild As String) As Integer

        mWild = wild
        mSearchFolders = New ArrayList
        mSearchSubs = New ArrayList
        mMatches = New ArrayList
        Dim drv
        Dim fld
        mCurDepth = 0
        ' prime the subsearch folder
        For Each drv In HelperClasses.Utility.AllDrives
            For Each fld In System.IO.Directory.GetDirectories(drv)
                mSearchSubs.Add(fld)
            Next
        Next

        RecursiveSearch()

        Return mMatches.Count

    End Function

End Class
