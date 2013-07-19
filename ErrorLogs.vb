Imports System.IO
Imports System.Text.RegularExpressions

Public Class ErrorLogs

    ' dgp rev 3/26/08 Parse the Source Code file
    Private Shared Function Parse_Errors(ByVal file As String) As Boolean

        Dim Code_String As String
        Dim sr As StreamReader
        Try
            sr = New StreamReader(file)
        Catch ex As Exception
            Return False
        End Try

        Code_String = sr.ReadToEnd
        sr.Close()

        Dim objMatchCollection As MatchCollection
        ' dgp rev 3/26/08 let's first remove the comment fields
        '        Code_String = Regex.Replace(Code_String, ";.*\n", "")
        objMatchCollection = Regex.Matches(Code_String, "(\%.*error.*occur.*$)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        If objMatchCollection.Count > 0 Then
            Dim item = objMatchCollection.Count
        End If
        Return objMatchCollection.Count > 0

    End Function

    Private Shared mErrorList As ArrayList
    Public Shared ReadOnly Property ErrorList() As ArrayList
        Get
            Return mErrorList
        End Get
    End Property

    ' dgp rev 9/26/08 Feedback Reporting
    Public Shared Function Scan_Files() As Boolean

        Dim item
        Dim today
        Dim fdate

        mErrorList = New ArrayList
        today = Format(Now(), "yyyyMMdd")
        For Each item In System.IO.Directory.GetFiles(Utility.MyAppPath)
            fdate = Format(System.IO.File.GetCreationTime(item), "yyyyMMdd")
            If (Parse_Errors(item)) Then mErrorList.Add(item)
            If (fdate = today) Then
                '                If (Parse_Errors(item)) Then mErrorList.Add(item)
            End If
        Next

        Scan_Files = (mErrorList.Count > 0)

    End Function

End Class
