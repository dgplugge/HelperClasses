' Name:     Color Table Manipulation Class for PV-Wave
' Author:   Donald G Plugge
' Date:     2/14/2011
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Drawing

Public Class ColorTable

    Private Shared mCustomColors As ArrayList
    Private Shared mCTPath As String = "F:\FlowRoot\Users\plugged\Settings"

    Private Shared mCustomTables As ArrayList = Nothing

    ' dgp rev 2/13/2011
    Private Shared Sub ScanTables()

        mCustomTables = New ArrayList
        mCurIndex = -1
        If Not System.IO.Directory.Exists(mCTPath) Then Return
        If System.IO.Directory.GetFiles(mCTPath, "*.ct").Length = 0 Then Return

        Dim table
        mCurIndex = 0
        For Each table In System.IO.Directory.GetFiles(mCTPath, "*.ct")
            mCustomTables.Add(System.IO.Path.GetFileNameWithoutExtension(table))
        Next

    End Sub

    Private Shared mCurIndex As Integer = -1

    ' dgp rev 2/13/2011
    Public Shared Property CurIndex As Integer
        Get
            Return mCurIndex
        End Get
        Set(ByVal value As Integer)
            mCurIndex = value
        End Set
    End Property

    Private Shared mOffSet As Int16 = 0
    Private Shared mCustomChunk() As Int32
    Private Shared mListSize As Int16

    ' dgp rev 2/14/2011 
    Private Shared Function MakeDefault() As Boolean

        If mDefaultFlag Then
            Try
                System.IO.File.Copy(mCurFile, System.IO.Path.Combine(mCTPath, "default.ct"))
                Return True
            Catch ex As Exception

            End Try
        End If
        Return False

    End Function

    Private Shared mCurFile As String

    Private Shared mChunkSize As Int16 = 16
    Private Shared mChunkOffset As Int16 = 0
    Public Shared Property ChunkOffset As Int16
        Get
            Return mChunkOffset
        End Get
        Set(ByVal value As Int16)
            mChunkOffset = value
        End Set
    End Property

    ' dgp rev 2/13/2011 read the current text file containing RGB values 
    Public Shared Function UpdateTable(ByVal intarr() As Int32, ByVal start As Int32) As Boolean

        If mCurIndex = -1 Then Return False
        If Not System.IO.Directory.Exists(mCTPath) Then Return False
        mCurFile = System.IO.Path.Combine(mCTPath, mCustomTables.Item(mCurIndex) + ".ct")

        Dim tmp_list = CurRGBList

        Try
            Dim sw As New StreamWriter(mCurFile, False)
            Dim idx
            Dim str
            Dim arr As New ArrayList
            Dim c As Color
            Dim last = start + mChunkSize - 1
            For idx = 0 To tmp_list.length - 1
                If idx >= start And idx <= last Then
                    c = System.Drawing.ColorTranslator.FromOle(intarr(idx))
                Else
                    c = System.Drawing.ColorTranslator.FromOle(tmp_list(idx))
                End If
                str = String.Format(" {0,3:D} {1,3:D} {2,3:D} ", c.R, c.G, c.B)
                sw.WriteLine(str)
            Next
            sw.Close()
            MakeDefault()
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    ' dgp rev 2/13/2011 read the current text file containing RGB values 
    Public Shared Function UpdateTable(ByVal intarr() As Int32) As Boolean

        Return UpdateTable(intarr, ChunkOffset)

    End Function

    ' dgp rev 2/13/2011 read the current text file containing RGB values 
    Private Shared Function mReadCurTable() As ArrayList

        mReadCurTable = New ArrayList
        If mCurIndex = -1 Then Exit Function
        If Not System.IO.Directory.Exists(mCTPath) Then Exit Function
        Dim tablefile As String = System.IO.Path.Combine(mCTPath, mCustomTables.Item(mCurIndex) + ".ct")
        If Not System.IO.File.Exists(tablefile) Then Exit Function

        Dim sr As New StreamReader(tablefile)
        Do While Not sr.EndOfStream
            mReadCurTable.Add(sr.ReadLine)
        Loop
        sr.Close()

    End Function

    Private Shared mRegEx As System.Text.RegularExpressions.Regex
    Private Shared mMatches As System.Text.RegularExpressions.MatchCollection

    Private Shared mDefaultFlag As Boolean = False
    Public Shared Property DefaultFlag As Boolean
        Get
            Return mDefaultFlag
        End Get
        Set(ByVal value As Boolean)
            mDefaultFlag = value
        End Set
    End Property

    ' dgp rev 2/13/2011
    Public Shared ReadOnly Property CurRGBChunk As Integer()
        Get
            Dim textlist = mReadCurTable()
            Dim express = "\s*(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})"
            Dim colorlist As New ArrayList
            Dim idx
            Dim mx = ChunkOffset + 15
            For idx = ChunkOffset To mx - 1
                mMatches = Regex.Matches(textlist(idx), express, RegexOptions.Singleline)
                If Not mMatches.Count = 0 Then
                    If mMatches.Item(0).Groups.Count = 4 Then
                        Dim byt = Microsoft.VisualBasic.RGB(mMatches.Item(0).Groups.Item(1).ToString, mMatches.Item(0).Groups.Item(2).ToString, mMatches.Item(0).Groups.Item(3).ToString)
                        colorlist.Add(byt)
                    End If
                End If
            Next

            If colorlist.Count > 0 Then
                Dim colorarr(colorlist.Count) As Int32
                Dim item
                For item = 0 To colorlist.Count - 1
                    colorarr(item) = colorlist.Item(item)
                Next
                Return colorarr
            End If
            Return Nothing

        End Get
    End Property

    ' dgp rev 2/13/2011
    Public Shared ReadOnly Property CurRGBList As Integer()
        Get
            Dim textlist = mReadCurTable()
            Dim line As String
            Dim express = "\s*(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})"
            Dim c As Color
            Dim colorlist As New ArrayList
            For Each line In textlist
                mMatches = Regex.Matches(line, express, RegexOptions.Singleline)
                If Not mMatches.Count = 0 Then
                    If mMatches.Item(0).Groups.Count = 4 Then
                        Dim byt = Microsoft.VisualBasic.RGB(mMatches.Item(0).Groups.Item(1).ToString, mMatches.Item(0).Groups.Item(2).ToString, mMatches.Item(0).Groups.Item(3).ToString)
                        colorlist.Add(byt)
                    End If
                End If
            Next

            If colorlist.Count > 0 Then
                Dim colorarr(colorlist.Count) As Int32
                Dim item
                For item = 0 To colorlist.Count - 1
                    colorarr(item) = colorlist.Item(item)
                Next
                Return colorarr
            End If
            Return Nothing

        End Get
    End Property

    ' dgp rev 2/13/2011 Any Valid Color tables
    Public Shared ReadOnly Property ColorTables As ArrayList
        Get
            If mCustomTables Is Nothing Then ScanTables()
            Return mCustomTables
        End Get
    End Property


    ' dgp rev 2/13/2011 Any Valid Color tables
    Public Shared ReadOnly Property ValidTables As Boolean
        Get
            If mCustomTables Is Nothing Then ScanTables()
            Return mCustomTables.Count > 0
        End Get
    End Property

End Class
