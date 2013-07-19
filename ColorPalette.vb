' Name:     Color Palette Manipulation Class for PV-Wave
' Author:   Donald G Plugge
' Date:     2/14/2011
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Drawing

Public Class ColorPalette

    Private Shared mCTPath As String = "F:\FlowRoot\Users\plugged\Settings"
    Private Shared mCustomPalettes As ArrayList = Nothing

    Public Shared Sub ResetPalettes()

        mCustomPalettes = Nothing

    End Sub

    ' dgp rev 2/13/2011
    Public Shared Sub ScanPalettes()

        mCustomPalettes = New ArrayList
        If Not System.IO.Directory.Exists(mCTPath) Then Return
        If System.IO.Directory.GetFiles(mCTPath, "*.ct").Length = 0 Then Return

        Dim palette
        For Each palette In System.IO.Directory.GetFiles(mCTPath, "*.ct")
            mCustomPalettes.Add(System.IO.Path.GetFileNameWithoutExtension(palette))
        Next

    End Sub

    Private Shared mCurIndex As Integer = -1

    ' dgp rev 2/13/2011
    Public Shared Property CurIndex As Integer
        Get
            Return mCurIndex
        End Get
        Set(ByVal value As Integer)
            If Not value = -1 Then
                If Not mCurIndex = value Then
                    mCurIndex = value
                    ReadCurPalette()
                End If
            End If
        End Set
    End Property

    ' dgp rev 2/14/2011 
    Private Shared Function MakeDefault() As Boolean

        If mDefaultFlag Then
            Try
                System.IO.File.Copy(mCurFileSpec, System.IO.Path.Combine(mCTPath, "default.ct"))
                Return True
            Catch ex As Exception

            End Try
        End If
        Return False

    End Function

    Private Shared mCurFileSpec As String

    Private Shared mPageSize As Int16 = 16
    Public Shared Property PageSize As Int16
        Get
            Return mPageSize
        End Get
        Set(ByVal value As Int16)
            mPageSize = value
        End Set
    End Property


    Public Shared ReadOnly Property PageCount As Int16
        Get
            If ColorArray.Count = 0 Then Return 0
            Return Math.Ceiling(ColorArray.Count / PageSize)
        End Get
    End Property

    Private Shared mPageNumber As Int16 = 1
    Public Shared Property PageNumber As Int16
        Get
            Return mPageNumber
        End Get
        Set(ByVal value As Int16)
            If value > 0 Then
                If value <= PageCount Then mPageNumber = value
            End If
        End Set
    End Property

    ' dgp rev 2/25/2011 Save the current Palette
    Public Shared Function SavePalette() As Boolean

        Return SavePalette(mCurFilename)

    End Function

    ' dgp rev 2/25/2011 Save the current Palette
    Public Shared Function RemovePalette() As Boolean

        Return RemovePalette(mCustomPalettes.Item(CurIndex))

    End Function

    ' dgp rev 2/25/2011 Save the current Palette
    Public Shared Function RemovePalette(ByVal name As String) As Boolean

        If Not System.IO.Directory.Exists(mCTPath) Then Return False
        mCurFileSpec = System.IO.Path.Combine(mCTPath, name + ".ct")
        Try
            If System.IO.File.Exists(mCurFileSpec) Then System.IO.File.Delete(mCurFileSpec)
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    ' dgp rev 2/25/2011 Save the current Palette
    Public Shared Function SavePalette(ByVal name As String) As Boolean

        If Not System.IO.Directory.Exists(mCTPath) Then Return False
        mCurFileSpec = System.IO.Path.Combine(mCTPath, name + ".ct")
        Dim sw As StreamWriter
        Try
            Dim str
            Dim arr As New ArrayList
            Dim c As Color
            sw = New StreamWriter(mCurFileSpec, False)

            For Each c In ColorArray
                str = String.Format(" {0,3:D} {1,3:D} {2,3:D} ", c.R, c.G, c.B)
                sw.WriteLine(str)
            Next
            sw.Close()
            MakeDefault()
        Catch ex As Exception
            sw.Close()
            Return False
        End Try
        Return True

    End Function

    ' dgp rev 2/24/2011 Validate the current filename
    Private Shared Function bIsValidFileName(ByVal filename As String) As Boolean

        Dim i As Integer

        ' Validate
        bIsValidFileName = True
        ' Can't be empty
        If String.IsNullOrEmpty(filename) Then
            bIsValidFileName = False
        End If
        ' Check invalid characters
        For i = 0 To Path.GetInvalidFileNameChars.GetUpperBound(0)
            If filename.IndexOf(Path.GetInvalidFileNameChars(i)) >= 0 Then
                ' Illegal character
                bIsValidFileName = False
                Exit For
            End If
        Next i

    End Function

    Private Shared mCurFilename As String
    Public Shared Function CheckFilename(ByVal name) As Boolean

        If Not bIsValidFileName(name) Then Return False
        Return True

    End Function


    Public Shared Function ChangeFilename(ByVal name) As Boolean

        If Not bIsValidFileName(name) Then Return False
        mCurFilename = name
        Return True

    End Function

    Private Shared mValid As Boolean = False
    Public Shared ReadOnly Property Valid As Boolean
        Get
            Return mValid
        End Get
    End Property

    ' dgp rev 2/13/2011 read the current text file containing RGB values 
    Private Shared Sub ReadCurPalette()

        mColorArray = New ArrayList
        Dim PaletteText = New ArrayList
        If CurIndex = -1 Then Exit Sub
        If Not System.IO.Directory.Exists(mCTPath) Then Exit Sub
        mCurFileSpec = System.IO.Path.Combine(mCTPath, mCustomPalettes.Item(CurIndex) + ".ct")
        If Not System.IO.File.Exists(mCurFileSpec) Then Exit Sub

        Try
            Dim sr As New StreamReader(mCurFileSpec)
            Do While Not sr.EndOfStream
                PaletteText.Add(sr.ReadLine)
            Loop
            mValid = PaletteText.Count > 0
            sr.Close()
            mCurFilename = System.IO.Path.GetFileNameWithoutExtension(mCustomPalettes.Item(CurIndex))
        Catch ex As Exception
            mValid = False
        End Try

        Dim express = "\s*(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})"
        Dim idx
        For idx = 0 To PaletteText.Count - 1
            mMatches = Regex.Matches(PaletteText(idx), express, RegexOptions.Singleline)
            If Not mMatches.Count = 0 Then
                If mMatches.Item(0).Groups.Count = 4 Then
                    Dim byt = Color.FromArgb(mMatches.Item(0).Groups.Item(1).ToString, mMatches.Item(0).Groups.Item(2).ToString, mMatches.Item(0).Groups.Item(3).ToString)
                    mColorArray.Add(byt)
                End If
            End If
        Next

    End Sub

    Private Shared mRegEx As System.Text.RegularExpressions.Regex
    Private Shared mMatches As System.Text.RegularExpressions.MatchCollection

    Private Shared mDefaultFlag As Boolean = False
    ' dgp rev 2/23/2011
    Public Shared Property DefaultFlag As Boolean
        Get
            Return mDefaultFlag
        End Get
        Set(ByVal value As Boolean)
            mDefaultFlag = value
        End Set
    End Property

    Private Shared mColorArray As ArrayList
    Public Shared Property ColorArray As ArrayList
        Get
            Return mColorArray
        End Get
        Set(ByVal value As ArrayList)
            mColorArray = value
        End Set
    End Property

    ' dgp rev 2/24/2011 Modify a single color in the current palette
    Public Shared Function TrimColors(ByVal idx As Integer) As Boolean

        If idx < 0 Then Return False
        If idx >= ColorArray.Count Then Return False
        Dim item
        Dim tmp As New ArrayList
        Dim cnt = 0
        Dim keep = ColorArray.Count - idx
        For Each item In ColorArray
            tmp.Add(item)
            cnt += 1
            If cnt = keep Then Exit For
        Next
        ColorArray = tmp
        Return True

    End Function

    ' dgp rev 2/24/2011 Modify a single color in the current palette
    Public Shared Function ModifyColor(ByVal idx As Integer, ByVal c As Color) As Boolean

        If idx < 0 Then Return False
        If idx >= ColorArray.Count Then Return False
        ColorArray(idx) = c
        Return True

    End Function

    ' dgp rev 2/23/2011
    Public Shared ReadOnly Property CurColorPage As Color()
        Get

            If ColorArray.Count > 0 Then
                If PageNumber = 0 Then PageNumber = 1
                Dim mx = (PageNumber * mPageSize)
                If ColorArray.Count < mx Then mx = ColorArray.Count
                Dim start As Integer = (PageNumber - 1) * mPageSize
                Dim colorarr(mx - start - 1) As Color
                Dim pos
                Dim idx = 0
                For pos = start To mx - 1
                    colorarr(idx) = mColorArray.Item(pos)
                    idx += 1
                Next
                Return colorarr
            End If
            Return Nothing

        End Get
    End Property


    ' dgp rev 2/13/2011 Any Valid Color palettes
    Public Shared ReadOnly Property ColorPalettes As ArrayList
        Get
            If mCustomPalettes Is Nothing Then ScanPalettes()
            Return mCustomPalettes
        End Get
    End Property

    ' dgp rev 2/13/2011 Any Valid Color palettes
    Public Shared ReadOnly Property ValidPalettes As Boolean
        Get
            If mCustomPalettes Is Nothing Then ScanPalettes()
            Return mCustomPalettes.Count > 0
        End Get
    End Property

    Shared Sub AppendColor(ByVal color As Color)

        mColorArray.Add(color)

    End Sub

End Class
