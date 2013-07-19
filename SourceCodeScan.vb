Public Class SourceCodeScan

    Public Delegate Sub FindEventHandler(ByVal File As String)
    Public Shared Event FindEvent As FindEventHandler
    Public Delegate Sub DoneEventHandler()
    Public Shared Event DoneEvent As DoneEventHandler

    Private Shared Sub ScanCodeRoot()

        Dim drv
        For Each drv In Utility.LocalDrives
            Dim path = System.IO.Path.Combine(drv, "MyCode")
            If System.IO.Directory.Exists(path) Then mMyCodeRoot = path
        Next

    End Sub

    Private Shared mMyCodeRoot = Nothing
    Public Shared Property MyCodeRoot As String
        Get
            If mMyCodeRoot Is Nothing Then ScanCodeRoot()
            Return mMyCodeRoot
        End Get
        Set(value As String)
            mMyCodeRoot = value
        End Set
    End Property

    Private Shared mUniqueFilename As Hashtable

    Private Shared Sub ScanWildSource()

    End Sub

    Private Shared mDateStart As DateTime = Now.AddDays(-7)
    Public Shared Property DateStart As DateTime
        Get
            Return mDateStart
        End Get
        Set(value As DateTime)
            mDateStart = value
        End Set
    End Property

    Private Shared mDateEnd As DateTime = Now
    Public Shared Property DateEnd As DateTime
        Get
            Return mDateEnd
        End Get
        Set(value As DateTime)
            mDateEnd = value
        End Set
    End Property

    ' dgp rev 2/25/2013
    Private Shared mWildSourcePathList = Nothing
    Public Shared ReadOnly Property WildSourcePathList As ArrayList
        Get
            If mWildSourcePathList Is Nothing Then ScanWildSource()
            Return mWildSourcePathList
        End Get
    End Property

    ' dgp rev 2/25/2013
    Private Shared Sub FileFound(File As String)

        If Not File.Length - 1 = File.Replace(".", "").Length Then
            Exit Sub
        End If

        mCompleteList.Add(File)
        Dim name = System.IO.Path.GetFileName(File)
        Dim filedate = System.IO.File.GetLastWriteTime(File)
        If mUniqueFilename.ContainsKey(name) Then
            Dim info As Hashtable = mUniqueFilename(name)
            Dim found = False
            If Not info.ContainsKey(filedate) Then
                info.Add(filedate, File)
                mUniqueFilename(name) = info
                RaiseEvent FindEvent(File)
            End If
        Else
            Dim newitem As New Hashtable
            newitem.Add(filedate, File)
            mUniqueFilename.Add(name, newitem)
            RaiseEvent FindEvent(File)
        End If

    End Sub

    Private Shared mCompleteList As ArrayList
    Private Shared mUniqueDates As ArrayList
    Private Shared mWildName As String = "*"
    Private Shared mScanExt As String = "vb"

    Public Shared Property ScanExt As String
        Get
            Return mScanExt
        End Get
        Set(value As String)
            mScanExt = value
        End Set
    End Property

    Public Shared Property WildName As String
        Get
            Return mWildName
        End Get
        Set(value As String)
            mWildName = value
        End Set
    End Property



    ' dgp rev 2/25/2013 
    Public Shared Sub RecursiveUniqueScan(path As String)

        Dim vbfil
        For Each vbfil In System.IO.Directory.GetFiles(path, String.Format("{0}.{1}", mWildName, mScanExt))
            FileFound(vbfil)
        Next

        If System.IO.Directory.GetDirectories(path).Length = 0 Then Exit Sub
        For Each item In System.IO.Directory.GetDirectories(path)
            RecursiveUniqueScan(item)
        Next

    End Sub

    ' dgp rev 2/25/2013 
    Public Shared Sub RecursiveScan(path As String)

        Dim vbfil
        For Each vbfil In System.IO.Directory.GetFiles(path, String.Format("{0}.{1}", mWildName, mScanExt))
            FileFound(vbfil)
        Next

        If System.IO.Directory.GetDirectories(path).Length = 0 Then Exit Sub
        For Each item In System.IO.Directory.GetDirectories(path)
            RecursiveScan(item)
        Next

    End Sub

    Private Shared mEarlierDateList As Hashtable
    Private Shared mLaterDateList As Hashtable
    Private Shared mEarlierDatePathList As Hashtable
    Private Shared mLaterDatePathList As Hashtable

    Private Shared mPreDateList As Hashtable
    Private Shared mPrePathList As Hashtable

    Private Shared Sub FileFoundCheckDate(File As String)

        Dim name = System.IO.Path.GetFileName(File)
        Dim filedate = System.IO.File.GetLastWriteTime(File)

        If mDateLimit > filedate Then
            If mPreDateList.ContainsKey(name) Then
                If filedate > mPreDateList(name) Then
                    mPreDateList(name) = filedate
                    mPrePathList(name) = File
                End If
            Else
                mPreDateList.Add(name, filedate)
                mPrePathList.Add(name, File)
            End If
        End If

        If filedate > mDateLimit Then
            If mFirstDateList.Contains(name) Then
                Dim hold = mFirstDateList.Item(name)
                Dim holdpath = mFirstDatePathList.Item(name)
                If filedate > mFirstDateList.Item(name) Then
                    mFirstDateList.Item(name) = filedate
                    mFirstDatePathList.Item(name) = File
                End If
                If mNextDateList.Contains(name) Then
                    If hold > mNextDateList.Item(name) Then
                        If Not hold = mFirstDateList(name) Then
                            mNextDateList.Item(name) = hold
                            mNextDatePathList.Item(name) = File
                        End If
                    End If
                Else
                    If mPreDateList.ContainsKey(name) Then
                        mNextDateList.Add(name, mPreDateList(name))
                        mNextDatePathList.Add(name, mPrePathList(name))
                    End If
                End If
            Else
                mFirstDateList.Add(name, filedate)
                mFirstDatePathList.Add(name, File)
                If mPreDateList.ContainsKey(name) Then
                    mNextDateList.Add(name, mPreDateList(name))
                    mNextDatePathList.Add(name, mPrePathList(name))
                End If
            End If
        End If
    End Sub

    Private Shared Sub FileFoundTwoDate(File As String)

        Dim name = System.IO.Path.GetFileName(File)
        Dim filedate = System.IO.File.GetLastWriteTime(File)

        If mEarlierDateList.Contains(name) Then
            If mLaterDateList.Contains(name) Then
                If filedate > mEarlierDateList(name) Then
                    ' file later than earliest
                    If filedate > mLaterDateList(name) Then
                        mEarlierDateList(name) = mLaterDateList(name)
                        mEarlierDatePathList(name) = mLaterDatePathList(name)
                        mLaterDateList(name) = filedate
                        mLaterDatePathList(name) = File
                    Else
                        If Not filedate = mLaterDateList(name) Then
                            mEarlierDateList(name) = filedate
                            mEarlierDatePathList(name) = File
                        End If
                    End If
                End If
            Else
                If filedate > mEarlierDateList(name) Then
                    ' second file and greater than first
                    mLaterDateList.Add(name, filedate)
                    mLaterDatePathList.Add(name, File)
                Else
                    If Not filedate = mEarlierDateList(name) Then
                        mLaterDateList.Add(name, mEarlierDateList(name))
                        mLaterDatePathList.Add(name, mEarlierDatePathList(name))
                        mEarlierDateList(name) = filedate
                        mEarlierDatePathList(name) = File
                    End If
                End If
            End If
        Else
            ' first file
            mEarlierDateList.Add(name, filedate)
            mEarlierDatePathList.Add(name, File)
        End If

    End Sub


    Public Shared Sub RecursiveTwoDateScan(path As String)

        For Each vbfil In System.IO.Directory.GetFiles(path, String.Format("{0}.{1}", mWildName, mScanExt))
            FileFoundTwoDate(vbfil)
        Next

        If System.IO.Directory.GetDirectories(path).Length = 0 Then Exit Sub
        For Each item In System.IO.Directory.GetDirectories(path)
            RecursiveTwoDateScan(item)
        Next

    End Sub



    Public Shared Sub RecursiveDateScan(path As String)

        For Each vbfil In System.IO.Directory.GetFiles(path, String.Format("{0}.{1}", mWildName, mScanExt))
            FileFoundCheckDate(vbfil)
        Next

        If System.IO.Directory.GetDirectories(path).Length = 0 Then Exit Sub
        For Each item In System.IO.Directory.GetDirectories(path)
            RecursiveDateScan(item)
        Next

    End Sub

    Private Shared mSortDates = Nothing
    Public Shared ReadOnly Property SortedDates As ArrayList
        Get
            If mSortDates Is Nothing Then UniqueSort()
            Return mSortDates
        End Get
    End Property

    Private Shared mCurFilename = Nothing
    Private Shared mKeys = Nothing
    Public Shared ReadOnly Property Keys As ArrayList
        Get
            If mKeys Is Nothing Then
                mKeys = New ArrayList
                For Each item In mUniqueFilename.Keys
                    mKeys.add(item)
                Next
                mKeys.Sort()
            End If
            Return mKeys
        End Get
    End Property


    Public Shared ReadOnly Property CurFilename As String
        Get
            If mCurFilename Is Nothing Then
                Dim keys = mUniqueFilename.Keys
                For Each item In keys
                    mCurFilename = item
                    Exit For
                Next
            End If
            Return mCurFilename
        End Get
    End Property

    Public Shared Sub UniqueSort()

        mSortDates = New ArrayList
        Dim keys = mUniqueFilename.Keys
        For Each item In keys
            mCurFilename = item
            Exit For
        Next
        For Each key In keys
            For Each item In mUniqueFilename(key)
                mSortDates.Add(item.key)
            Next
            Exit For
        Next
        mSortDates.Sort()

    End Sub

    Public Shared ReadOnly Property UniqueFilenameList As Hashtable
        Get
            Return mUniqueFilename(Keys(0))
        End Get
    End Property

    Private Shared mFileDateHash As Hashtable
    Private Shared mFileAnyHash As Hashtable
    Public Shared ReadOnly Property EarlierDateList As Hashtable
        Get
            Return mEarlierDateList
        End Get
    End Property

    Public Shared ReadOnly Property EarlierDatePathList As Hashtable
        Get
            Return mEarlierDatePathList
        End Get
    End Property

    Public Shared ReadOnly Property LaterDateList As Hashtable
        Get
            Return mLaterDateList
        End Get
    End Property

    Public Shared ReadOnly Property LaterDatePathList As Hashtable
        Get
            Return mLaterDatePathList
        End Get
    End Property

    Public Shared Sub ScanTwoDatesStart()

        mEarlierDateList = New Hashtable
        mLaterDateList = New Hashtable
        mEarlierDatePathList = New Hashtable
        mLaterDatePathList = New Hashtable
        AddHandler FindEvent, AddressOf FileFound
        RecursiveTwoDateScan(MyCodeRoot)
        RaiseEvent DoneEvent()

    End Sub


    Public Shared Sub ScanDateStart()

        mFileDateHash = New Hashtable
        mFileAnyHash = New Hashtable
        AddHandler FindEvent, AddressOf FileFound
        RecursiveDateScan(MyCodeRoot)

    End Sub

    Public Shared Sub ScanStart()

        mCompleteList = New ArrayList
        mUniqueDates = New ArrayList
        mUniqueFilename = New Hashtable
        AddHandler FindEvent, AddressOf FileFound
        RecursiveScan(MyCodeRoot)

        UniqueSort()

    End Sub

    Private Shared mFirstDateList As Hashtable
    Private Shared mNextDateList As Hashtable

    Public Shared ReadOnly Property FirstDateList As Hashtable
        Get
            Return mFirstDateList
        End Get
    End Property

    Public Shared ReadOnly Property NextDateList As Hashtable
        Get
            Return mNextDateList
        End Get
    End Property

    Private Shared mFirstDatePathList As Hashtable
    Private Shared mNextDatePathList As Hashtable

    Public Shared ReadOnly Property FirstDatePathList As Hashtable
        Get
            Return mFirstDatePathList
        End Get
    End Property

    Public Shared ReadOnly Property NextDatePathList As Hashtable
        Get
            Return mNextDatePathList
        End Get
    End Property

    Private Shared mDaysAgo As Integer = 30
    Private Shared mDateLimit As Date = Now.Date.AddDays(-mDaysAgo)

    Public Shared ReadOnly Property DateKeys As ArrayList
        Get
            Dim keys As New ArrayList
            If Not mFirstDateList.Count = 0 Then
                For Each key In mFirstDateList.Keys
                    keys.Add(key)
                Next
            End If
            Return keys
        End Get
    End Property

    Public Shared Sub DateScanStart()

        mFirstDateList = New Hashtable
        mNextDateList = New Hashtable
        mFirstDatePathList = New Hashtable
        mNextDatePathList = New Hashtable
        mPreDateList = New Hashtable
        mPrePathList = New Hashtable
        RecursiveDateScan(MyCodeRoot)



    End Sub

End Class
