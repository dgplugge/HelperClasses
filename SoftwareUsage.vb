' Name:     Software Usage Tracking
' Author:   Donald G Plugge
' Date:     6/4/2012
' Purpose:  Track the usage of various software packages by user and machine
Imports System.Xml.Linq
Imports System.Linq

Public Class SoftwareUsage

    ' dgp rev 11/29/2012
    Private Shared mUsersList As ArrayList
    Public Shared ReadOnly Property UsersList As ArrayList
        Get
            Return mUsersList
        End Get
    End Property

    ' dgp rev 11/29/2012
    Private Shared mDateList As ArrayList
    Public Shared ReadOnly Property DateList As ArrayList
        Get
            Return mDateList
        End Get
    End Property

    ' dgp rev 11/29/2012
    Private Shared mSoftwareList As ArrayList
    Public Shared ReadOnly Property SoftwareList As ArrayList
        Get
            Return mSoftwareList
        End Get
    End Property

    ' dgp rev 11/29/2012
    Private Shared mMachineList As ArrayList
    Public Shared ReadOnly Property MachineList As ArrayList
        Get
            Return mMachineList
        End Get
    End Property
    Private Shared mTimestampList As ArrayList

    ' dgp rev 11/29/2012
    Public Shared ReadOnly Property TimestampList As ArrayList
        Get
            Return mTimestampList
        End Get
    End Property



    Private Shared mUserElements
    Private Shared mCurUser

    ' dgp rev 11/29/2012
    Private Shared mFilteredList = Nothing
    Private Shared mFilteredCount = 0
    Public Shared ReadOnly Property FilteredCount As Integer
        Get
            Return mFilteredCount
        End Get
    End Property

    Public Shared ReadOnly Property FilteredList As List(Of XElement)
        Get
            If mFilteredList Is Nothing Then mFilteredList = (From item In SoftwareUsageXElement.Descendants("SoftwareUsage") Select item).ToList
            Return mFilteredList
        End Get
    End Property

    ' dgp rev 11/29/2012
    Public Shared Sub FilterByUser(ByVal user As String)

        mCurUser = user
        mSoftwareList = New ArrayList
        mDateList = New ArrayList
        mUsersList = New ArrayList
        mMachineList = New ArrayList


        Dim TempArray As New List(Of XElement)
        If Not SoftwareUsageXElement.IsEmpty Then
            For Each report In FilteredList
                If report.Elements("User").Value.ToLower = user.ToLower Then
                    If report.Elements("Software") IsNot Nothing And report.Elements("Date") IsNot Nothing Then
                        Try
                            mSoftwareList.Add(report.Elements("Software").Value)
                            mDateList.Add(report.Elements("Date").Value)
                            If Not mUsersList.Contains(user.ToLower) Then mUsersList.Add(report.Elements("User").Value.ToLower)
                            mMachineList.Add(report.Elements("Machine").Value)
                        Catch ex As Exception

                        End Try
                    End If
                    TempArray.Add(report)
                End If
            Next
            mFilteredCount = TempArray.Count
            mFilteredList = TempArray
        End If
        mUsersList.Sort()

    End Sub

    Private Shared mCurApp As String
    ' dgp rev 11/29/2012
    Public Shared Sub FilterBySoftware(ByVal app As String)

        mCurApp = app
        mSoftwareList = New ArrayList
        mDateList = New ArrayList
        mUsersList = New ArrayList
        mMachineList = New ArrayList

        Dim TempArray As New List(Of XElement)
        If Not SoftwareUsageXElement.IsEmpty Then
            For Each report In FilteredList
                If report.Elements("Software").Value.ToLower = app.ToLower Then
                    If report.Elements("Software") IsNot Nothing And report.Elements("Date") IsNot Nothing Then
                        Try
                            mSoftwareList.Add(report.Elements("Software").Value)
                            mDateList.Add(report.Elements("Date").Value)
                            mUsersList.Add(report.Elements("User").Value)
                            mMachineList.Add(report.Elements("Machine").Value)
                        Catch ex As Exception

                        End Try
                    End If
                    TempArray.Add(report)
                End If
            Next
            mFilteredCount = TempArray.Count
            mFilteredList = TempArray
        End If

    End Sub

    ' dgp rev 11/29/2012
    Public Shared Sub ResetFilter()

        mFilteredList = Nothing

    End Sub

    Private Shared mUniqueUser = Nothing
    Public Shared ReadOnly Property UniqueUser As ArrayList
        Get
            If mUniqueUser Is Nothing Then ScanUnique()
            Return mUniqueUser
        End Get
    End Property
    Private Shared mUniqueSoftware = Nothing
    Public Shared ReadOnly Property UniqueSoftware As ArrayList
        Get
            If mUniqueSoftware Is Nothing Then ScanUnique()
            Return mUniqueSoftware
        End Get
    End Property
    Private Shared mUniqueUserSoftware = Nothing
    Public Shared ReadOnly Property UniqueUserSoftware As ArrayList
        Get
            If mUniqueUserSoftware Is Nothing Then ScanUnique()
            Return mUniqueUserSoftware
        End Get
    End Property
    Private Shared mUniqueDate = Nothing
    Public Shared ReadOnly Property UniqueDate As ArrayList
        Get
            If mUniqueDate Is Nothing Then ScanUnique()
            Return mUniqueDate
        End Get
    End Property
    Private Shared mUniqueVersion = Nothing
    Public Shared ReadOnly Property UniqueVersion As ArrayList
        Get
            If mUniqueVersion Is Nothing Then ScanUnique()
            Return mUniqueVersion
        End Get
    End Property

    Private Shared mUniqueTime = Nothing
    Public Shared ReadOnly Property UniqueTime As ArrayList
        Get
            If mUniqueTime Is Nothing Then ScanUnique()
            Return mUniqueTime
        End Get
    End Property

    ' dgp rev 11/28/2012 
    Private Shared Sub ScanUnique()

        mUniqueUser = New ArrayList
        mUniqueUserSoftware = New ArrayList
        mUniqueSoftware = New ArrayList
        mUniqueDate = New ArrayList
        mUniqueVersion = New ArrayList
        mUniqueTime = New ArrayList
        mFilteredList = Nothing

        Dim combo = ""
        Dim flag As Boolean = False
        If Not SoftwareUsageXElement.IsEmpty Then
            For Each report In FilteredList
                flag = report.Elements("Software").Value IsNot Nothing
                flag = flag And report.Elements("Date").Value IsNot Nothing
                flag = flag And report.Elements("Software").Value IsNot Nothing
                flag = flag And report.Elements("User").Value IsNot Nothing
                flag = flag And report.Elements("Timestamp").Value IsNot Nothing
                flag = flag And report.Elements("Error").Value Is Nothing
                If flag Then
                    Try
                        combo = report.Elements("User").Value + report.Elements("Software").Value
                        If Not mUniqueUserSoftware.Contains(combo.ToLower) Then
                            mUniqueUserSoftware.Add(combo.ToLower)
                            mUniqueDate.Add(String.Format("{0} {1}", report.Elements("Date").Value, Format(Convert.ToDateTime(report.Elements("Timestamp").Attributes("Runtime").FirstOrDefault.Value), "HH:mm")))
                            If report.Elements("Timestamp").Attributes("Runtime") IsNot Nothing Then mUniqueTime.Add(Format(Convert.ToDateTime(report.Elements("Timestamp").Attributes("Runtime").FirstOrDefault.Value), "HH:mm"))
                            mUniqueSoftware.Add(report.Elements("Software").Value)
                            mUniqueUser.Add(report.Elements("User").Value)
                            mUniqueVersion.Add(report.Elements("Software").Value)
                        End If
                    Catch ex As Exception

                    End Try
                End If
            Next
        End If

    End Sub



    ' dgp rev 11/28/2012 
    Public Shared Sub ScanUsage()

        mUsersList = New ArrayList
        mTimestampList = New ArrayList
        mMachineList = New ArrayList

        mSoftwareList = New ArrayList
        mDateList = New ArrayList
        mFilteredList = Nothing

        Dim TempArray As New List(Of XElement)
        If Not SoftwareUsageXElement.IsEmpty Then
            For Each report In FilteredList
                If report.Elements("Software") IsNot Nothing And report.Elements("Date") IsNot Nothing Then
                    Try
                        mSoftwareList.Add(report.Elements("Software").Value)
                        mDateList.Add(report.Elements("Date").Value)
                        If Not mUsersList.Contains(report.Elements("User").Value.ToLower) Then mUsersList.Add(report.Elements("User").Value.ToLower)
                        mMachineList.Add(report.Elements("Machine").Value)
                    Catch ex As Exception

                    End Try
                End If
                TempArray.Add(report)
            Next
            mFilteredCount = TempArray.Count
            mFilteredList = TempArray
        End If
        mUsersList.Sort()

    End Sub

    ' dgp rev 11/10/2011
    Private Shared mSoftwareUsageLogName = Nothing
    Public Shared ReadOnly Property SoftwareUsageLogName() As String
        Get
            If (mSoftwareUsageLogName Is Nothing) Then mSoftwareUsageLogName = String.Format("{0}.xml", "SoftwareUsage")
            Return mSoftwareUsageLogName
        End Get
    End Property

    ' dgp rev 11/10/2011
    Public Shared ReadOnly Property SoftwareUsageLogFullSpec As String
        Get
            Return System.IO.Path.Combine(InstallInfo.GlobalLogServerPath, SoftwareUsageLogName)
        End Get
    End Property

    ' dgp rev 11/29/2012
    Private Shared mSoftwareUsageXElement As XElement
    Public Shared Property SoftwareUsageXElement As XElement
        Get
            Try
                If System.IO.File.Exists(SoftwareUsageLogFullSpec) Then
                    mSoftwareUsageXElement = XElement.Load(SoftwareUsageLogFullSpec)
                Else
                    mSoftwareUsageXElement = New XElement("Packages", NewUsageLogElement)
                End If
            Catch ex As Exception

            End Try

            Return mSoftwareUsageXElement
        End Get
        Set(ByVal value As XElement)
            mSoftwareUsageXElement = value
        End Set
    End Property

    ' dgp rev 11/29/2012
    Private Shared mMessage As String = ""
    Private Shared Function NewUsageLogElement() As XElement

        Try
            NewUsageLogElement =
                New XElement("SoftwareUsage",
                                  New XElement("Timestamp", New XAttribute("Runtime", DateTime.Now)),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("Software", VerInfo.AssemblyFullInfo),
                                  New XElement("Machine", Environment.MachineName),
                                  New XElement("User", Environment.UserName))

        Catch ex As Exception
            mMessage = ex.Message
            NewUsageLogElement =
                New XElement("SoftwareUsage",
                                  New XElement("Timestamp", New XAttribute("modified", DateTime.Now)),
                                  New XElement("Error", ex.Message),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("Machine", Environment.MachineName),
                                  New XElement("User", Environment.UserName))
        End Try

    End Function

    ' dgp rev 5/31/2011
    Public Shared Function UsageLogging() As Boolean

        Dim LocalValid = False
        UsageLogging = False
        Try
            Dim mMainTree As XElement

            mMainTree = SoftwareUsageXElement
            If System.IO.File.Exists(SoftwareUsageLogFullSpec) Then
                Dim lastEntry As XElement = mMainTree.Element("SoftwareUsage")
                lastEntry.AddBeforeSelf(NewUsageLogElement)
            End If
            If Utility.Create_Tree(InstallInfo.GlobalLogServerPath) Then
                mMainTree.Save(SoftwareUsageLogFullSpec)
            End If
        Catch ex As Exception

        End Try

    End Function

End Class
