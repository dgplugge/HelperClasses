' Name:     Software Usage Tracking
' Author:   Donald G Plugge
' Date:     6/4/2012
' Purpose:  Track the usage of various software packages by user and machine
Imports System.Xml.Linq
Imports System.Linq

Public Class SoftwareUpgrade

    Private Shared mMessage As String = ""

    Private Shared mServerPath As String = "\\nt-eib-10-6b16\Distribution\Versions"
    Public Shared ReadOnly Property ServerPath As String
        Get
            Return mServerPath
        End Get
    End Property

    Private Shared mXMLFilename = Nothing
    Public Shared Property XMLFilename As String
        Get
            If mXMLFilename Is Nothing Then
                If Environment.UserName.ToLower = "plugged" Then
                    mXMLFilename = "TestConfig.xml"
                Else
                    mXMLFilename = "SoftwareConfig.xml"
                End If
            End If
            Return mXMLFilename
        End Get
        Set(ByVal value As String)
            value = String.Format("{0}.xml", value.ToLower.Replace(".xml", ""))
            mXMLFilename = value
        End Set
    End Property

    ' dgp rev 8/29/2012 
    Private Shared mXMLStore = Nothing
    Public Shared ReadOnly Property UpgradeStore As HelperClasses.XMLSettings
        Get
            If mXMLStore Is Nothing Then
                mXMLStore = New HelperClasses.XMLSettings(System.IO.Path.Combine(ServerPath, XMLFilename))
            End If
            Return mXMLStore
        End Get
    End Property

    ' dgp rev 11/10/2011
    Private Shared mSoftwareUpgradeLogName = Nothing
    Public Shared ReadOnly Property SoftwareUpgradeLogName() As String
        Get
            If (mSoftwareUpgradeLogName Is Nothing) Then mSoftwareUpgradeLogName = String.Format("{0}.xml", "SoftwareUpgrade")
            Return mSoftwareUpgradeLogName
        End Get
    End Property

    ' dgp rev 11/10/2011
    Public Shared ReadOnly Property SoftwareUpgradeLogFullSpec As String
        Get
            Return System.IO.Path.Combine(InstallInfo.GlobalLogServerPath, SoftwareUpgradeLogName)
        End Get
    End Property

    Private Shared mLogger As HelperClasses.Logger
    Private Shared Property Logger As HelperClasses.Logger
        Get
            If mLogger Is Nothing Then mLogger = New HelperClasses.Logger("SoftwareUpgrade")
            Return mLogger
        End Get
        Set(ByVal value As HelperClasses.Logger)
            mLogger = value
        End Set
    End Property

    ' dgp rev 10/19/2010
    Private Shared Function NewUpgradeLogElement() As XElement

        Try
            NewUpgradeLogElement =
                New XElement("SoftwareUpgrade",
                                  New XElement("Timestamp", New XAttribute("Runtime", DateTime.Now)),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("Software", VerInfo.AssemblyFullInfo),
                                  New XElement("Machine", Environment.MachineName),
                                  New XElement("User", Environment.UserName))

        Catch ex As Exception
            mMessage = ex.Message
            NewUpgradeLogElement =
                New XElement("SoftwareUpgrade",
                                  New XElement("Timestamp", New XAttribute("modified", DateTime.Now)),
                                  New XElement("Error", ex.Message),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("Machine", Environment.MachineName),
                                  New XElement("User", Environment.UserName))
        End Try

    End Function

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

    Private Shared mFilteredCount = 0
    Public Shared ReadOnly Property FilteredCount As Integer
        Get
            Return mFilteredCount
        End Get
    End Property

    ' dgp rev 11/29/2012
    Public Shared Sub ResetFilter()

        mFilteredList = Nothing

    End Sub

    ' dgp rev 11/29/2012
    Public Shared Sub FilterBySoftware(ByVal app As String)

        mSoftwareList = New ArrayList
        mDateList = New ArrayList
        mUsersList = New ArrayList
        mMachineList = New ArrayList

        Dim TempArray As New List(Of XElement)
        If Not SoftwareUpgradeXElement.IsEmpty Then
            For Each report In FilteredList
                If report.Elements("Software").Value.ToLower = app.ToLower Then
                    If report.Elements("Software") IsNot Nothing And report.Elements("Date") IsNot Nothing Then
                        Try
                            mSoftwareList.Add(report.Elements("Software").Value)
                            mDateList.Add(report.Elements("Date").Value)
                            mUsersList.Add(report.Elements("User").Value.ToLower)
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

    ' dgp rev 11/28/2012 
    Private Shared Sub ScanUnique()

        mUniqueUser = New ArrayList
        mUniqueUserSoftware = New ArrayList
        mUniqueSoftware = New ArrayList
        mUniqueDate = New ArrayList
        mUniqueVersion = New ArrayList
        mFilteredList = Nothing

        Dim combo = ""
        Dim flag
        If Not SoftwareUpgradeXElement.IsEmpty Then
            For Each report In FilteredList
                flag = report.Elements("Software").Value IsNot Nothing
                flag = flag And report.Elements("Date").Value IsNot Nothing
                flag = flag And report.Elements("Software").Value IsNot Nothing
                flag = flag And report.Elements("User").Value IsNot Nothing
                flag = flag And report.Elements("Timestamp").Value IsNot Nothing
                flag = flag And report.Elements("Error").Value Is Nothing
                If flag IsNot Nothing Then
                    Try
                        combo = report.Elements("User").Value + report.Elements("Software").Value
                        If Not mUniqueUserSoftware.Contains(combo.ToLower) Then
                            mUniqueUserSoftware.Add(combo.ToLower)
                            mUniqueDate.Add(String.Format("{0} {1}", report.Elements("Date").Value, Format(Convert.ToDateTime(report.Elements("Timestamp").Attributes("Runtime").FirstOrDefault.Value), "HH:mm")))
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



    ' dgp rev 11/29/2012
    Public Shared Sub FilterByUser(ByVal user As String)

        mSoftwareList = New ArrayList
        mDateList = New ArrayList
        mUsersList = New ArrayList
        mMachineList = New ArrayList

        Dim TempArray As New List(Of XElement)
        If Not SoftwareUpgradeXElement.IsEmpty Then
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

    ' dgp rev 11/28/2012 
    Public Shared Sub ScanUpgrade()

        mUsersList = New ArrayList
        mTimestampList = New ArrayList
        mMachineList = New ArrayList

        mSoftwareList = New ArrayList
        mDateList = New ArrayList
        mFilteredList = Nothing

        Dim TempArray As New List(Of XElement)
        If Not SoftwareUpgradeXElement.IsEmpty Then
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

    Private Shared mFilteredList = Nothing

    Public Shared ReadOnly Property FilteredList As List(Of XElement)
        Get
            If mFilteredList Is Nothing Then mFilteredList = (From item In SoftwareUpgradeXElement.Descendants("SoftwareUpgrade") Select item).ToList
            Return mFilteredList
        End Get
    End Property

    Private Shared mSoftwareUpgradeXElement As XElement
    Public Shared ReadOnly Property SoftwareUpgradeXElement As XElement
        Get
            Try
                If System.IO.File.Exists(SoftwareUpgradeLogFullSpec) Then
                    mSoftwareUpgradeXElement = XElement.Load(SoftwareUpgradeLogFullSpec)
                Else
                    mSoftwareUpgradeXElement = New XElement("Packages", NewUpgradeLogElement)
                End If
            Catch ex As Exception

            End Try

            Return mSoftwareUpgradeXElement
        End Get
    End Property

    ' dgp rev 5/31/2011
    Public Shared Function AppendUpgradeLog() As Boolean

        Try

            Dim mMainTree As XElement

            Logger.Log_Info(String.Format("Software Upgrade {0}", SoftwareUpgradeLogFullSpec))
            mMainTree = SoftwareUpgradeXElement
            If System.IO.File.Exists(SoftwareUpgradeLogFullSpec) Then
                Logger.Log_Info(String.Format("Appending"))
                Dim lastEntry As XElement = mMainTree.Element("SoftwareUpgrade")
                lastEntry.AddBeforeSelf(NewUpgradeLogElement)
            Else
                Logger.Log_Info(String.Format("Creating New Log"))
            End If
            Logger.Log_Info(String.Format("Check global path {0}", InstallInfo.GlobalLogServerPath))
            If Utility.Create_Tree(InstallInfo.GlobalLogServerPath) Then
                mMainTree.Save(SoftwareUpgradeLogFullSpec)
                Logger.Log_Info(String.Format("Successful Append {0}", SoftwareUpgradeLogFullSpec))
            Else
                Logger.Log_Info(String.Format("{0} Not Found", InstallInfo.GlobalLogServerPath))
            End If
        Catch ex As Exception
            Logger.Log_Info(String.Format("Error {0}", ex.Message))
            Return False
        End Try
        Return True

    End Function

    Public Sub New(ByVal filespec As String)

        mXMLStore = New XMLSettings(filespec)

    End Sub

    Private mUserInfo As UserInfo
    Public Property MyUserInfo As UserInfo
        Get
            Return mUserInfo
        End Get
        Set(ByVal value As UserInfo)
            mUserInfo = value
        End Set
    End Property

    Private Shared mPVWCmd As Process

    Private Shared mTogShell As Boolean = True
    Private Shared mTogWindow As Boolean = True
    Public Delegate Sub WorkEventHandler(ByVal SomeString As String)
    Public Shared Event WorkEvent As WorkEventHandler

    Private Shared Sub ParseUninstall()

        Dim info = SoftwareHandler.ActiveUninstallString.ToString().Split(" ")

        mActiveUninstallFile = info(0)
        mActiveUninstallAttr = SoftwareHandler.ActiveUninstallString.ToString.Substring(mActiveUninstallFile.Length)

    End Sub

    Private Shared mActiveUninstallFile = Nothing
    Public Shared ReadOnly Property ActiveUninstallFile As String
        Get
            If mActiveUninstallFile Is Nothing Then ParseUninstall()
            Return mActiveUninstallFile
        End Get
    End Property
    Private Shared mActiveUninstallAttr = Nothing
    Public Shared ReadOnly Property ActiveUninstallAttr As String
        Get
            If mActiveUninstallAttr Is Nothing Then ParseUninstall()
            Return mActiveUninstallAttr
        End Get
    End Property

    Private Shared mCurUserInfo As UserInfo
    ' dgp rev 12/4/2012
    Public Shared Function DoesApplicationNeedUpgrade() As Boolean

        DoesApplicationNeedUpgrade = False
        Try

            HelperClasses.SoftwareUpgrade.XMLFilename = "SoftwareUpgrades"
            If Environment.OSVersion.ToString.ToLower.Contains("nt 6") Then
                SoftwareUpgrade.UpgradeStore.CurElementName = "Upgrade7"
            Else
                SoftwareUpgrade.UpgradeStore.CurElementName = "Upgrade"
            End If
            mCurUserInfo = New UserInfo(Environment.UserName)
            If mCurUserInfo.UpgradeRequired() Then

                Try

                    If SoftwareHandler.ActiveUninstallString IsNot Nothing Then
                        Dim insUninstall As New frmUninstallUpgrade
                        insUninstall.FormUnInstallString = SoftwareHandler.ActiveUninstallString.ToString
                        insUninstall.FormInstallString = mCurUserInfo.UpgradeAppPath
                        insUninstall.ShowDialog()
                        DoesApplicationNeedUpgrade = SoftwareUpgrade.InstallSuccessfull And SoftwareUpgrade.UnInstallSuccessfull
                    Else
                        MsgBox(String.Format("No Active UninstallString", MsgBoxStyle.Information))
                    End If
                Catch ex As Exception
                    MsgBox(String.Format("{0} -- {1} in {2}", ex.Message, mCurUserInfo.InstallAppName, System.IO.Path.GetDirectoryName(HelperClasses.VerInfo.CurrentUpdateLocation)), MsgBoxStyle.Information)
                End Try

            End If
        Catch ex As Exception

        End Try

    End Function

    Public Shared Sub InstallApp(App As String)

        mInstallSuccessfull = False
        Try

            mPVWCmd = New Process


            ' a new process is created from running the PV-Wave program
            mPVWCmd.StartInfo.UseShellExecute = mTogShell
            ' redirect IO
            ' PVWCmd.StartInfo.RedirectStandardOutput = True
            '       PVWCmd.StartInfo.RedirectStandardError = True
            '      PVWCmd.StartInfo.RedirectStandardInput = True
            ' don't even bring up the console window
            mPVWCmd.StartInfo.CreateNoWindow = mTogWindow
            ' executable command line info
            mPVWCmd.StartInfo.FileName = App

            '        PVWCmd.StartInfo.RedirectStandardInput = True

            '        PVWCmd.StartInfo.WorkingDirectory = Dist_Str
            mPVWCmd.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(HelperClasses.VerInfo.CurrentUpdateLocation)

            '            PVWCmd.EnableRaisingEvents = True
            '   
            ' Add an event handler.
            '
            '           AddHandler PVWCmd.Exited, AddressOf Me.ProcessExited

            mPVWCmd.Start()
            mPVWCmd.WaitForExit()
            mInstallSuccessfull = True
        Catch ex As Exception

        End Try



    End Sub

    Public Shared Sub InstallApp()

        mInstallSuccessfull = False
        Try

            mPVWCmd = New Process


            ' a new process is created from running the PV-Wave program
            mPVWCmd.StartInfo.UseShellExecute = mTogShell
            ' redirect IO
            ' PVWCmd.StartInfo.RedirectStandardOutput = True
            '       PVWCmd.StartInfo.RedirectStandardError = True
            '      PVWCmd.StartInfo.RedirectStandardInput = True
            ' don't even bring up the console window
            mPVWCmd.StartInfo.CreateNoWindow = mTogWindow
            ' executable command line info
            mPVWCmd.StartInfo.FileName = mCurUserInfo.UpgradeAppPath

            '        PVWCmd.StartInfo.RedirectStandardInput = True

            '        PVWCmd.StartInfo.WorkingDirectory = Dist_Str
            mPVWCmd.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(HelperClasses.VerInfo.CurrentUpdateLocation)

            '            PVWCmd.EnableRaisingEvents = True
            '   
            ' Add an event handler.
            '
            '           AddHandler PVWCmd.Exited, AddressOf Me.ProcessExited

            mPVWCmd.Start()
            mPVWCmd.WaitForExit()
            mInstallSuccessfull = True
        Catch ex As Exception

        End Try



    End Sub

    Private Shared mUnInstallSuccessfull As Boolean = False
    Public Shared ReadOnly Property UnInstallSuccessfull As Boolean
        Get
            Return mUnInstallSuccessfull
        End Get
    End Property

    Private Shared mInstallSuccessfull As Boolean = False
    Public Shared ReadOnly Property InstallSuccessfull As Boolean
        Get
            Return mInstallSuccessfull
        End Get
    End Property


    Public Shared Sub UnInstallApp()

        mUnInstallSuccessfull = False
        Try

            mPVWCmd = New Process


            ' a new process is created from running the PV-Wave program
            mPVWCmd.StartInfo.UseShellExecute = mTogShell
            ' redirect IO
            ' PVWCmd.StartInfo.RedirectStandardOutput = True
            '       PVWCmd.StartInfo.RedirectStandardError = True
            '      PVWCmd.StartInfo.RedirectStandardInput = True
            ' don't even bring up the console window
            mPVWCmd.StartInfo.CreateNoWindow = mTogWindow
            ' executable command line info
            mPVWCmd.StartInfo.FileName = ActiveUninstallFile

            '        PVWCmd.StartInfo.RedirectStandardInput = True

            '        PVWCmd.StartInfo.WorkingDirectory = Dist_Str
            mPVWCmd.StartInfo.Arguments = ActiveUninstallAttr

            '            PVWCmd.EnableRaisingEvents = True
            '   
            ' Add an event handler.
            '
            '           AddHandler PVWCmd.Exited, AddressOf Me.ProcessExited

            mPVWCmd.Start()
            mPVWCmd.WaitForExit()
            mUnInstallSuccessfull = True

        Catch ex As Exception

        End Try
        RaiseEvent WorkEvent("Done")

    End Sub




End Class
