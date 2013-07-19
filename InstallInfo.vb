' Name:     InstallInfo
' Author:   Donald G Plugge
' Date:     3/8/2012
' Purpose:  PC Installation Info
Imports System.Threading
Imports Microsoft.Win32
Imports System.IO
Imports System.Xml.Linq

Public Class InstallInfo

    ' dgp rev 2/16/2012
    Private Shared mRegPath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

    ' dgp rev 2/17/2012 Text Report
    Public Shared Function TextReport() As String

        Return ProductReport()

    End Function

    Private Shared mDotNetList As ArrayList = Nothing
    Private Shared mFrameworkPath = System.IO.Path.Combine(Environment.GetEnvironmentVariable("WinDir"), "Microsoft.Net\Framework")
    ' dgp rev 3/6/2012
    Public Shared ReadOnly Property DotNetVersions As ArrayList
        Get
            If mDotNetList Is Nothing Then
                mDotNetList = New ArrayList
                If System.IO.Directory.Exists(mFrameworkPath) Then
                    If System.IO.Directory.GetDirectories(mFrameworkPath).Length > 0 Then
                        Dim item
                        For Each item In System.IO.Directory.GetDirectories(mFrameworkPath)
                            If System.IO.Path.GetFileName(item).StartsWith("v") Then mDotNetList.Add(item)
                        Next
                    End If
                End If
            End If
            Return mDotNetList
        End Get
    End Property

    Private Shared mMemory = Nothing
    ' dgp rev 2/17/2012 
    Public Shared ReadOnly Property MemoryReport As String
        Get
            If mMemory Is Nothing Then
                mMemory = String.Format("{0}Physical Memory: {0}", vbCrLf)
                mMemory += Format((My.Computer.Info.TotalPhysicalMemory * 1024.0 ^ -3.0), "##0.### GB")
                mMemory += vbCrLf
                mMemory += vbCrLf
            End If
            Return mMemory
        End Get
    End Property


    ' dgp rev 2/17/2012 
    Public Shared ReadOnly Property DotNetReport As String
        Get
            DotNetReport = String.Format("{0}.NET Frameworks: {0}", vbCrLf)
            Dim item
            For Each item In DotNetVersions
                DotNetReport += item + vbCrLf
            Next
            DotNetReport += vbCrLf
            Return DotNetReport
        End Get
    End Property

    ' dgp rev 2/17/2012 Text Report
    Public Shared ReadOnly Property Products As ArrayList
        Get
            Products = New ArrayList
            Dim item As Hashtable
            For Each item In KeyMatchList
                If item.ContainsKey("DisplayName") Then Products.Add(item("DisplayName"))
            Next
        End Get
    End Property

    Private Shared mLogger As Logger

    ' dgp rev 3/15/2012 
    Public Shared Sub Analytics(ByVal objLOG As Logger)

        mLogger = objLOG
        Try
            OneTimeScan()
        Catch ex As Exception

        End Try
        Try
            SoftwareUsage.UsageLogging()
        Catch ex As Exception

        End Try

    End Sub

    ' dgp rev 3/15/2012 
    Public Shared Sub Analytics()

        Try
            OneTimeScan()
        Catch ex As Exception

        End Try
        Try
            SoftwareUsage.UsageLogging()
        Catch ex As Exception

        End Try

    End Sub

    Private Shared mServer As String = "NT-EIB-10-6B16"
    Public Shared ReadOnly Property Server As String
        Get
            Return mServer
        End Get
    End Property

    Private Shared mUserShare As String = "Root2"
    Public Shared ReadOnly Property UserShare As String
        Get
            Return mUserShare
        End Get
    End Property

    Private Shared mSoftwareShare As String = "Software"
    Public Shared ReadOnly Property GlobalSoftwareShare As String
        Get
            Return mSoftwareShare
        End Get
    End Property

    ' dgp rev 7/15/09 Path to Server
    Private Shared mUserLogServerPath = Nothing
    Public Shared ReadOnly Property UserLogServerPath() As String
        Get
            mUserLogServerPath = String.Format("\\{0}\{1}\Users\{2}", Server, UserShare, Environment.UserName)
            mUserLogServerPath = System.IO.Path.Combine(mUserLogServerPath, "Logging")
            Return mUserLogServerPath
        End Get
    End Property

    ' dgp rev 7/15/09 Path to Server
    Private Shared mUserHelpServerPath = Nothing
    Public Shared ReadOnly Property UserHelpServerPath() As String
        Get
            mUserHelpServerPath = String.Format("\\{0}\{1}\Users\{2}", Server, UserShare, Environment.UserName)
            mUserHelpServerPath = System.IO.Path.Combine(mUserHelpServerPath, "Help")
            Return mUserHelpServerPath
        End Get
    End Property
    ' dgp rev 7/15/09 Path to Server
    Private Shared mGlobalLogServerPath = Nothing
    Public Shared ReadOnly Property GlobalLogServerPath() As String
        Get
            mGlobalLogServerPath = String.Format("\\{0}\{1}", Server, GlobalSoftwareShare)
            mGlobalLogServerPath = System.IO.Path.Combine(mGlobalLogServerPath, "Logging")
            Return mGlobalLogServerPath
        End Get
    End Property

    ' dgp rev 7/15/09 Path to Server
    Private Shared mGlobalHelpServerPath = Nothing
    Public Shared ReadOnly Property GlobalHelpServerPath() As String
        Get
            mGlobalHelpServerPath = String.Format("\\{0}\{1}", Server, GlobalSoftwareShare)
            mGlobalHelpServerPath = System.IO.Path.Combine(mGlobalHelpServerPath, "Help")
            Return mGlobalHelpServerPath
        End Get
    End Property

    Private Shared mLogThread As Thread
    Private Shared Sub LaunchLogThread()

        mLogThread = New Threading.Thread(New Threading.ThreadStart(AddressOf Analytics))
        mLogThread.Name = "Software Logging"

        ' dgp rev 10/19/09 Reinsert the event handler
        '        AddHandler objPVW.ExitedEvent, AddressOf ProcessDone

        mLogThread.Start()

    End Sub

    Private Shared mLogDoc As XDocument

    Private Shared mNewBranch As XElement

    ' dgp rev 3/6/2012
    Public Shared Sub OneTimeScan()

        If Not ScanCheck Then
            Dim objReport = New HelperClasses.EmailReporting()
            Dim Report As String = VerInfo.AssemblyFullInfo + vbCrLf
            Report += HelperClasses.InstallInfo.MemoryReport()
            Report += HelperClasses.InstallInfo.DotNetReport()
            Report += HelperClasses.InstallInfo.AssembliesReport()
            Report += HelperClasses.InstallInfo.ProductReport()
            objReport.EmailText = Report
            objReport.SendReport("ProductScan")
            Dim sw As New StreamWriter(LogSpec)
            sw.WriteLine(Report)
            sw.Close()
            SoftwareUpgrade.AppendUpgradeLog()
        End If

    End Sub

    Private Shared mScanCheck = Nothing
    ' dgp rev 3/6/2012
    Private Shared ReadOnly Property ScanCheck As Boolean
        Get
            If mScanCheck Is Nothing Then
                mScanCheck = System.IO.File.Exists(LogSpec)
            End If
            Return mScanCheck
        End Get
    End Property

    ' dgp rev 3/6/2012
    Private Shared Sub RegistryScan()

        mUniqueNames = New ArrayList
        mKeyMatchList = New ArrayList
        Dim RegKey As RegistryKey
        Try
            RegKey = Registry.LocalMachine.OpenSubKey(mRegPath, False)
            ScanRegistryProducts(RegKey)
        Catch ex As Exception

        End Try
        Try
            RegKey = Registry.CurrentUser.OpenSubKey(mRegPath, False)
            ScanRegistryProducts(RegKey)
        Catch ex As Exception

        End Try

    End Sub

    ' dgp rev 5/29/2012
    Private Shared Sub ScanRegistryAssemblies(ByVal RegKey)

        If RegKey Is Nothing Then Return
        Dim PPList As New ArrayList
        Dim item
        For Each item In RegKey.GetSubKeyNames
            If item.ToString.ToLower.Contains("microsoft.office.interop.powerpoint") Then
                If item.ToString.ToLower.Contains("pia") Then
                    PPList.Add(item.ToString)
                End If
            End If
        Next

    End Sub
    Private Shared mRegAssembliesPath As String = "SOFTWARE\Classes\Installer\Assemblies"

    Private Shared mAssemblyPath = "C:\Windows\assembly"
    Private Shared mPIAFilesList As ArrayList
    Private Shared mPIAFilesPath As ArrayList
    Private Shared mPIAFilesVers As ArrayList

    Private Shared mAssembliesReport As String

    Private Shared Sub ScanPIAs()

        Dim GACItem
        Dim PPTItem
        Dim vers

        Dim CurGAC
        Dim CurName
        Dim CurVer

        mAssembliesReport = String.Format("{0}Assemblies:{0}", vbCrLf)
        mPIAFilesList = New ArrayList
        mPIAFilesPath = New ArrayList
        mPIAFilesVers = New ArrayList
        If System.IO.Directory.Exists(mAssemblyPath) Then
            For Each GACItem In System.IO.Directory.GetDirectories(mAssemblyPath)
                CurGAC = System.IO.Path.GetFileNameWithoutExtension(GACItem)
                For Each PPTItem In System.IO.Directory.GetDirectories(GACItem)
                    If PPTItem.ToLower.Contains("microsoft.office.interop.powerpoint") Then
                        CurName = System.IO.Path.GetFileName(PPTItem)
                        mPIAFilesList.Add(CurName)
                        mPIAFilesPath.Add(System.IO.Path.GetDirectoryName(PPTItem))
                        For Each vers In System.IO.Directory.GetDirectories(PPTItem)
                            CurVer = System.IO.Path.GetFileNameWithoutExtension(vers)
                            mPIAFilesVers.Add(CurVer)
                            mAssembliesReport += String.Format("{0} {1} {2} {3}", CurGAC, CurName, CurVer, vbCrLf)
                        Next
                    End If
                Next
            Next
        End If

    End Sub


    ' dgp rev 6/4/2012
    Private Shared Sub RegisteredAssemblies()

        Dim RegKey As RegistryKey
        Try
            RegKey = Registry.LocalMachine.OpenSubKey(mRegAssembliesPath, False)
            ScanRegistryAssemblies(RegKey)
            ScanPIAs()
        Catch
        End Try

    End Sub

    ' dgp rev 2/17/2012 
    Public Shared ReadOnly Property AssembliesReport As String
        Get
            Try
                RegisteredAssemblies()
            Catch ex As Exception
                mAssembliesReport += String.Format("Error: {0}", ex.Message)
            End Try

            Return mAssembliesReport
        End Get
    End Property


    ' dgp rev 2/17/2012 Text Report
    Public Shared Function ProductReport() As String

        If KeyMatchList Is Nothing Then
            Return "No Registry Matches Found"
        End If
        Dim item As Hashtable
        ProductReport = ""
        For Each item In KeyMatchList
            ProductReport += "Product: " + vbCrLf
            Dim Keys = item.Keys
            Dim key
            For Each key In Keys
                ProductReport += String.Format("{0}: {1}{2}", key, item(key), vbCrLf)
            Next
            ProductReport += vbCrLf
        Next

    End Function

    ' dgp rev 2/17/2012 Text Report
    Private Shared mKeyMatchList As ArrayList = Nothing
    Public Shared ReadOnly Property KeyMatchList As ArrayList
        Get
            If mKeyMatchList Is Nothing Then RegistryScan()
            Return mKeyMatchList
        End Get
    End Property

    Private Shared mUniqueNames As ArrayList

    ' dgp rev 2/16/2012
    Private Shared Sub AddItem(ByVal subregkey As RegistryKey)

        Dim item = subregkey.Name
        If mUniqueNames.Contains(item) Then Return
        mUniqueNames.Add(item)

        Dim keys As New ArrayList
        keys.Add("Contact")
        keys.Add("DisplayName")
        keys.Add("InstallSource")
        keys.Add("DisplayVersion")
        keys.Add("InstallDate")
        keys.Add("Publisher")
        keys.Add("UninstallString")

        Dim keyword
        Dim info As New Hashtable
        For Each keyword In keys
            If subregkey.GetValue(keyword) IsNot Nothing Then info.Add(keyword, subregkey.GetValue(keyword))
        Next
        mKeyMatchList.Add(info)

    End Sub

    ' dgp rev 2/15/2012
    Private Shared Sub ScanRegistryProducts(ByVal RegKey)

        Dim SubRegKey As RegistryKey
        Dim name As String
        Dim auth As String
        Dim source As String
        Dim pub As String
        If RegKey Is Nothing Then Return
        Dim item
        Dim uninstal
        For Each item In RegKey.GetSubKeyNames
            SubRegKey = RegKey.OpenSubKey(item, False)
            name = SubRegKey.GetValue("DisplayName")
            auth = SubRegKey.GetValue("Contact")
            source = SubRegKey.GetValue("InstallSource")
            pub = SubRegKey.GetValue("Publisher")
            uninstal = SubRegKey.GetValue("UninstallString")
            If uninstal IsNot Nothing Then If uninstal.ToLower.Contains("dfshim") Then AddItem(SubRegKey)
            If auth IsNot Nothing Then If auth.ToLower.Contains("plugge") Then AddItem(SubRegKey)
            If source IsNot Nothing Then If source.ToLower.Contains("nt-eib-10-6b16") Then AddItem(SubRegKey)
            If pub IsNot Nothing Then If pub.ToLower.Contains("flow control") Then AddItem(SubRegKey)
            If name IsNot Nothing Then If name.ToLower.Contains("pv-wave") Then AddItem(SubRegKey)
        Next

    End Sub

    Private Shared PVWCmd As Process
    Private Shared mProcThread As Thread

    ' dgp rev 3/6/2012
    Private Shared mLogSpec = Nothing
    Public Shared ReadOnly Property LogSpec As String
        Get
            If mLogSpec Is Nothing Then
                mLogSpec = System.IO.Path.Combine(LogPath, String.Format("{0}.log", VerInfo.AssemblyFullInfo))
            End If
            Return mLogSpec
        End Get
    End Property

    ' dgp rev 3/6/2012
    Private Shared mLogPath = Nothing
    Public Shared ReadOnly Property LogPath As String
        Get
            If mLogPath Is Nothing Then
                mLogPath = System.Environment.GetEnvironmentVariable("APPDATA")
                mLogPath = System.IO.Path.Combine(mLogPath, "Flow Control")
            End If
            Return mLogPath
        End Get
    End Property

End Class
