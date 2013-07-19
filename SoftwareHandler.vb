Imports Microsoft.Win32
Imports System.Threading
Imports HelperClasses

Public Class SoftwareHandler

    Private Shared mTrackerRemove As Boolean = False
    Private Shared mTrackerInstall As Boolean = False
    Private Shared mTrackerRemovalString As String = ""
    Private Shared mTrackerInstallString As String = "\\nt-eib-10-6b16\www\BetaDistribution\RunTracker\RunTracker.application"
    Private Shared mActiveInstallString As String = "\\nt-eib-10-6b16\www\TestDistribution\TestDeployBeta\DeployTestActive.application"

    Private Shared mUniqueNames As ArrayList

    ' dgp rev 2/16/2012
    Private Shared mRegPath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
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

    ' dgp rev 3/6/2012
    Public Shared Sub RegistryScan()

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

    Public Shared ReadOnly Property AssemblyName As String
        Get
            Return System.IO.Path.GetFileNameWithoutExtension(HelperClasses.VerInfo.AssemblyName)
        End Get
    End Property

    Private Shared ReadOnly Property OldVersion As String
        Get
            Return AssemblyName.Replace("7", "")
        End Get
    End Property

    Private Shared Function NewVersion() As Boolean

        Return (AssemblyName = AssemblyName.Replace("7", ""))

    End Function

    Private Shared Function ValidState() As Boolean

        ValidState = AssemblyName.Contains("7")

    End Function

    Private Shared mActiveUninstallString = Nothing
    Public Shared ReadOnly Property ActiveUninstallString As String
        Get
            If mActiveUninstallString Is Nothing Then ScanUninstall()
            Return mActiveUninstallString
        End Get
    End Property

    Private Shared Sub ScanUninstall()

        '        HelperClasses.InstallInfo.Products
        Dim ver As Hashtable
        mUninstallList = New ArrayList
        mUninstallString = ""
        SoftwareHandler.RegistryScan()
        For Each ver In SoftwareHandler.KeyMatchList
            mUninstallList.add(ver.Item("UninstallString"))
            If ver.ContainsKey("UninstallString") Then
                If ver.Item("UninstallString").tolower.contains(String.Format("{0}.application", AssemblyName.ToLower)) Then
                    mUninstallString = ver.Item("UninstallString")
                    mActiveUninstallString = mUninstallString
                End If
            End If
        Next
    End Sub

    Private Shared mUninstallList = Nothing
    Public Shared ReadOnly Property UninstallList As ArrayList
        Get
            If mUninstallList Is Nothing Then ScanUninstall()
            Return mUninstallList
        End Get
    End Property

    Private Shared mUninstallString = Nothing
    Public Shared ReadOnly Property UninstallString As String
        Get
            If mUninstallString Is Nothing Then ScanUninstall()
            Return mUninstallString
        End Get
    End Property

    ' dgp rev 2/17/2012 Text Report
    Private Shared mKeyMatchList As ArrayList = Nothing
    Public Shared ReadOnly Property KeyMatchList As ArrayList
        Get
            If mKeyMatchList Is Nothing Then RegistryScan()
            Return mKeyMatchList
        End Get
    End Property

    Private Shared mLogPath = Nothing

    ' dgp rev 3/21/2012
    Private Shared Sub InstallMSI(fullspec)

        Dim PVWCmd = New Process

        If Not System.IO.File.Exists(fullspec) Then
            MsgBox(String.Format("MSI not found - {0}", fullspec))
        Else
            mCurAppFile = System.IO.Path.GetFileName(mPIA)
            mCurAppPath = System.IO.Path.GetDirectoryName(mPIA)

            ' a new process is created from running the PV-Wave program
            PVWCmd.StartInfo.UseShellExecute = False
            PVWCmd.StartInfo.UserName = "aaplugged"
            PVWCmd.StartInfo.Domain = "NIH"

            Dim pass = "Cracker55jack%%"

            Dim secret As New System.Security.SecureString
            Dim letter
            For Each letter In pass
                secret.AppendChar(letter)
            Next

            PVWCmd.StartInfo.Password = secret

            '        PVWCmd.StartInfo.FileName = System.IO.Path.Combine(mCurAppPath, mCurAppFile)
            PVWCmd.StartInfo.FileName = mCurAppFile
            PVWCmd.StartInfo.WorkingDirectory = mCurAppPath

            mLogPath = System.Environment.GetEnvironmentVariable("APPDATA")
            mLogPath = System.IO.Path.Combine(mLogPath, "Flow Control")
            Dim mLogFullSpec = System.IO.Path.Combine(mLogPath, "Obsolete.Log")

            PVWCmd.StartInfo.Arguments = String.Format(" /I  /quiet ")
            '        PVWCmd.StartInfo.Arguments = String.Format(" /I ""{0}"" /quiet > ""{1}""", mObsolete, mLogFullSpec)

            PVWCmd.StartInfo.CreateNoWindow = False

            PVWCmd.Start()
        End If

    End Sub



    ' dgp rev 3/21/2012
    Private Shared Sub InstallObsolete()

        Dim PVWCmd = New Process

        ' a new process is created from running the PV-Wave program
        PVWCmd.StartInfo.UseShellExecute = False
        PVWCmd.StartInfo.UserName = "aaplugged"
        PVWCmd.StartInfo.Domain = "NIH"

        Dim pass = "Cracker55jack%%"

        Dim secret As New System.Security.SecureString
        Dim letter
        For Each letter In pass
            secret.AppendChar(letter)
        Next

        PVWCmd.StartInfo.Password = secret

        '        PVWCmd.StartInfo.FileName = System.IO.Path.Combine(mCurAppPath, mCurAppFile)
        PVWCmd.StartInfo.FileName = "msiexec.exe"
        PVWCmd.StartInfo.WorkingDirectory = mCurAppPath
        mLogPath = System.Environment.GetEnvironmentVariable("APPDATA")
        mLogPath = System.IO.Path.Combine(mLogPath, "Flow Control")
        Dim mLogFullSpec = System.IO.Path.Combine(mLogPath, "Obsolete.Log")

        PVWCmd.StartInfo.Arguments = String.Format(" /I ""{0}"" /quiet ", mObsolete)
        '        PVWCmd.StartInfo.Arguments = String.Format(" /I ""{0}"" /quiet > ""{1}""", mObsolete, mLogFullSpec)

        PVWCmd.StartInfo.CreateNoWindow = False

        PVWCmd.Start()

    End Sub

    ' dgp rev 3/21/2012
    Private Shared Sub UserUnInstallApp(cmd)

        Dim PVWCmd = New Process

        ' a new process is created from running the PV-Wave program
        PVWCmd.StartInfo.UseShellExecute = False

        Dim info = cmd.ToString.Split(" ")
        Dim arg = cmd.ToString.Substring(info(0).Length)

        ' redirect IO
        ' PVWCmd.StartInfo.RedirectStandardOutput = True
        '       PVWCmd.StartInfo.RedirectStandardError = True
        '      PVWCmd.StartInfo.RedirectStandardInput = True
        ' don't even bring up the console window
        PVWCmd.StartInfo.CreateNoWindow = False
        ' executable command line info


        PVWCmd.StartInfo.FileName = info(0)

        '        PVWCmd.StartInfo.EnvironmentVariables("pvwave_setup_flag") = "skip"

        '        PVWCmd.StartInfo.RedirectStandardInput = True

        '        PVWCmd.StartInfo.WorkingDirectory = Dist_Str
        '        PVWCmd.StartInfo.WorkingDirectory = mCurAppPath
        PVWCmd.StartInfo.Arguments = arg


        '            PVWCmd.EnableRaisingEvents = True
        '   
        ' Add an event handler.
        '
        '           AddHandler PVWCmd.Exited, AddressOf Me.ProcessExited

        PVWCmd.Start()
        PVWCmd.WaitForExit()

        RaiseEvent RemovalEventDoneHandler()

    End Sub



    ' dgp rev 3/21/2012
    Private Shared Sub PrivUnInstallApp(cmd)

        Dim PVWCmd = New Process

        ' a new process is created from running the PV-Wave program
        PVWCmd.StartInfo.UseShellExecute = False
        PVWCmd.StartInfo.UserName = "aaplugged"
        PVWCmd.StartInfo.Domain = "NIH"

        Dim pass = "Cracker44jack$$"

        Dim secret As New System.Security.SecureString
        Dim letter
        For Each letter In pass
            secret.AppendChar(letter)
        Next

        PVWCmd.StartInfo.Password = secret

        Dim info = cmd.ToString.Split(" ")
        Dim arg = cmd.ToString.Substring(info(0).Length)

        ' redirect IO
        ' PVWCmd.StartInfo.RedirectStandardOutput = True
        '       PVWCmd.StartInfo.RedirectStandardError = True
        '      PVWCmd.StartInfo.RedirectStandardInput = True
        ' don't even bring up the console window
        PVWCmd.StartInfo.CreateNoWindow = False
        ' executable command line info


        PVWCmd.StartInfo.FileName = info(0)

        '        PVWCmd.StartInfo.EnvironmentVariables("pvwave_setup_flag") = "skip"

        '        PVWCmd.StartInfo.RedirectStandardInput = True

        '        PVWCmd.StartInfo.WorkingDirectory = Dist_Str
        '        PVWCmd.StartInfo.WorkingDirectory = mCurAppPath
        PVWCmd.StartInfo.Arguments = arg


        '            PVWCmd.EnableRaisingEvents = True
        '   
        ' Add an event handler.
        '
        '           AddHandler PVWCmd.Exited, AddressOf Me.ProcessExited

        PVWCmd.Start()
        PVWCmd.WaitForExit()

        RaiseEvent RemovalEventDoneHandler()

    End Sub

    Private Shared mCurAppFile As String
    Private Shared mCurAppPath As String

    Private Shared mObsolete As String = "\\nt-eib-10-6b16\www\BetaDistribution\RunTrackerDeploy.msi"
    Private Shared mPIA As String = "\\nt-eib-10-6b16\www\BetaDistribution\office2003piaredist\o2003pia.msi"

    Public Shared Sub InstallApp()

        Dim PVWCmd = New Process

        ' a new process is created from running the PV-Wave program
        PVWCmd.StartInfo.UseShellExecute = False
        PVWCmd.StartInfo.UserName = "aaplugged"
        PVWCmd.StartInfo.Domain = "NIH"

        Dim pass = "Cracker55jack%"

        Dim secret As New System.Security.SecureString
        Dim letter
        For Each letter In pass
            secret.AppendChar(letter)
        Next

        PVWCmd.StartInfo.Password = secret


        ' a new process is created from running the PV-Wave program
        PVWCmd.StartInfo.UseShellExecute = True
        ' redirect IO
        ' PVWCmd.StartInfo.RedirectStandardOutput = True
        '       PVWCmd.StartInfo.RedirectStandardError = True
        '      PVWCmd.StartInfo.RedirectStandardInput = True
        ' don't even bring up the console window
        PVWCmd.StartInfo.CreateNoWindow = False
        ' executable command line info

        PVWCmd.StartInfo.FileName = mCurAppFile

        '        PVWCmd.StartInfo.EnvironmentVariables("pvwave_setup_flag") = "skip"

        '        PVWCmd.StartInfo.RedirectStandardInput = True

        '        PVWCmd.StartInfo.WorkingDirectory = Dist_Str
        PVWCmd.StartInfo.WorkingDirectory = mCurAppPath
        mLogPath = System.Environment.GetEnvironmentVariable("APPDATA")
        mLogPath = System.IO.Path.Combine(mLogPath, "Flow Control")
        Dim mLogFullSpec = System.IO.Path.Combine(mLogPath, String.Format("{0}.Log", System.IO.Path.GetFileNameWithoutExtension(mCurAppFile)))
        PVWCmd.StartInfo.Arguments = String.Format(" > {0}", mLogFullSpec.ToString)


        '            PVWCmd.EnableRaisingEvents = True
        '   
        ' Add an event handler.
        '
        '           AddHandler PVWCmd.Exited, AddressOf Me.ProcessExited

        PVWCmd.Start()

    End Sub

    Private Shared mSuiteDistribution = "\\nt-eib-10-6b16\Software\EIB Flow Lab"
    Public Shared ReadOnly Property SuiteDistribution As String
        Get
            Return mSuiteDistribution
        End Get
    End Property

    Public Shared Function CheckPath(item) As Boolean

        mCurAppFile = ""
        Dim AppPath = System.IO.Path.Combine(mSuiteDistribution, item)
        Try
            Dim fld
            Dim fil
            For Each fld In System.IO.Directory.GetDirectories(AppPath)
                For Each fil In System.IO.Directory.GetFiles(fld, "*.application")
                    mCurAppPath = fld
                    mCurAppFile = System.IO.Path.GetFileName(fil)
                Next
            Next
        Catch ex As Exception

        End Try
        Return (Not mCurAppFile = "")

    End Function

    Public Shared ReadOnly Property TrackerInstallFlag As Boolean
        Get
            Return mTrackerInstall
        End Get
    End Property

    Public Shared ReadOnly Property TrackerRemoveFlag As Boolean
        Get
            Return mTrackerRemove
        End Get
    End Property

    Public Shared Sub DebugReinstallObsolete()

        mCurAppFile = System.IO.Path.GetFileName(mObsolete)
        mCurAppPath = System.IO.Path.GetDirectoryName(mObsolete)
        InstallObsolete()

    End Sub

    Delegate Sub RemovalEvent()
    Public Shared Event RemovalEventDoneHandler As RemovalEvent

    Public Shared Sub RemovalLauncher()

        AddHandler RemovalEventDoneHandler, AddressOf EmailResults
        Dim objThread As New Thread(New ThreadStart(AddressOf RemoveObsoleteTracker))
        objThread.Name = "Remove Tracker"
        objThread.Start()

    End Sub

    Private Shared Sub EmailResults()

        Dim objEmail As New EmailReporting
        objEmail.SendLog(mLogFullSpec)

    End Sub

    Private Shared mLogFullSpec

    Public Shared Sub RemoveObsoleteTracker()

        mLogPath = System.Environment.GetEnvironmentVariable("APPDATA")
        mLogPath = System.IO.Path.Combine(mLogPath, "Flow Control")
        mLogFullSpec = System.IO.Path.Combine(mLogPath, "TrackerRemoval.Log")
        mTrackerRemovalString = mTrackerRemovalString.ToUpper.Replace("/I", "/X")
        mTrackerRemovalString += String.Format(" /quiet /log ""{0}""", mLogFullSpec)
        Dim results = Shell(mTrackerRemovalString, vbNormalFocus)
        'UserUnInstallApp(mTrackerRemovalString)

    End Sub

    ' dgp rev 6/7/2012
    Public Shared Sub InstallActive()

        If mTrackerInstall Then
            mCurAppFile = System.IO.Path.GetFileName(mActiveInstallString)
            mCurAppPath = System.IO.Path.GetDirectoryName(mActiveInstallString)
            InstallApp()
        End If

    End Sub


    ' dgp rev 6/7/2012
    Public Shared Sub InstallPIA()

        mCurAppFile = System.IO.Path.GetFileName(mPIA)
        mCurAppPath = System.IO.Path.GetDirectoryName(mPIA)
        InstallApp()

    End Sub


    ' dgp rev 6/7/2012
    Public Shared Sub InstallnewTracker()

        If mTrackerInstall Then
            mCurAppFile = System.IO.Path.GetFileName(mTrackerInstallString)
            mCurAppPath = System.IO.Path.GetDirectoryName(mTrackerInstallString)
            InstallApp()
        End If

    End Sub

    ' dgp rev 6/7/2012
    Public Shared Sub AnalyzeBeta()

        mTrackerInstall = True
        mTrackerRemove = False
        Dim record

        For Each record In KeyMatchList
            If record("DisplayName").ToString.ToLower.Contains("beta") Then
                If record("DisplayName").ToString.ToLower.Contains("deploy") Then
                    If record("UninstallString").ToString.ToLower.Contains("dfshim") Then
                        mTrackerRemovalString = record("UninstallString")
                        mTrackerRemove = True
                    Else
                        mTrackerInstall = False
                    End If
                End If
            End If
        Next

    End Sub



    ' dgp rev 6/7/2012
    Public Shared Sub AnalyzeTracker()

        mTrackerInstall = True
        mTrackerRemove = False
        Dim record

        For Each record In KeyMatchList
            If record("DisplayName").ToString.ToLower.Contains("tracker") Then
                If record("UninstallString").ToString.ToLower.Contains("msiexec") Then
                    mTrackerRemovalString = record("UninstallString")
                    mTrackerRemove = True
                Else
                    mTrackerInstall = False
                End If
            End If
        Next

        If SoftwareHandler.TrackerRemoveFlag Then
            If MsgBox("Obsolete version of Run Tracker needs to be removed", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                SoftwareHandler.RemovalLauncher()
            End If
        End If

        If SoftwareHandler.TrackerInstallFlag Then
            If MsgBox("Install New Version of Run Tracker", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                SoftwareHandler.InstallnewTracker()
            End If
        End If

    End Sub

End Class
