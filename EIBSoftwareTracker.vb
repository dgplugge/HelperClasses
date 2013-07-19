Imports HelperClasses

Public Class EIBSoftwareTracker

    Private Shared mWWWDistributionRoot = "\\nt-eib-10-6b16\WWW\BetaDistribution\"
    Public Shared ReadOnly Property WWWDistributionRoot As String
        Get
            Return mWWWDistributionRoot
        End Get
    End Property

    Private Shared mWWWDistributionList = Nothing
    Public Shared ReadOnly Property WWWDistributionList As ArrayList
        Get
            If mWWWDistributionList Is Nothing Then
                mWWWDistributionList = New ArrayList
                If System.IO.Directory.Exists(WWWDistributionRoot) Then
                    If Not System.IO.Directory.GetFiles(WWWDistributionRoot).Length = 0 Then
                        For Each fil In System.IO.Directory.GetFiles(WWWDistributionRoot, "*.application")
                            mWWWDistributionList.Add(System.IO.Path.GetFileName(fil))
                        Next
                    End If
                End If
            End If
            Return mWWWDistributionList
        End Get
    End Property

    

    Private Shared mServerPath As String = "\\nt-eib-10-6b16\Distribution\Versions"
    Public Shared ReadOnly Property ServerPath As String
        Get
            Return mServerPath
        End Get
    End Property

    Private Shared mConfigName = Nothing
    Public Shared Property ConfigName As String
        Get
            If mConfigName Is Nothing Then
                If Environment.UserName.ToLower = "plugged" Then
                    '                    mConfigName = "TestConfig.xml"
                    mConfigName = "SoftwareConfig.xml"
                Else
                    mConfigName = "SoftwareConfig.xml"
                End If
            End If
            Return mConfigName
        End Get
        Set(value As String)
            mXMLStore = Nothing
            mConfigName = String.Format("{0}.XML", System.IO.Path.GetFileNameWithoutExtension(value))
        End Set
    End Property

    ' dgp rev 8/29/2012 
    Private Shared mXMLStore = Nothing
    Public Shared ReadOnly Property XMLStore As XMLSettings
        Get
            If mXMLStore Is Nothing Then
                mXMLStore = New XMLSettings(System.IO.Path.Combine(ServerPath, ConfigName))
            End If
            Return mXMLStore
        End Get
    End Property

    Private Shared mCurUserInfo As UserInfo

    Private Shared mCurUser = Environment.UserName
    Public Shared Property CurUser As String
        Get
            Return mCurUser
        End Get
        Set(value As String)
            mCurUser = value
        End Set
    End Property

    Private Shared mCurElementName = Nothing
    Public Shared ReadOnly Property CurElementName As String
        Get
            If mCurElementName Is Nothing Then
                '            ConfigName = "SoftwareUpgrades.XML"
                If Environment.OSVersion.ToString.ToLower.Contains("nt 6") Then
                    mCurElementName = "Upgrade7"
                Else
                    mCurElementName = "Upgrade"
                End If
            End If
            Return mCurElementName
        End Get
    End Property

    Private Shared mCurApp = String.Format("{0}.application", System.IO.Path.GetFileNameWithoutExtension(VerInfo.AssemblyName))
    Public Shared ReadOnly Property CurApp As String
        Get
            Return mCurApp
        End Get
    End Property

    ' dgp rev 12/19/2012 
    Private Shared Sub ScanUpgrades()

        mUpgradeAppPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(HelperClasses.VerInfo.CurrentUpdateLocation), XMLStore.AttributeComboMatch(CurUser, "Uninstall", CurApp, "Install"))
        mUpgradeRequired = System.IO.File.Exists(mUpgradeAppPath)

    End Sub

    Private Shared mUpgradeRequired = Nothing
    Public Shared ReadOnly Property UpgradeRequired As Boolean
        Get
            If mUpgradeRequired Is Nothing Then ScanUpgrades()
            Return mUpgradeRequired
        End Get
    End Property


    Private Shared mUpgradeAppPath = Nothing
    Public Shared ReadOnly Property UpgradeAppPath As String
        Get
            If mUpgradeAppPath Is Nothing Then ScanUpgrades()
            Return mUpgradeAppPath
        End Get
    End Property

    Private Shared mUpgradeFlag = False
    Public Shared ReadOnly Property UpgradeFlag As Boolean
        Get
            Return mUpgradeFlag
        End Get
    End Property

    Private Shared Sub ErrorTracking(mess As String)

        Logger.Log_Info(mess)
        Logger.Flush()

    End Sub

    Private Shared mLogger As HelperClasses.Logger
    Public Shared Property Logger As HelperClasses.Logger
        Get
            If mLogger Is Nothing Then mLogger = New HelperClasses.Logger("EIBSoftwareTracker")
            Return mLogger
        End Get
        Set(ByVal value As HelperClasses.Logger)
            mLogger = value
        End Set
    End Property

    Private Shared Sub SendReport()

        If UpgradeFlag Then
            ErrorTracking(String.Format("Config File: {0}", XMLStore.FullFileSpec))
            Logger.Clean_Up()
            Dim report = New HelperClasses.EmailReporting
            report.Message = "Successful Upgrade"
            report.SendLog(Logger.FilePath)
        End If

    End Sub

    ' dgp rev 12/4/2012
    Public Shared Function DoesApplicationNeedUpgrade() As Boolean

        DoesApplicationNeedUpgrade = False
        Try
            XMLStore.CurElementName = CurElementName
            If XMLStore.AttributeComboExists(CurUser, "Uninstall", CurApp, "Install") Then
                Try
                    ErrorTracking(String.Format("User: {0} Application: {1}", CurUser, CurApp))
                    ErrorTracking(String.Format("Uninstall: {0}", SoftwareHandler.ActiveUninstallString))
                    If SoftwareHandler.ActiveUninstallString IsNot Nothing Then
                        Try
                            Dim insUninstall As New frmUninstallUpgrade
                            insUninstall.FormUnInstallString = SoftwareHandler.ActiveUninstallString.ToString
                            insUninstall.FormInstallString = UpgradeAppPath
                            insUninstall.ShowDialog()
                            ErrorTracking(String.Format("Upgrade: {0}", UpgradeAppPath))
                            DoesApplicationNeedUpgrade = SoftwareUpgrade.InstallSuccessfull And SoftwareUpgrade.UnInstallSuccessfull
                        Catch ex As Exception
                            DoesApplicationNeedUpgrade = False
                        End Try
                    Else
                        ErrorTracking("No uninstall string")
                        MsgBox(String.Format("No Active UninstallString", MsgBoxStyle.Information))
                    End If
                Catch ex As Exception
                    MsgBox(String.Format("{0} -- {1} in {2}", ex.Message, mCurUserInfo.InstallAppName, System.IO.Path.GetDirectoryName(HelperClasses.VerInfo.CurrentUpdateLocation)), MsgBoxStyle.Information)
                End Try
            End If
        Catch ex As Exception
            ErrorTracking(ex.Message)
        End Try
        SendReport()

    End Function



End Class
