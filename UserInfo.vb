Imports System.Xml.Linq
Imports System.Linq

Public Class UserInfo

    Private mUserName As String

    Private mInstall As String
    Public ReadOnly Property InstallAppName As String
        Get
            Return mInstall
        End Get
    End Property

    Private mUpgradeFlag As Boolean = False
    Public ReadOnly Property UpgradeFlag As Boolean
        Get
            Return mUpgradeFlag
        End Get
    End Property

    Private mCommands As Hashtable
    Public ReadOnly Property Commands As Hashtable
        Get
            Return mCommands
        End Get
    End Property

    Private mAttributes As Hashtable

    Private mValid As Boolean = False
    Public ReadOnly Property Valid As Boolean
        Get
            Return mValid
        End Get
    End Property

    Private mUpgradeRequired = Nothing
    Public ReadOnly Property UpgradeRequired As Boolean
        Get
            If mUpgradeRequired Is Nothing Then ScanUpgrades()
            Return mUpgradeRequired
        End Get
    End Property


    Private mUpgradeAppPath = Nothing
    Public ReadOnly Property UpgradeAppPath As String
        Get
            If mUpgradeAppPath Is Nothing Then ScanUpgrades()
            Return mUpgradeAppPath
        End Get
    End Property


    Public Sub ScanUpgrades()

        mUpgradeRequired = False

        mCommands = New Hashtable
        If Not SoftwareUpgrade.UpgradeStore.CheckUser(mUserName) Then Exit Sub
        mValid = True

        Dim UniqueCommand As New ArrayList
        Dim orig As XElement = SoftwareUpgrade.UpgradeStore.CurXMLDoc


        Dim x = SoftwareUpgrade.UpgradeStore.AttributeComboMatch(mUserName, "Uninstall", String.Format("{0}.application", VerInfo.AssemblyName), "Install")
        Dim items = (From item In orig.Elements(mUserName).Elements(SoftwareUpgrade.UpgradeStore.CurElementName) Select item).ToList
        If items IsNot Nothing Then
            If items.Count > 0 Then
                Dim command As XElement
                For Each command In items
                    Dim attribs = (From attrib In command.Attributes("Uninstall") Select attrib).ToList
                    If Not attribs.Count = 0 Then
                        If attribs(0).Value.ToLower = String.Format("{0}.application", VerInfo.AssemblyName).ToLower Then
                            Dim attribII = (From attrib In command.Attributes("Install") Select attrib).ToList
                            If Not attribII.Count = 0 Then
                                If Not attribII(0).Value.ToLower = attribs(0).Value.ToLower Then
                                    mUpgradeAppPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(HelperClasses.VerInfo.CurrentUpdateLocation), attribII(0).Value)
                                    mUpgradeRequired = System.IO.File.Exists(mUpgradeAppPath)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If

    End Sub

    Public Sub New(ByVal user As String)

        mUserName = user

    End Sub

End Class
