' Name:     VerInfo
' Author:   Donald G Plugge`1   `                                                           
' Date:     3/8/2012
' Purpose:  PC Versioning Info
Imports Microsoft.Build.Tasks.Deployment.ManifestUtilities
Imports System.Deployment.Application
Imports System.Windows.Forms
Imports System.Collections.Specialized
Imports System.Web

Public Class VerInfo

    Private Shared mMyAss As System.Reflection.Assembly = Nothing
    Private Shared mMyGUID As Guid

    ' dgp rev 3/30/2011 
    Public Shared ReadOnly Property GetGUID() As String
        Get
            If mMyAss Is Nothing Then
                mMyAss = System.Reflection.Assembly.GetExecutingAssembly
                mMyGUID = mMyAss.ManifestModule.ModuleVersionId
            End If
            Return mMyGUID.ToString
        End Get
    End Property

    Private Shared mVerDeployed = Nothing
    Private Shared mFullVer As Version

    ' dgp rev 3/30/2011 
    Public Shared ReadOnly Property GetVersion() As String
        Get
            If mVerDeployed Is Nothing Then
                If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
                    mFullVer = Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
                Else
                    mFullVer = My.Application.Info.Version
                End If
                mVerDeployed = mFullVer.Major.ToString
                mVerDeployed = mVerDeployed + "." + mFullVer.Minor.ToString
                mVerDeployed = mVerDeployed + "." + mFullVer.Revision.ToString
            End If
            Return mVerDeployed

        End Get

    End Property

    Private Shared mManifest = Nothing
    ' dgp rev 1/30/2012

    Public Shared Function GetQueryStringParameters() As NameValueCollection
        Dim NameValueTable As New NameValueCollection()

        If (ApplicationDeployment.IsNetworkDeployed) Then
            Dim QueryString As String = ApplicationDeployment.CurrentDeployment.ActivationUri.Query
            NameValueTable = HttpUtility.ParseQueryString(QueryString)
        End If

        GetQueryStringParameters = NameValueTable
    End Function

    Public Shared ReadOnly Property Current7Location() As String
        Get
            Try

                If (ApplicationDeployment.IsNetworkDeployed) Then
                    Try
                        If ApplicationDeployment.CurrentDeployment IsNot Nothing Then
                            If ApplicationDeployment.CurrentDeployment.UpdatedApplicationFullName IsNot Nothing Then
                                Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
                                If AD.UpdatedApplicationFullName IsNot Nothing Then
                                    Return AD.UpdatedApplicationFullName.ToString.Replace(".application", "7.application")
                                Else
                                    Return "UpdateLocation is nothing"
                                End If
                            Else
                                Return "CurrentDeployment is nothing"
                            End If
                        Else
                            Return "CurrentDeployment is nothing"
                        End If
                    Catch ex As Exception
                        Return "info " + ex.Message
                    End Try
                End If
            Catch ex As Exception
                Return "IsNetwork error " + ex.Message
            End Try
            Return ""
        End Get
    End Property

    Private Shared mMockDeploy = "file://nt-eib-10-6b16/WWW/BetaDistribution/SoftwareReplacement.application#SoftwareReplacement.application, Version=1.0.0.74, Culture=neutral, PublicKeyToken=8e44a03860ffce06, processorArchitecture=msil/SoftwareReplacement.exe, Version=1.0.0.74, Culture=neutral, PublicKeyToken=8e44a03860ffce06, processorArchitecture=msil, type=win32"
    Private Shared mShortCut = "file://nt-eib-10-6b16/WWW/BetaDistribution/SoftwareReplacement.application#SoftwareReplacement.application, Culture=neutral, PublicKeyToken=8e44a03860ffce06, processorArchitecture=msil"

    Private Shared mCurrentUpdateLocation = Nothing
    Public Shared ReadOnly Property CurrentUpdateLocation() As String
        Get
            If mCurrentUpdateLocation Is Nothing Then
                mCurrentUpdateLocation = "//nt-eib-10-6b16/WWW/BetaDistribution/SoftwareReplacement.application"
                Try
                    If (ApplicationDeployment.IsNetworkDeployed) Then
                        Try
                            If ApplicationDeployment.CurrentDeployment IsNot Nothing Then
                                If ApplicationDeployment.CurrentDeployment.UpdatedApplicationFullName IsNot Nothing Then
                                    Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
                                    If AD.UpdatedApplicationFullName IsNot Nothing Then
                                        Dim Info = AD.UpdatedApplicationFullName.ToString()
                                        Dim file = Info.Split("#")
                                        Dim path = file(0).Split(":")
                                        mCurrentUpdateLocation = path(1)
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    Else
                        Dim Info = mMockDeploy
                        Dim file = Info.Split("#")
                        Dim path = file(0).Split(":")
                        mCurrentUpdateLocation = path(1)
                    End If
                Catch ex As Exception
                End Try
            End If
            Return mCurrentUpdateLocation
        End Get
    End Property


    Public Shared ReadOnly Property AssemblyFullInfo As String
        Get
            Try
                Dim path = System.IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath.ToString)
                If mManifest Is Nothing Then
                    Try
                        mManifest = ManifestReader.ReadManifest(String.Format("{0}.exe.manifest", path), True)
                        Return String.Format("{0} {1}", mManifest.AssemblyIdentity.Name, mManifest.AssemblyIdentity.Version)
                    Catch ex As Exception
                        Return String.Format("No Manifest for {0}", path)
                    End Try
                End If
                Try
                    Return String.Format("{0} {1}", mManifest.AssemblyIdentity.Name, mManifest.AssemblyIdentity.Version)
                Catch ex As Exception
                    Return String.Format("No Manifest for {0}", path)
                End Try
            Catch ex As Exception
                Return Application.ExecutablePath.ToString
            End Try
        End Get
    End Property

    Private Shared mAssemblyName = Nothing
    Public Shared ReadOnly Property AssemblyName As String
        Get
            Dim path = System.IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath.ToString)
            mAssemblyName = My.Application.Info.AssemblyName
            If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
                If mManifest Is Nothing Then
                    Try
                        mManifest = ManifestReader.ReadManifest(String.Format("{0}.exe.manifest", path), True)
                        mAssemblyName = mManifest.AssemblyIdentity.Name
                    Catch ex As Exception
                        mAssemblyName = My.Application.Info.AssemblyName
                    End Try
                End If
            End If
            Return mAssemblyName
        End Get
    End Property

    Public Shared ReadOnly Property AssemblyVersion As String
        Get
            Dim path = System.IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath.ToString)
            If mManifest Is Nothing Then
                Try
                    mManifest = ManifestReader.ReadManifest(String.Format("{0}.exe.manifest", path), True)
                    Return mManifest.AssemblyIdentity.Version
                Catch ex As Exception
                    Return String.Format("No Manifest for {0}", path)
                End Try
            End If
            Return mManifest.AssemblyIdentity.Version
        End Get
    End Property

    Public Shared ReadOnly Property Manifest() As Manifest

        Get
            If mManifest Is Nothing Then
                Try
                    mManifest = ManifestReader.ReadManifest(String.Format("{0}.exe.manifest", System.IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath.ToString)), True)
                Catch ex As Exception

                End Try
            End If

            Return mManifest
        End Get

    End Property


End Class
