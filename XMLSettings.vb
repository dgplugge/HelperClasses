' Name:     XMLStore
' Author:   Donald G Plugge
' Date:     5/2/2012
' Purpose:  XML Store of information
' <{Filename}>
'   <Global>
'       <{settings} {keyword=value>
'   <Users>
'       <{username}>
'          <{settings} {keyword =value>
Imports System.Threading
Imports Microsoft.Win32
Imports System.IO
Imports System.Xml.Linq
Imports HelperClasses
Imports System.Linq

Public Class XMLSettings

    Private mMessage As String

    Private mCurElementName = "Settings"
    Private mCurAttributeName = ""

    Public Property CurElementName As String
        Get
            Return mCurElementName
        End Get
        Set(ByVal value As String)
            mCurElementName = value
        End Set
    End Property

    Public Property CurAttributeName As String
        Get
            Return mCurAttributeName
        End Get
        Set(ByVal value As String)
            mCurAttributeName = value
        End Set
    End Property

    Public ReadOnly Property CurAttributeValue(ByVal user As String) As String
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Users").Elements(CurUser).Elements(CurElementName).Attributes(CurAttributeName) Select item).ToList
            If lst.Count = 0 Then Return ""
            Return lst(0).Value
        End Get
    End Property

    Public ReadOnly Property CurAttributeValue As String
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Global").Elements(CurElementName).Attributes(CurAttributeName) Select item).ToList
            If lst.Count = 0 Then Return ""
            Return lst(0).Value
        End Get
    End Property

    ' dgp rev 5/2/2012 Path to Server
    Private mShareName As String = "Distribution"
    Public ReadOnly Property GlobalShareName As String
        Get
            Return mShareName
        End Get
    End Property

    ' dgp rev 5/2/2012 
    Private mGlobalServerPath = Nothing
    Public ReadOnly Property GlobalServerPath() As String
        Get
            If mGlobalServerPath Is Nothing Then
                mGlobalServerPath = String.Format("\\{0}\{1}", "NT-EIB-10-6B16", GlobalShareName)
                mGlobalServerPath = System.IO.Path.Combine(mGlobalServerPath, "Versions")
            End If
            Return mGlobalServerPath
        End Get
    End Property

    ' dgp rev 5/2/2012 
    Public Function ClearData() As Boolean

        ClearData = True
        If System.IO.File.Exists(FullFileSpec) Then
            Try
                System.IO.File.Delete(FullFileSpec)
            Catch ex As Exception
                ClearData = False
            End Try
        End If

    End Function

    ' dgp rev 5/3/2012 
    Private mGroupName As String = "FlowControl"
    Public Property GroupName As String
        Get
            Return mGroupName
        End Get
        Set(ByVal value As String)
            mGroupName = value
        End Set
    End Property

    ' dgp rev 5/3/2012 
    Private mCurXMLDoc = Nothing

    ' dgp rev 5/3/2012 
    Public ReadOnly Property CurXMLDoc As XElement
        Get
            If mCurXMLDoc Is Nothing Then
                If System.IO.File.Exists(FullFileSpec) Then
                    mCurXMLDoc = XElement.Load(FullFileSpec)
                    mGroupName = mCurXMLDoc.Name.LocalName.ToString
                Else
                    mGroupName = System.IO.Path.GetFileNameWithoutExtension(FullFileSpec)
                    mCurXMLDoc = New XElement(GroupName)
                End If
            End If
            Return mCurXMLDoc
        End Get
    End Property

    Private mFound

    ' Instance portion
    ' dgp rev 5/2/2012 Path to Server
    Private mCurShareName As String = "Upload"
    Public Property CurShareName As String
        Get
            Return mCurShareName
        End Get
        Set(ByVal value As String)
            mCurShareName = System.IO.Path.GetDirectoryName(value)
        End Set

    End Property

    Private mValid As Boolean = False

    ' dgp rev 5/2/2012 
    Private mCurServerPath = Nothing
    Private mCurFileName = String.Format("{0}.xml", Environment.UserName)
    Public Property CurFileName As String
        Get
            Return mCurFileName
        End Get
        Set(ByVal value As String)
            mCurFileName = String.Format("{0}.xml", value.ToLower.Replace(".xml", ""))
        End Set
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property CurServerPath() As String
        Get
            If mCurServerPath Is Nothing Then
                mCurServerPath = String.Format("\\{0}\{1}", "NT-EIB-10-6B16", CurShareName)
                mCurServerPath = System.IO.Path.Combine(mGlobalServerPath, "Tracking")
            End If
            Return mCurServerPath
        End Get
    End Property

    Private Function SaveXML() As Boolean

        SaveXML = True
        Try
            mCurXMLDoc.Save(FullFileSpec)
            mCurXMLDoc = Nothing
        Catch ex As Exception
            SaveXML = False
        End Try

    End Function

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property RemoveElementWithAttr(user As String, ByVal key As String, val As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(key) Select item).ToList
            If lst.Count = 0 Then Return True
            Dim remov = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(key) Where item.Value = val Select item.Parent).First
            remov.Remove()
            SaveXML()
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property RemoveElementWithAttr(ByVal key As String, ByVal val As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Global").Elements(CurElementName).Attributes(key) Select item).ToList
            If lst.Count = 0 Then Return True
            Dim remov = (From item In orig.Elements("Global").Elements(CurElementName).Attributes(key) Where item.Value = val Select item.Parent).First
            remov.Remove()
            SaveXML()
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddValue(ByVal key As String, ByVal newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If Not orig.Elements("Global").Count = 0 Then
                If Not orig.Elements("Global").Elements(CurElementName).Count = 0 Then
                    If Not orig.Elements("Global").Elements(CurElementName).Attributes(key).Count = 0 Then
                        orig.Elements("Global").Elements(CurElementName).Attributes(key).First.Value = newval
                    Else
                        Dim newele As XElement = New XElement(CurElementName, New XAttribute(key, newval))
                        orig.Elements("Global").First.Add(newele)
                    End If
                Else
                    Dim newele As XElement = New XElement(CurElementName, New XAttribute(key, newval))
                    orig.Elements("Global").First.Add(newele)
                End If
            Else
                Dim newele As XElement = New XElement("Global", New XElement(CurElementName, New XAttribute(key, newval)))
                orig.Add(newele)
            End If
            SaveXML()

            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddValue(ByVal key As ArrayList, ByVal newval As ArrayList) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If Not key.Count = newval.Count Then Return False
            Dim replace = True

            Dim NewEle1 As New XElement(CurElementName)
            For idx = 0 To key.Count - 1
                NewEle1.Add(New XAttribute(XName.Get(key(idx)), newval(idx)))
            Next

            If Not orig.Elements("Global").Count = 0 Then
                If Not orig.Elements("Global").Elements(CurElementName).Count = 0 Then
                    orig.Elements("Global").First.Add(NewEle1)
                Else
                    orig.Elements("Global").First.Add(NewEle1)
                End If
            Else
                Dim newele As XElement = New XElement("Global", NewEle1)
                orig.Add(newele)
            End If
            SaveXML()

            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddValue(user As String, ByVal key As ArrayList, ByVal newval As ArrayList) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If Not key.Count = newval.Count Then Return False

            Dim NewEle1 As New XElement(CurElementName)
            For idx = 0 To key.Count - 1
                NewEle1.Add(New XAttribute(XName.Get(key(idx)), newval(idx)))
            Next

            If Not orig.Elements("Users").Count = 0 Then
                If Not orig.Elements("Users").Elements(user).Count = 0 Then
                    If Not orig.Elements("Users").Elements(user).Elements(CurElementName).Count = 0 Then
                        orig.Elements("Users").Elements(user).First.Add(NewEle1)
                    Else
                        orig.Elements("Users").Elements(user).First.Add(NewEle1)
                    End If
                Else
                    Dim newele As XElement = New XElement(user, NewEle1)
                    orig.Elements("Users").First.Add(newele)
                End If
            Else
                Dim newele As XElement = New XElement("Users", New XElement(user, NewEle1))
                orig.Add(newele)
            End If
            SaveXML()

            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property ReplaceValues(ByVal key As String, val As String, keyarr As ArrayList, ByVal valarr As ArrayList) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            Dim parents
            Dim kid As XElement

            Dim NewEle1 As New XElement(CurElementName)
            For idx = 0 To keyarr.Count - 1
                NewEle1.Add(New XAttribute(XName.Get(keyarr(idx)), valarr(idx)))
            Next

            parents = (From item In orig.Elements("Global").Elements(CurElementName) Where item.Attribute(XName.Get(key)).Value = val Select item).ToList
            For Each kid In parents
                kid.ReplaceWith(NewEle1)
            Next

            SaveXML()

            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property ReplaceValues(user As String, ByVal key As String, val As String, keyarr As ArrayList, ByVal valarr As ArrayList) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            Dim parents
            Dim kid As XElement

            Dim NewEle1 As New XElement(CurElementName)
            For idx = 0 To keyarr.Count - 1
                NewEle1.Add(New XAttribute(XName.Get(keyarr(idx)), valarr(idx)))
            Next

            parents = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName) Where item.Attribute(XName.Get(key)).Value = val Select item).ToList
            For Each kid In parents
                kid.ReplaceWith(NewEle1)
            Next

            SaveXML()

            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property ModValue(ByVal key As ArrayList, ByVal newval As ArrayList) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If Not key.Count = newval.Count Then Return False
            Dim replace = True
            Dim ReplaceIndex = 0
            For idx = 0 To key.Count - 1
                Dim lst = (From item In orig.Elements("Global").Elements(CurElementName).Attributes(XName.Get(key(idx))) Select item).ToList
                If Not lst.Count = 0 Then ReplaceIndex += 1
            Next

            If ReplaceIndex = key.Count Then
                For idx = 0 To key.Count - 1
                    orig.Elements("Global").Elements(CurElementName).Attributes(XName.Get(key(idx))).First.Value = newval(idx)
                Next
            Else
                Dim NewEle1 As New XElement(CurElementName)
                For idx = 0 To key.Count - 1
                    NewEle1.Add(New XAttribute(XName.Get(key(idx)), newval(idx)))
                Next

                If Not orig.Elements("Global").Count = 0 Then
                    If Not orig.Elements("Global").Elements(CurElementName).Count = 0 Then
                        orig.Elements("Global").First.Add(NewEle1)
                    Else
                        orig.Elements("Global").First.Add(NewEle1)
                    End If
                Else
                    Dim newele As XElement = New XElement("Global", NewEle1)
                    orig.Add(newele)
                End If
            End If
            SaveXML()

            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property ModValue(user As String, ByVal key As ArrayList, ByVal newval As ArrayList) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If Not key.Count = newval.Count Then Return False
            Dim replace = True
            Dim ReplaceIndex = 0
            For idx = 0 To key.Count - 1
                Dim lst = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(XName.Get(key(idx))) Select item).ToList
                If Not lst.Count = 0 Then ReplaceIndex += 1
            Next

            If ReplaceIndex = key.Count Then
                For idx = 0 To key.Count - 1
                    orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(XName.Get(key(idx))).First.Value = newval(idx)
                Next
            Else
                Dim NewEle1 As New XElement(CurElementName)
                For idx = 0 To key.Count - 1
                    NewEle1.Add(New XAttribute(XName.Get(key(idx)), newval(idx)))
                Next

                If Not orig.Elements("Users").Elements(user).Count = 0 Then
                    If Not orig.Elements("Users").Elements(user).Elements(CurElementName).Count = 0 Then
                        orig.Elements("Users").Elements(user).First.Add(NewEle1)
                    Else
                        orig.Elements("Users").Elements(user).First.Add(NewEle1)
                    End If
                Else
                    Dim newele As XElement = New XElement("Users", New XElement(user, NewEle1))
                    orig.Add(newele)
                End If
            End If
            SaveXML()

            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AttributeComboMatch(user As String, ByVal sourceName As String, ByVal sourceValue As String, ByVal targetName As String) As String
        Get
            Dim orig As XElement = CurXMLDoc

            Dim lst = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName) Where item.Attribute(sourceName).Value.Equals(sourceValue) And item.Attribute(targetName) IsNot Nothing Select item).ToList
            If Not lst.Count = 0 Then Return lst.First.Attribute(targetName).Value
            Return AttributeComboMatch(sourceName, sourceValue, targetName)

        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AttributeComboMatch(ByVal sourceName As String, ByVal sourceValue As String, ByVal targetName As String) As String
        Get
            Dim orig As XElement = CurXMLDoc

            Dim lst = (From item In orig.Elements("Global").Elements(CurElementName) Where item.Attribute(sourceName).Value.Equals(sourceValue) And item.Attribute(targetName) IsNot Nothing Select item).ToList
            If lst.Count = 0 Then Return ""
            Return lst.First.Attribute(targetName).Value
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AttributeComboExists(ByVal sourceName As String, ByVal sourceValue As String, ByVal targetName As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Global").Elements(CurElementName) Where item.Attribute(sourceName).Value.Equals(sourceValue) And item.Attribute(targetName) IsNot Nothing Select item).ToList
            Return Not lst.Count = 0
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AttributeComboExists(user As String, ByVal sourceName As String, ByVal sourceValue As String, ByVal targetName As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName) Where item.Attribute(sourceName).Value.Equals(sourceValue) And item.Attribute(targetName) IsNot Nothing Select item).ToList
            If Not lst.Count = 0 Then Return True
            Return AttributeComboExists(sourceName, sourceValue, targetName)
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddUpgrade(ByVal user As String, ByVal Uninstall As String, ByVal Install As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If CheckUser(user) Then
                Dim items = (From item In orig.Elements(user).Elements(SoftwareUpgrade.UpgradeStore.CurElementName).Attributes("Uninstall") Select item).ToList
                If items.Count = 0 Then
                    Dim newele As XElement = New XElement(SoftwareUpgrade.UpgradeStore.CurElementName, New XAttribute("Uninstall", Uninstall), New XAttribute("Install", Install))
                    orig.Element(user).Add(newele)
                    SaveXML()
                Else
                    orig.Elements(user).Elements(SoftwareUpgrade.UpgradeStore.CurElementName).Attributes("Uninstall").First.Value = Uninstall
                    orig.Elements(user).Elements(SoftwareUpgrade.UpgradeStore.CurElementName).Attributes("Install").First.Value = Install
                    SaveXML()
                End If
            Else
                Dim newele As XElement = New XElement(SoftwareUpgrade.UpgradeStore.CurElementName, New XAttribute("Uninstall", Uninstall), New XAttribute("Install", Install))
                orig.Add(newele)
                SaveXML()
            End If
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddValue(ByVal user As String, ByVal key As String, ByVal newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            mCurUser = user
            If Not orig.Elements("Users").Count = 0 Then
                If Not orig.Elements("Users").Elements(CurUser).Count = 0 Then
                    If Not orig.Elements("Users").Elements(CurUser).Elements(CurElementName).Count = 0 Then
                        If Not orig.Elements("Users").Elements(CurUser).Elements(CurElementName).Attributes(key).Count = 0 Then
                            orig.Elements("Users").Elements(CurUser).Elements(CurElementName).Attributes(key).First.Value = newval
                        Else
                            Dim newele As XElement = New XElement(CurElementName, New XAttribute(key, newval))
                            orig.Elements("Users").Elements(CurUser)(0).Add(newele)
                        End If
                    Else
                        Dim newele As XElement = New XElement(CurElementName, New XAttribute(key, newval))
                        orig.Elements("Users").Elements(CurUser)(0).Add(newele)
                    End If
                Else
                    Dim newele As XElement = New XElement(CurUser, New XElement(CurElementName, New XAttribute(key, newval)))
                    orig.Elements("Users")(0).Add(newele)
                End If
            Else
                Dim newele As XElement = New XElement("Users", New XElement(CurUser, New XElement(CurElementName, New XAttribute(key, newval))))
                orig.Add(newele)
            End If
            SaveXML()
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property ModifyValue(ByVal user As String, ByVal key As String, ByVal newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim items = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(key) Select item)
            If items Is Nothing Then Return False
            orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(key).First.Value = newval
            SaveXML()
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify global keyword
    Public ReadOnly Property ModifyValue(ByVal key As String, ByVal newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim items = (From item In orig.Elements("Global").Elements(CurElementName).Attributes(key) Select item)
            If items Is Nothing Then Return False
            orig.Elements("Global").Elements(CurElementName).Attributes(key).First.Value = newval
            SaveXML()
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 global key value
    Public ReadOnly Property GetValue(ByVal key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc

            Try
                Dim items = (From item In orig.Elements("Global").Elements(CurElementName).Attributes(key) Select item.Value).ToList
                If items IsNot Nothing Then
                    Return items(0).ToString
                End If
            Catch ex As Exception
                Return ""
            End Try
            Return ""
        End Get
    End Property

    Public Function CheckUser(ByVal user As String) As Boolean

        CheckUser = False
        Dim orig As XElement = CurXMLDoc

        Dim items = (From item In orig.Elements("Users").Elements(user)).ToList
        Return Not (items.Count = 0)

    End Function

    ' dgp rev 8/29/2012 specific user key value
    Public ReadOnly Property GetValue(ByVal user As String, ByVal key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc

            If Not CheckUser(user) Then Return ""

            Try
                Dim items = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(key) Select item.Value).ToList
                If items.Count = 0 Then Return ""
                Return items(0).ToString
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property

    ' dgp rev 8/29/2012 specific user key value
    Public ReadOnly Property GetAttributes(ByVal user As String) As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            GetAttributes = New ArrayList
            If CheckUser(user) Then
                Try
                    Dim items = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes Select item).ToList
                    If items.Count = 0 Then Return GetAttributes
                    For Each item In items
                        GetAttributes.Add(item)
                    Next
                Catch ex As Exception
                End Try
            End If
            Return GetAttributes
        End Get
    End Property

    ' dgp rev 8/29/2012 specific user key value
    Public ReadOnly Property GetAttribute(ByVal user As String) As String
        Get
            Dim orig As XElement = CurXMLDoc

            If Not CheckUser(user) Then Return ""

            Try
                Dim items = (From item In orig.Elements(user).Elements(CurElementName).Attributes(CurAttributeName) Select item).ToList
                If items.Count = 0 Then Return ""
                Return items(0).Value
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property



    ' dgp rev 8/29/2012 specific user key value
    Public ReadOnly Property GetAttribute(ByVal user As String, ByVal key As String) As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc

            Dim Attribes As New ArrayList
            If Not CheckUser(user) Then Return Attribes

            Try
                Dim items = (From item In orig.Elements(user).Elements(key).Attributes Select item).ToList
                If items.Count = 0 Then Return Attribes
                Return Attribes
            Catch ex As Exception
                Return Attribes
            End Try
        End Get
    End Property

    ' dgp rev 8/29/2012 retrieve list of users
    Public ReadOnly Property GetUsers As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements() Where Not item.Name = CurElementName Select item).ToList
            If items IsNot Nothing Then
                If items.Count > 0 Then
                    Dim item As XElement
                    For Each item In items
                        NewArray.Add(item.Name.LocalName)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property

    ' dgp rev 8/29/2012 retrieve list of keys for select user
    Public ReadOnly Property GetElements(ByVal user As String) As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements("Users").Elements(user) Select item).ToList
            If items IsNot Nothing Then
                If items.Count > 0 Then
                    Dim item As XElement
                    For Each item In items.Elements
                        If Not NewArray.Contains(item.Name.LocalName) Then NewArray.Add(item.Name.LocalName)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property

    ' dgp rev 8/29/2012 retrieve list of keys for select user
    Public ReadOnly Property GetSubElements(ByVal user As String) As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName) Select item).ToList
            If items IsNot Nothing Then
                If items.Count > 0 Then
                    Dim item As XElement
                    For Each item In items
                        NewArray.Add(item.Attributes.ElementAt(0).Name.LocalName)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property



    ' dgp rev 8/29/2012 retrieve list of global keys
    Public ReadOnly Property GetElements As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements("Global").Elements Select item).ToList
            If items IsNot Nothing Then
                If items.Count > 0 Then
                    Dim item As XElement
                    For Each item In items
                        NewArray.Add(item)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property AttribExists(ByVal user As String, ByVal key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements(user).Elements(CurElementName).Attributes(key) Select item).ToList
            Return Not (lst.Count = 0)
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property Attributes(user As String) As List(Of XAttribute)
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes Select item).ToList
            Return lst
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property Attributes() As List(Of XAttribute)
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Global").Elements(CurElementName).Attributes Select item).ToList
            Return lst
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property AttributeExists(ByVal user As String, ByVal attr As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes(attr) Select item).ToList
            Return Not (lst.Count = 0)
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property AttributesExists(user As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Users").Elements(user).Elements(CurElementName).Attributes Select item).ToList
            Return Not (lst.Count = 0)
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property AttributesExists() As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Global").Elements(CurElementName).Attributes Select item).ToList
            Return Not (lst.Count = 0)
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property AttributeExists(ByVal key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements("Global").Elements(CurElementName).Attributes(key) Select item).ToList
            Return Not (lst.Count = 0)
        End Get
    End Property

    Private mCurUser As String = Environment.UserName
    Public ReadOnly Property CurUser As String
        Get
            Return mCurUser
        End Get
    End Property

    ' dgp rev 5/2/2012 
    Public ReadOnly Property FullFileSpec As String
        Get
            Return mFullFileSpec
        End Get
    End Property

    Private mFullFileSpec As String

    ' dgp rev 5/23/2012 New XML persistence file
    Public Sub New(ByVal spec As String)

        mFullFileSpec = spec
        Dim uri As Uri = New Uri(spec)
        If uri.IsUnc Then
        Else
        End If
        CurShareName = uri.AbsolutePath

        CurFileName = System.IO.Path.GetFileNameWithoutExtension(spec)
        mValid = System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(spec))

    End Sub

    ' dgp rev 5/23/2012 New XML persistence file
    Public Sub New(ByVal Share As String, ByVal Filename As String)

        Dim uri As Uri = New Uri(Share)
        If uri.IsUnc Then
        Else
        End If
        mFullFileSpec = System.IO.Path.Combine(Share, String.Format("{0}.xml", Filename.ToLower.Replace(".xml", "")))
        CurShareName = uri.AbsolutePath

        CurFileName = Filename
        mValid = System.IO.Directory.Exists(Share)

    End Sub
End Class
