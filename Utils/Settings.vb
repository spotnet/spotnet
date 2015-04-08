Imports System.Xml
Imports System.Configuration
Imports System.Collections.Specialized

Friend Class PortableSettingsProvider
    Inherits SettingsProvider

    Const SETTINGSROOT As String = "Settings" 'XML Root Node

    Public Overrides Sub Initialize(ByVal name As String, ByVal col As NameValueCollection)
        MyBase.Initialize(Me.ApplicationName, col)
    End Sub

    Public Overrides Property ApplicationName() As String
        Get
            Return Spotname
        End Get
        Set(ByVal value As String)
            'Do nothing
        End Set
    End Property

    Overridable Function GetAppSettingsPath() As String

        Return SettingsFolder()

    End Function

    Overridable Function GetAppSettingsFilename() As String
        'Used to determine the filename to store the settings
        Return "settings.xml"
    End Function

    Public Overrides Sub SetPropertyValues(ByVal context As SettingsContext, ByVal propvals As SettingsPropertyValueCollection)
        'Iterate through the settings to be stored
        'Only dirty settings are included in propvals, and only ones relevant to this provider
        For Each propval As SettingsPropertyValue In propvals
            SetValue(propval)
        Next

        Try
            SettingsXML.Save(IO.Path.Combine(GetAppSettingsPath, GetAppSettingsFilename))
        Catch ex As Exception
            'Ignore if cant save, device been ejected
        End Try
    End Sub

    Public Overrides Function GetPropertyValues(ByVal context As SettingsContext, ByVal props As SettingsPropertyCollection) As SettingsPropertyValueCollection
        'Create new collection of values
        Dim values As SettingsPropertyValueCollection = New SettingsPropertyValueCollection()

        'Iterate through the settings to be retrieved
        For Each setting As SettingsProperty In props

            Dim value As SettingsPropertyValue = New SettingsPropertyValue(setting)
            value.IsDirty = False
            value.SerializedValue = GetValue(setting)
            values.Add(value)
        Next
        Return values
    End Function

    Private m_SettingsXML As Xml.XmlDocument = Nothing

    Private ReadOnly Property SettingsXML() As Xml.XmlDocument
        Get
            'If we dont hold an xml document, try opening one.  
            'If it doesnt exist then create a new one ready.
            If m_SettingsXML Is Nothing Then
                m_SettingsXML = New Xml.XmlDocument

                Try
                    m_SettingsXML.Load(IO.Path.Combine(GetAppSettingsPath, GetAppSettingsFilename))
                Catch ex As Exception
                    'Create new document
                    Dim dec As XmlDeclaration = m_SettingsXML.CreateXmlDeclaration("1.0", "utf-8", String.Empty)
                    m_SettingsXML.AppendChild(dec)

                    Dim nodeRoot As XmlNode

                    nodeRoot = m_SettingsXML.CreateNode(XmlNodeType.Element, SETTINGSROOT, "")
                    m_SettingsXML.AppendChild(nodeRoot)
                End Try
            End If

            Return m_SettingsXML
        End Get
    End Property

    Private Function GetValue(ByVal setting As SettingsProperty) As String
        Dim ret As String = ""

        Try
            ret = SettingsXML.SelectSingleNode(SETTINGSROOT & "/" & setting.Name).InnerText

        Catch ex As Exception
            If Not setting.DefaultValue Is Nothing Then
                ret = setting.DefaultValue.ToString
            Else
                ret = ""
            End If
        End Try

        Return ret
    End Function

    Private Sub SetValue(ByVal propVal As SettingsPropertyValue)

        Dim SettingNode As Xml.XmlElement

        Try
            SettingNode = CType(SettingsXML.SelectSingleNode(SETTINGSROOT & "/" & propVal.Name), XmlElement)
        Catch ex As Exception
            SettingNode = Nothing
        End Try

        'Check to see if the node exists, if so then set its new value
        If Not SettingNode Is Nothing Then
            If propVal.SerializedValue Is Nothing Then
                SettingNode.InnerText = ""
            Else
                SettingNode.InnerText = propVal.SerializedValue.ToString
            End If

        Else

            'Store the value as an element of the Settings Root Node
            SettingNode = SettingsXML.CreateElement(propVal.Name)
            SettingNode.InnerText = propVal.SerializedValue.ToString
            SettingsXML.SelectSingleNode(SETTINGSROOT).AppendChild(SettingNode)

        End If

    End Sub

End Class
