<%
Class PageFabricator
    Private m_bodyStyle                 'For inline body styling
    Private m_bolDebugMode              'Boolean the determines if class is in debug mode
    Private m_bolFooterSent             'Boolean that determines if the food has already been sent
    Private m_bolHeaderSent             'Boolean that determines if the header has already been sent
    Private m_description               'String containing page description, if anym_description
    Private m_favIcon                   'String containing favicon (if any)
    Private m_footerTemplates           'Array containing templates to add to footer
    Private m_FSO                       'Generic File System Object
    Private m_headerHTML                'String containing html to display with header
    Private m_inlineCSS                 'String containing css styling to display with header
    Private m_javascripts               'Array containing JavaScript references
    Private m_jsInjectValues            'Dictionary containing javascript values to inject
    Private m_meta                      'Dictionary containing meta values
    Private m_pathShared                'Contains link to shared directory
    Private m_stylesheets               'Array containing StyleSheet references to include
    Private m_title
    Private m_titleSuffix

    Public Property Get bodyStyle
        bodyStyle = m_bodyStyle
    End Property
    
    Public Property Let bodyStyle(css)
        If NOT IsNull(m_bodyStyle) Then
            css = Trim(CStr(css))
            If LEN(css) = 0 Then
                css = Null
            End If
        End If
        m_bodyStyle = css
    End Property

    Public Property Get bolFooterSent
        bolFooterSent = m_bolFooterSent
    End Property

    Public Property Get bolHeaderSent
        bolHeaderSent = m_bolHeaderSent
    End Property

    Public Property Get DebugMode
        DebugMode = false
        If NOT IsNull(m_bolDebugMode) Then
            DebugMode = m_bolDebugMode
        End If
    End Property

    Public Property Let DebugMode(value)
        m_bolDebugMode = CBool(value)
    End Property

    Public Property Get Description
        Description = m_description
    End Property

    Public Property Let Description(value)
        m_description = CStr(Trim(value))
    End Property

    Public Property Get Title
        Title = m_title
    End Property

    Public Property Let Title(value)
        m_title = CStr(Trim(value))
    End Property

    Public Property Get TitleSuffix
        TitleSuffix = m_titleSuffix
    End Property

    Public Property Let TitleSuffix(value)
        m_titleSuffix = CStr(Trim(value))
    End Property

    Public Function footer_addHtml(strValue)
        If NOT IsEmpty(strValue) AND NOT IsNull(strValue) Then
            strValue = CStr(Trim(strValue))

            Redim Preserve m_footerTemplates(uBound(m_footerTemplates) + 1)
            m_footerTemplates(uBound(m_footerTemplates)) = array("html", strValue)

            footer_addHtml = True
        Else
            footer_addHtml = False
        End If
    End Function


    Public Function footer_addFile(filename)
        If NOT IsEmpty(filename) AND NOT IsNull(filename) Then
            filename = CStr(Trim(filename))

            Redim Preserve m_footerTemplates(uBound(m_footerTemplates) + 1)
            m_footerTemplates(uBound(m_footerTemplates)) = array("file", filename)
            footer_addFile = True
        Else
            footer_addFile = False
        End If
    End Function


    Public Function header_addHtml(strValue)
        If me.bolHeaderSent Then
            header_addHtml = False
        Else
            strValue = Trim(CStr(strValue))
            If LEN(strValue) > 0 Then
                m_headerHTML = m_headerHTML & vbCr & strValue
            End If

            header_addHtml = True
        End If
    End Function


    Public Function inject_css(css)
        css = Trim(CStr(css))
        If Not m_bolDebugMode Then css = Replace(css, " ", "", 1, -1) End If
        m_inlineCSS = m_inlineCSS & vbCr & vbTab & css
    End Function


    Public Function inject_jsValue(name, value)
        name = Trim(CStr(name))
        value = Trim(CStr(value))

        m_jsInjectValues.item(name) = value

        inject_jsValue = True
    End Function


    'Constructor
    Public Function Init
        Dim temp: temp = 0
        Dim Url: Set Url = New UrlParser

        m_bolHeaderSent = False
        m_bolFooterSent = False
        m_description = ""
        m_favIcon = ""
        m_footerHTML = ""
        m_footerTemplates = array()
        Set m_FSO = CreateObject("Scripting.FileSystemObject")
        m_headerHTML = ""
        m_inlineCSS = ""
        m_javascripts = array()
        Set m_jsInjectValues = CreateObject("Scripting.Dictionary")
        m_stylesheets = array()
        Set m_meta = CreateObject("Scripting.Dictionary")
        m_title = ""
        m_titleSuffix = ""

        'Determine m_pathShared
        Url.path = Trim(Request.ServerVariables("PATH_INFO"))

        m_pathShared = Url.pathSeparator & Url.directory(0) & Url.pathSeparator

        If LCase(Url.directory(0)) = "shared" Then
            temp = LCase(Url.directory(1))
            If temp = "dev" OR temp = "staging" Then
                m_pathShared = m_pathShared & Url.directory(1) & Url.pathSeparator
            End If
        Else
            m_pathShared = m_pathShared & "Shared" & Url.pathSeparator
        End If

        'Build default meta tags
        me.set_Meta "dcterms.dateCopyrighted", CStr(Year(Date))
        me.set_Meta "dcterms.rightsHolder", "All rights reserved"
        me.set_meta "robots", "NOINDEX,NOFOLLOW"
    End Function


    Public Function link_raw(resourceFile)
        Dim ext: ext = False
        Dim returnValue: returnValue = False
        Dim temp: temp = ""
        Dim validExtensions: validExtensions = array()
        Dim Url: Set Url = New UrlParser

        Common.append_array validExtensions, "css", True
        Common.append_array validExtensions, "js", True

        resourceFile = Replace(Trim(CStr(resourceFile)), "\", "/")
        Url.path = resourceFile

        temp = LCase(Url.fileExtension)
        If Common.in_array(temp, validExtensions) Then
            returnValue = true
            Select Case temp
                Case "css"
                    Common.append_array m_stylesheets, Url.fullPath, True
                Case "js"
                    Common.append_array m_javascripts, Url.fullPath, True
                Case Else
                    returnValue = False
            End Select
        End If
    End Function


    Public Function link_shared(resourceFile)
        If Left(resourceFile, 1) = "/" Then
            resourceFile = MID(resourceFile, 2)
        End If

        link_raw(m_pathShared & resourceFile)
    End Function


    Public Function set_Meta(name, value)
        name = LCase(Trim(Cstr(name)))
        value = Trim(Cstr(value))

        m_meta.item(name) = value

        set_Meta = true
    End Function


    Public Function output_Header
        Dim element
        Dim index
        Dim temp
        Dim tempKey
        Dim tempKeys

        m_bolHeaderSent = True

        Response.Write("<!DOCTYPE html>" & vbCr)
        Response.Write("<html class=""no-js"" lang=""en"">" & vbCr)
        Response.Write("<head>" & vbCr)

        'Debug Mode Warning Comment
        If Me.DebugMode Then
            Response.Write(vbTab & "<!-- PageFabricator is in Debug Mode -->" & vbCr)
        End If

        'Output Title
        Response.Write(vbTab & "<title>")
        If LEN(me.Title) > 0 Then
             Response.Write(me.Title)
        Else
            Response.Write("Developed by EUD")
        End If
        If LEN(me.TitleSuffix) > 0 Then Response.Write(" " & me.TitleSuffix) End If
        Response.Write("</title>" & vbCr)
        me.inject_jsValue "title", """" & me.Title & """"

        Response.Write(vbTab & "<meta charset=""utf-8"">" & vbCr)
        Response.Write(vbTab & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & vbCr)

        'Output metas
        For Each element in m_meta
            Response.Write(vbTab & "<meta name=""" & element & """ content=""" & m_meta(element) & """>" & vbCr)
        Next

        'Output Style-Sheets
        'm_stylesheets = sort_urlArray(m_stylesheets)
        Response.Write(vbTab & "<!-- StyleSheets -->" & vbCr)
        For index = LBound(m_stylesheets) To UBound(m_stylesheets)
            Response.Write(vbTab & "<link href=""" & m_stylesheets(index) & """ rel=""stylesheet"" type=""text/css"">" & vbCr)
        Next

        'Injected inline CSS
        If LEN(Trim(m_inlineCSS)) > 0 Then
            Response.Write(vbTab & "<!-- Injected Styling -->" & vbCr)
            Response.Write(vbTab & "<style>")
            Response.Write(m_inlineCSS)
            Response.Write(vbCr & vbTab & "</style>" & vbCr)
        End If

        'Output Javascript
        Response.Write(vbTab & "<!-- JavaScripts -->" & vbCr)
        For index = LBound(m_javascripts) To UBound(m_javascripts)
            Response.Write(vbTab & "<script src=""" & m_javascripts(index) & """></script>" & vbCr)
        Next

        inject_jsValue "headerSent", "true"
        'Injected Javascript Variables
        If m_jsInjectValues.Count > 0 Then
            'Inject Javascript Variables
            Response.Write(vbTab & "<!-- Injected JavaScript Values -->" & vbCr)
            Response.Write(vbTab & "<script type=""text/javascript"">" & vbCr)
            Response.Write(vbTab & vbTab & "window.PAGE = {};" & vbCr)
            For Each element in m_jsInjectValues
                Response.Write(vbTab & vbTab & "PAGE." & element & " = " & m_jsInjectValues(element) & ";" & vbCr)
                m_jsInjectValues.Remove(element)
            Next
            Response.Write(vbTab & "</script>" & vbCr)
        End If

        'Close HTML Header
        Response.Write("</head>" & vbCr)

        'Open Body Tag
        If LEN(m_bodyStyle) > 0 Then
            Response.Write("<body style=""" & m_bodyStyle & """>" & vbCr)
        Else
            Response.Write("<body>" & vbCr)
        End If

        'Output header html
        If Me.DebugMode Then
            Response.Write(vbCr & "<!-- HEADER HTML -->" & vbCr)
        End If
        Response.Write(m_headerHTML)
        If Me.DebugMode Then
            Response.Write(vbCr & "<!-- /HEADER HTML -->" & vbCr & vbCr)
        End If

        inject_jsValue "footerSent", "true"
    End Function


    Public Function output_Footer
        m_bolFooterSent = True
        If Me.DebugMode Then
            Response.Write("<!-- Footer Output -->")
        End If

        If Me.DebugMode Then
            Response.Write(vbCr & "<!-- FOOTER TEMPLATES -->" & vbCr)
        End If

        For index = LBound(m_footerTemplates) To UBound(m_footerTemplates)
            templateType = LCase(Trim(m_footerTemplates(index)(0)))
            templateData = m_footerTemplates(index)(1)

            If templateType = "html" Then
                Response.Write(templateData)
            ElseIf templateType = "file" Then
                response.write(templateData)
                Server.Execute(templateData)
            Else
                Response.Write("<!-- RECIEVED AN INVALID TEMPLATE TYPE -->")
            End If

            Response.Write(vbCr)
        Next
        
        If Me.DebugMode Then
            Response.Write(vbCr & "<!-- /FOOTER TEMPLATES -->" & vbCr & vbCr)
        End If

        'Injected Javascript Variables that didn't make it into header
        If m_jsInjectValues.Count > 0 Then
            'Inject Javascript Variables
            Response.Write(vbTab & "<!-- Injected JavaScript Values -->" & vbCr)
            Response.Write(vbTab & "<script type=""text/javascript"">" & vbCr)
            Response.Write(vbTab & vbTab & "window.PAGE = (window.PAGE || {});" & vbCr)
            For Each element in m_jsInjectValues
                Response.Write(vbTab & vbTab & "PAGE." & element & " = " & m_jsInjectValues(element) & ";" & vbCr)
            Next
            Response.Write(vbTab & "</script>" & vbCr)
        End If

        Response.Write("</body>" & vbCr)
        Response.Write("</html>")
        If Me.DebugMode Then
            Response.Write("<!-- End of Footer Output -->")
        End If
        Response.End
    End Function

End Class
%>
