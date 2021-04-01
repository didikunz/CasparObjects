'Copyright by Media Support - Didi Kunz didi@mediasupport.ch
'The same license conditions apply as CasparCG server has.
'
Imports System.Text
Imports System.Xml
Imports System.Drawing
Imports System.IO

''' <summary>
''' Template object
''' </summary>
''' <remarks>Used to send template commands to Caspar. Can also be used for GetTemplates</remarks>
<Serializable()> _
Public Class Template

#Region "Properties"

   Private _Fields As List(Of TemplateField) = New List(Of TemplateField)

   ''' <summary>
   ''' Name of the template
   ''' </summary>
   ''' <remarks>Set by the GetTemplates function</remarks>
   Public Property Name As String = ""

   ''' <summary>
   ''' Author of the template
   ''' </summary>
   ''' <remarks>Set by the GetTemplates function</remarks>
   Public Property Author As String = ""

   ''' <summary>
   ''' AuthorEMail of the template
   ''' </summary>
   ''' <remarks>Set by the GetTemplates function</remarks>
   Public Property AuthorEMail As String = ""

   ''' <summary>
   ''' Info of the template
   ''' </summary>
   ''' <remarks>Set by the GetTemplates function</remarks>
   Public Property Info As String = ""

   ''' <summary>
   ''' Width of the template
   ''' </summary>
   ''' <remarks>Set by the GetTemplates function</remarks>
   Public Property Width As Integer = 0

   ''' <summary>
   ''' Height of the template
   ''' </summary>
   ''' <remarks>Set by the GetTemplates function</remarks>
   Public Property Height As Integer = 0

   ''' <summary>
   ''' Frame rate of the template
   ''' </summary>
   ''' <remarks>Set by the GetTemplates function</remarks>
   Public Property FrameRate As Integer = 0

   ''' <summary>
   ''' Format TemplateDataText as JSON, if false use XML
   ''' </summary>
   Public Property UseJSON As Boolean

   ''' <summary>
   ''' set to True if InfoFields are already added
   ''' </summary>
   Public Property InfoFieldsAdded As Boolean = False

   ''' <summary>
   ''' List of filelds
   ''' </summary>
   ''' <remarks>Used to send template fields to a template. Also set by the GetTemplates function</remarks>
   Public ReadOnly Property Fields As List(Of TemplateField)
      Get
         Return _Fields
      End Get
   End Property

#End Region

#Region "Methods"

   ''' <summary>
   ''' Formats the Fields list as XML
   ''' </summary>
   ''' <returns>A XML formated string</returns>
   ''' <remarks>Uses the XML format, that the template commands need</remarks>
   Public Function TemplateDataText() As String

      '---XML Format:
      '<templateData>
      '   <componentData id=\"f0\">
      '      <data id=\"text\" value=\"Niklas P Andersson\"></data>
      '   </componentData>
      '   <componentData id=\"f1\">
      '      <data id=\"text\" value=\"developer\"></data>
      '   </componentData>
      '   <componentData id=\"f2\">
      '      <data id=\"text\" value=\"Providing an example\"></data>
      '   </componentData>
      '</templateData>

      Dim quote As String = "\" + ChrW(&H22)
      Dim sb As StringBuilder = New StringBuilder

      If UseJSON = False Then
         '----------As XML

         sb.Append("<templateData>")

         For Each tf As TemplateField In _Fields

            sb.AppendFormat("<componentData id={0}{1}{0}>", quote, tf.Name)

            If tf.RenderAsElement = False Then

               'render as attribute
               sb.AppendFormat("<data id={0}{2}{0} value={0}{1}{0}", quote, tf.Value.Replace(ChrW(&H22), "\\&quot;").Replace("'", "\\&apos;"), IIf(tf.AttributeType = TemplateField.enumAttributeType.Text, "text", "image"))
               For Each kvp As KeyValuePair(Of String, String) In tf.Attachements
                  sb.AppendFormat(" {1}={0}{2}{0}", quote, kvp.Key, kvp.Value.Replace(ChrW(&H22), "'"))
               Next
               sb.Append(" />")

            Else

               'render as element
               sb.AppendFormat("<data id={0}text{0}", quote)
               For Each kvp As KeyValuePair(Of String, String) In tf.Attachements
                  sb.AppendFormat(" {1}={0}{2}{0}", quote, kvp.Key, kvp.Value.Replace(ChrW(&H22), "'"))
               Next
               sb.Append(">")
               sb.Append(tf.Value.Replace(ChrW(&H22), "'"))
               sb.Append("</data>")
            End If

            sb.Append("</componentData>")

         Next

         sb.Append("</templateData>")

      Else
         '----------As JSON
         sb.Append("{")

         For Each tf As TemplateField In _Fields
            sb.AppendFormat("{0}{1}{0}: {0}{2}{0}, ", quote, tf.Name, tf.Value.Replace(ChrW(&H22), "'"))
         Next

         sb.Remove(sb.Length - 2, 2)
         sb.Append("}")

      End If


      Return sb.ToString

   End Function

   ''' <summary>
   ''' Add TemplateField to this Template
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Value">Value of the field</param>
   ''' <remarks>Checks for null (Nothing) values</remarks>
   Public Sub AddField(Name As String, Value As String)
      If Value IsNot Nothing Then
         Fields.Add(New TemplateField(Name, Value))
      End If
   End Sub

   ''' <summary>
   ''' Add TemplateField to this Template
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Value">Value of the field</param>
   ''' <param name="Encoding">Enables text encoding for the value</param>
   ''' <remarks>Checks for null (Nothing) values</remarks>
   Public Sub AddField(Name As String, Value As String, Encoding As Boolean)
      If Value IsNot Nothing Then
         Fields.Add(New TemplateField(Name, Value, Encoding))
      End If
   End Sub

   ''' <summary>
   ''' Add TemplateField to this Template
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Value">Value of the field</param>
   ''' <param name="Encoding">Enables text encoding for the value</param>
   ''' <param name="RenderAsElement">Render value as XML element insted of an attribute</param>
   ''' <remarks>Checks for null (Nothing) values</remarks>
   Public Sub AddField(Name As String, Value As String, Encoding As Boolean, RenderAsElement As Boolean)
      Fields.Add(New TemplateField(Name, Value, Encoding, RenderAsElement))
   End Sub

   ''' <summary>
   ''' Add TemplateField to this Template
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Value">Value of the field</param>
   ''' <param name="Encoding">Enables text encoding for the value</param>
   ''' <param name="RenderAsElement">Render value as XML element insted of an attribute</param>
   ''' <param name="AttributeType">The type attribute</param>
   ''' <remarks>Checks for null (Nothing) values</remarks>
   Public Sub AddField(Name As String, Value As String, Encoding As Boolean, RenderAsElement As Boolean, AttributeType As TemplateField.enumAttributeType)
      If Value IsNot Nothing Then
         Fields.Add(New TemplateField(Name, Value, Encoding, RenderAsElement, AttributeType))
      End If
   End Sub

   ''' <summary>
   ''' Add TemplateField to this Template
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Value">Value of the field</param>
   Public Sub AddField(Name As String, Value As Integer)
      Fields.Add(New TemplateField(Name, Value.ToString))
   End Sub

   ''' <summary>
   ''' Add TemplateField to this Template
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Value">Value of the field</param>
   Public Sub AddField(Name As String, Value As Boolean)
      Fields.Add(New TemplateField(Name, Value.ToString))
   End Sub

   ''' <summary>
   ''' Add picture field to this Template as URI
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="PictureFilename">Full path to the picture file</param>
   ''' <remarks>Formats PictureFilename as URI, checks if file exist</remarks>
   Public Sub AddPictureField(Name As String, PictureFilename As String)
      AddPictureField(Name, PictureFilename, False, False)
   End Sub

   ''' <summary>
   ''' Add picture field to this Template as URI or as Base64
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="PictureFilename">Full path to the picture file</param>
   ''' <param name="SendAsBase64">Load the file and convert it to Base64 if true, send the filename URI otherwise</param>
   Public Sub AddPictureField(Name As String, PictureFilename As String, SendAsBase64 As Boolean)
      AddPictureField(Name, PictureFilename, SendAsBase64, False)
   End Sub

   ''' <summary>
   ''' Add picture field to this Template as URI or as Base64
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="PictureFilename">Full path to the picture file</param>
   ''' <param name="SendAsBase64">Load the file and convert it to Base64 if true, send the filename URI otherwise</param>
   ''' <param name="SendAsImageAttribute">Send the picture-filename as image attrivute (XML-only)</param>
   Public Sub AddPictureField(Name As String, PictureFilename As String, SendAsBase64 As Boolean, SendAsImageAttribute As Boolean)
      If PictureFilename IsNot Nothing AndAlso IO.File.Exists(PictureFilename) Then
         If SendAsBase64 = False Then
            If SendAsImageAttribute Then
               Fields.Add(New TemplateField(Name, New System.Uri(PictureFilename).ToString, TemplateField.enumAttributeType.Image))
            Else
               Fields.Add(New TemplateField(Name, New System.Uri(PictureFilename).ToString))
            End If
         Else
            Try
               AddPictureField(Name, New Bitmap(PictureFilename))
            Catch ex As Exception
               'ignore
            End Try
         End If
      End If
   End Sub

   ''' <summary>
   ''' Add picture field to this Template as Base64
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Picture">A System.Drawing.Image, that will be converted as a PNG Base64 encoded string</param>
   ''' <remarks>Encode picture as Base64 string, ignores errors</remarks>
   Public Sub AddPictureField(Name As String, Picture As Image)

      Try
         Dim ms As MemoryStream = New MemoryStream
         Picture.Save(ms, Imaging.ImageFormat.Png)

         Dim imageBytes() As Byte = ms.ToArray

         Dim base64String As String = Convert.ToBase64String(imageBytes)
         Fields.Add(New TemplateField(Name, base64String))

      Catch ex As Exception
         'ignore
      End Try

   End Sub

   ''' <summary>
   ''' Add color field to this Template
   ''' </summary>
   ''' <param name="Name">Name of the field</param>
   ''' <param name="Color">System.Drawing.Color</param>
   ''' <remarks>Formats the color value as hex-string</remarks>
   Public Sub AddColorField(Name As String, Color As Color)
      Fields.Add(New TemplateField(Name, String.Format("0x{0:X8}", Color.ToArgb)))
   End Sub

   ''' <summary>
   ''' Clear the list of fields
   ''' </summary>
   Public Sub Clear()
      Fields.Clear()
   End Sub

   ''' <summary>
   ''' Formated ToString function
   ''' </summary>
   ''' <returns>A String with the Template's name, dimensions and author</returns>
   Public Overrides Function ToString() As String
      If Me.Width > 0 Then
         Return String.Format("{0}, {1}x{2}, {3}", Me.Name, Me.Width, Me.Height, Me.Author)
      Else
         Return TemplateDataText()
      End If
   End Function

#End Region

#Region "Shared Methods"

   ''' <summary>
   ''' Create a Template object by parsing a XML string
   ''' </summary>
   ''' <param name="XmlText">The XML string to parse</param>
   ''' <returns>A Template object parsed</returns>
   ''' <remarks>Use the XML format, that the template commands need</remarks>
   Public Shared Function Parse(XmlText As String) As Template

      Dim ti As Template = New Template
      Dim doc As XmlDocument = New XmlDocument()

      doc.LoadXml(XmlText.TrimEnd(ChrW(&H0)))

      If doc.FirstChild.Name = "templateData" Then

         '<templateData>
         ' <componentData id = "f0" >
         '      <data id="text" value="Test f0"/>
         ' </componentData>
         ' <componentData id = "f1" >
         '      <data id="text" value="Test f1"/>
         ' </componentData>
         ' <componentData id = "f2" >
         '      <data id="text" value="Test f2"/>
         ' </componentData>
         ' <componentData id = "f3" >
         '      <data id="text" value="Test f3"/>
         ' </componentData>
         '</templateData>

         For Each fld As XmlNode In doc.FirstChild.ChildNodes
            ti.Fields.Add(New TemplateField(fld.Attributes("id").Value, fld.FirstChild.Attributes(1).Value))
         Next

      Else

         '<?xml version="1.0" encoding="utf-8"?>
         '<template version="1.8.0" authorName="Didi Kunz" authorEmail="didi@mediasupport.ch" templateInfo="" originalWidth="1024" originalHeight="576" originalFrameRate="50">
         '   <components>
         '      <component name="CasparTextField">
         '         <property name="text" type="string" info="String data"/>
         '      </component>
         '   </components>
         '   <keyframes/>
         '   <instances>
         '      <instance name="f0" type="CasparTextField"/>
         '   </instances>
         '</template>

         Dim nd As XmlNode = doc.SelectSingleNode("template")
         If nd IsNot Nothing AndAlso nd.Attributes IsNot Nothing AndAlso nd.Attributes.Count > 1 Then
            ti.Author = nd.Attributes("authorName").Value
            ti.AuthorEMail = nd.Attributes("authorEmail").Value
            ti.Info = nd.Attributes("templateInfo").Value
            ti.Width = Integer.Parse(nd.Attributes("originalWidth").Value)
            ti.Height = Integer.Parse(nd.Attributes("originalHeight").Value)
            ti.FrameRate = Integer.Parse(nd.Attributes("originalFrameRate").Value)
         End If

         nd = doc.SelectSingleNode("template/instances")
         If nd IsNot Nothing Then
            For Each fld As XmlNode In nd.ChildNodes
               ti.Fields.Add(New TemplateField(fld.Attributes("name").Value))
            Next
         End If

      End If

      Return ti

   End Function

#End Region

End Class

