'Copyright by Media Support - Didi Kunz didi@mediasupport.ch
'The same license conditions apply as CasparCG server has.
'
Imports System.Text

''' <summary>
''' TemplateField object
''' </summary>
''' <remarks>Used to send fields to templates or retrieve them</remarks>
<Serializable()> _
Public Class TemplateField
   Implements IComparable(Of TemplateField)

   Public Enum enumAttributeType
      Text
      Image
   End Enum

#Region "Private vars"

   Private _Attachements As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))

#End Region

#Region "Properties"

   ''' <summary>
   ''' Field name
   ''' </summary>
   Public Property Name As String = ""

   ''' <summary>
   ''' Field value
   ''' </summary>
   Public Property Value As String = ""

   ''' <summary>
   ''' Attachements are additional non standard data that are added to the output as separate attributes. Currently only works in XML.
   ''' </summary>
   ''' <returns>A list of KeyValuePairs</returns>
   Public ReadOnly Property Attachements As List(Of KeyValuePair(Of String, String))
      Get
         Return _Attachements
      End Get
   End Property

   ''' <summary>
   ''' Render the value as an XML Element instead of a Value attribute. Only works in XML.
   ''' </summary>
   Public Property RenderAsElement As Boolean = False

   ''' <summary>
   ''' Set the type attribute for this value. Only works in XML.
   ''' </summary>
   ''' <returns></returns>
   Public Property AttributeType As enumAttributeType = enumAttributeType.Text

#End Region

#Region "Shared"

   Public Shared Function EncodeText(text As String, encodeAsHTML As Boolean) As String

      If encodeAsHTML = False Then

         Return text

      Else 'HTML

         Dim stringBuilder As New StringBuilder
         Dim i As Integer
         Dim c As Char
         For i = 0 To text.Length - 1
            c = text(i)
            Select Case Asc(c)
               Case 34  '"
                  stringBuilder.Append("&#34;")
               Case 38  '&
                  stringBuilder.Append("&#38;")
               Case 39  ''
                  stringBuilder.Append("&#39;")
               Case 60  '<
                  stringBuilder.Append("&lt;")
               Case 62  '>
                  stringBuilder.Append("&gt;")
               Case 92  '\
                  stringBuilder.Append("&#92;")
               Case Else
                  stringBuilder.Append(c)
            End Select
         Next
         Return stringBuilder.ToString

      End If

   End Function

#End Region

#Region "Methods"

   ''' <summary>
   ''' Formated ToString function
   ''' </summary>
   ''' <returns>A string containing name and value</returns>
   Public Overrides Function ToString() As String
      Return String.Format("{0} = '{1}'", Me.Name, Me.Value)
   End Function

   ''' <summary>
   ''' CompareTo for sorting
   ''' </summary>
   ''' <param name="other">A TemplateField object to compare this with</param>
   ''' <returns>An integer indicating if bigger, smaller or equal</returns>
   Public Function CompareTo(other As TemplateField) As Integer Implements System.IComparable(Of TemplateField).CompareTo
      Return Name.CompareTo(other.Name)
   End Function

   ''' <summary>
   ''' Add a new Attachement to the list. 
   ''' </summary>
   ''' <param name="Name">Name of the attachement => name of the attribute.</param>
   ''' <param name="Value">Value of the attachement => value of the attribute.</param>
   Public Sub AddAttachement(Name As String, Value As String)
      AddAttachement(Name, Value, False)
   End Sub

   ''' <summary>
   ''' Add a new Attachement to the list. 
   ''' </summary>
   ''' <param name="Name">Name of the attachement => name of the attribute.</param>
   ''' <param name="Value">Value of the attachement => value of the attribute.</param>
   ''' <param name="EncodeAsHTML">Text encoding for the value</param>
   Public Sub AddAttachement(Name As String, Value As String, EncodeAsHTML As Boolean)
      _Attachements.Add(New KeyValuePair(Of String, String)(Name, EncodeText(Value, EncodeAsHTML)))
   End Sub

#End Region

#Region "Constructors"

   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name="Name">The name of the field</param>
   Public Sub New(Name As String)
      Me.Name = Name
   End Sub

   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name="Name">The name of the field</param>
   ''' <param name="Value">The value of the field</param>
   Public Sub New(Name As String, Value As String)
      Me.Name = Name
      Me.Value = Value
   End Sub

   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name="Name">The name of the field</param>
   ''' <param name="Value">The value of the field</param>
   ''' <param name="Type">The type attribute</param>
   Public Sub New(Name As String, Value As String, Type As enumAttributeType)
      Me.Name = Name
      Me.Value = Value
      Me.AttributeType = Type
   End Sub

   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name="Name">The name of the field</param>
   ''' <param name="Value">The value of the field</param>
   ''' <param name="EncodeAsHTML">Enables text encoding for the value</param>
   Public Sub New(Name As String, Value As String, EncodeAsHTML As Boolean)
      Me.Name = Name
      Me.Value = EncodeText(Value, EncodeAsHTML)
   End Sub

   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name="Name">The name of the field</param>
   ''' <param name="Value">The value of the field</param>
   ''' <param name="EncodeAsHTML">Enables text encoding for the value</param>
   ''' <param name="Type">The type attribute</param>
   Public Sub New(Name As String, Value As String, EncodeAsHTML As Boolean, Type As enumAttributeType)
      Me.Name = Name
      Me.Value = EncodeText(Value, EncodeAsHTML)
      Me.AttributeType = Type
   End Sub


   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name="Name">The name of the field</param>
   ''' <param name="Value">The value of the field</param>
   ''' <param name="EncodeAsHTML">Enables text encoding for the value</param>
   ''' <param name="RenderAsElement">Render value as XML element insted of an attribute</param>
   Public Sub New(Name As String, Value As String, EncodeAsHTML As Boolean, RenderAsElement As Boolean)
      Me.Name = Name
      Me.Value = EncodeText(Value, EncodeAsHTML)
      Me.RenderAsElement = RenderAsElement
   End Sub

   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name="Name">The name of the field</param>
   ''' <param name="Value">The value of the field</param>
   ''' <param name="EncodeAsHTML">Enables text encoding for the value</param>
   ''' <param name="RenderAsElement">Render value as XML element insted of an attribute</param>
   ''' <param name="Type">The type attribute</param>
   Public Sub New(Name As String, Value As String, EncodeAsHTML As Boolean, RenderAsElement As Boolean, Type As enumAttributeType)
      Me.Name = Name
      Me.Value = EncodeText(Value, EncodeAsHTML)
      Me.RenderAsElement = RenderAsElement
      Me.AttributeType = Type
   End Sub
#End Region

End Class

