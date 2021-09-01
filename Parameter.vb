<Serializable>
Public Class Parameter

   Public Enum enumFieldType
      ftString
      ftInteger
      ftReal
      ftImage
      ftClip
      ftList
      ftColor
      ftFilename
      ftCSV
   End Enum

   Public Property Name As String = ""
   Public Property Type As enumFieldType = enumFieldType.ftString
   Public Property Info As String = ""

   Public Function GetFieldType() As String
      Select Case Me.Type
         Case enumFieldType.ftInteger
            Return "integer"
         Case enumFieldType.ftReal
            Return "real"
         Case enumFieldType.ftImage
            Return "image"
         Case enumFieldType.ftClip
            Return "clip"
         Case enumFieldType.ftList
            Return "list"
         Case enumFieldType.ftColor
            Return "color"
         Case enumFieldType.ftFilename
            Return "filename"
         Case enumFieldType.ftCSV
            Return "csv"
         Case Else
            Return "string"
      End Select
   End Function

   Public Sub New()
   End Sub

   Public Sub New(name As String, type As enumFieldType)
      Me.Name = name
      Me.Type = type
      Select Case type
         Case enumFieldType.ftImage
            Info = "Image URL"
         Case enumFieldType.ftColor
            Info = "Color as hex string"
         Case Else
            Info = ""
      End Select
   End Sub

   Public Sub New(name As String, type As enumFieldType, info As String)
      Me.Name = name
      Me.Type = type
      Select Case type
         Case enumFieldType.ftImage
            Me.Info = info + " Image URL"
         Case enumFieldType.ftColor
            Me.Info = info + " Color as hex string"
         Case Else
            Me.Info = info
      End Select
   End Sub


End Class
