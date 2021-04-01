Public Class TemplateParameter

   Public Property id As String
   Public Property type As String
   Public Property info As String

   Public Sub New()
      id = ""
      type = ""
      info = ""
   End Sub

   Public Sub New(theID As String, theType As String, theInfo As String)
      id = theID
      type = theType
      info = theInfo
   End Sub

End Class
