'Copyright by Media Support - Didi Kunz didi@mediasupport.ch
'The same license conditions apply as CasparCG server has.
'
''' <summary>
''' ReturnInfo object
''' </summary>
''' <remarks>Gets Caspars server response</remarks>
<Serializable()> _
Public Class ReturnInfo

   ''' <summary>
   ''' Numeric return value 
   ''' </summary>
   Public Property Number As Integer

   ''' <summary>
   ''' The message returned
   ''' </summary>
   Public Property Message As String

   ''' <summary>
   ''' Data returned
   ''' </summary>
   ''' <remarks>Some commands return data</remarks>
   Public Property Data As String

   ''' <summary>
   ''' Parameter less constructor
   ''' </summary>
   Public Sub New()
   End Sub

   ''' <summary>
   ''' Parameterized constructor
   ''' </summary>
   ''' <param name="Number">Numeric return value</param>
   ''' <param name="Message">The message returned</param>
   ''' <param name="Data">Some commands return data</param>
   ''' <remarks></remarks>
   Public Sub New(Number As Integer, Message As String, Data As String)
      _Number = Number
      _Message = Message
      _Data = Data
   End Sub

End Class

