Public Interface ICapturable

   Property CaptureInProgress As Boolean

   Sub Capture(Caspar As CasparCG, Channel As Integer)

   Event Finished As EventHandler

End Interface
