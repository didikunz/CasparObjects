Imports System.Collections
Imports System.Windows.Forms

Public Class DelayQueue

   Private Class DelayedCommand
      Public Property Delay As CasparCG.Retard
      Public Property Command As String

      Public Sub New(delay As Integer, command As String)
         Me.Delay.Amount = delay
         Me.Command = command
      End Sub
   End Class


   Private _Caspar As CasparCG = Nothing
   Private _queue As List(Of DelayedCommand) = New List(Of DelayedCommand)
   Private _currentDelay As Integer = 0
   Private _nextCommand As String = ""

   Private WithEvents _timer As Timer = New Timer()

   Public WriteOnly Property Caspar As CasparCG
      Set(value As CasparCG)
         _Caspar = value
      End Set
   End Property

   Public Sub Add(delay As Integer, command As String)

      If _currentDelay > 0 Then
         delay -= _currentDelay
         If delay < 0 Then
            delay = 0
         End If
      End If

      If _queue.Count = 0 Then
         _queue.Add(New DelayedCommand(delay, command))
      ElseIf _queue.Count = 1 Then

         If _queue(0).Delay.Amount >= delay Then
            _queue.Insert(0, New DelayedCommand(delay, command))
         Else
            _queue.Add(New DelayedCommand(delay, command))
         End If

      Else

         If _queue(0).Delay.Amount >= delay Then
            _queue.Insert(0, New DelayedCommand(delay, command))
         Else

            Dim i As Integer = 0
            Do
               Dim dc1 As DelayedCommand = _queue(i)
               Dim dc2 As DelayedCommand = Nothing
               If i < _queue.Count - 1 Then
                  dc2 = _queue(i + 1)
               End If

               If delay >= dc1.Delay.Amount Then

                  If dc2 IsNot Nothing Then

                     If delay < dc2.Delay.Amount Then
                        _queue.Insert(i + 1, New DelayedCommand(delay, command))
                        Exit Do
                     End If

                  Else
                     _queue.Add(New DelayedCommand(delay, command))
                     Exit Do
                  End If

               Else

                  If dc2 Is Nothing Then
                     _queue.Add(New DelayedCommand(delay, command))
                     Exit Do
                  End If

               End If

               i += 1
               If i > _queue.Count - 1 Then
                  Exit Do
               End If
            Loop

         End If
      End If

      StartTimer()

   End Sub

   Private Sub StartTimer()

      If _queue.Count > 0 Then

         If Not _timer.Enabled Then

            Dim top As DelayedCommand = _queue(0)
            _currentDelay = top.Delay.Amount
            _nextCommand = top.Command
            _queue.RemoveAt(0)

            For Each dc As DelayedCommand In _queue
               dc.Delay.Amount -= _currentDelay
            Next

            If _currentDelay = 0 Then
               _Caspar.Execute(_nextCommand)
               StartTimer()
            Else
               _timer.Interval = _currentDelay
               _timer.Start()
            End If

         End If

      End If

   End Sub

   Private Sub _timer_Tick(sender As Object, e As EventArgs) Handles _timer.Tick

      _timer.Stop()
      _currentDelay = 0
      _Caspar.Execute(_nextCommand)
      StartTimer()

   End Sub

   Public Sub New()
   End Sub

   Public Sub New(Caspar As CasparCG)
      _Caspar = Caspar
   End Sub

End Class
