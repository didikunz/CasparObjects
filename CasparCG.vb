'Copyright by Media Support - Didi Kunz didi@mediasupport.ch
'The same license conditions apply as CasparCG server has.
'
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Diagnostics
Imports System.ComponentModel
Imports System.Threading
Imports System.Globalization

''' <summary>
''' CasparCG connection object
''' </summary>
''' <remarks>Wrapper around a System.Net.Sockets.Socket object</remarks>
<Serializable()> _
Public Class CasparCG
   Implements ICloneable

#Region "Enums and local classes"

   ''' <summary>
   ''' Media clip types
   ''' </summary>
   ''' <remarks>Used for GetMediaClipsNames function</remarks>
   Public Enum MediaTypes
      ''' <summary>
      ''' All media types
      ''' </summary>
      All
      ''' <summary>
      ''' Only videos
      ''' </summary>
      Movie
      ''' <summary>
      ''' Only still pictures
      ''' </summary>
      Still
      ''' <summary>
      ''' Only audio clips
      ''' </summary>
      Audio
   End Enum

   ''' <summary>
   ''' Return Value of the Connect function
   ''' </summary>
   Public Enum enumConnectResult
      ''' <summary>
      ''' Succesfully connected to CasparCG
      ''' </summary>
      crSuccessfull
      ''' <summary>
      ''' The machine could not be pinged, network error or machine is not running
      ''' </summary>
      crMachineNotAvailable
      ''' <summary>
      ''' CasparCG does not respond, it is not started on the machine
      ''' </summary>
      crCasparCGNotStarted
      ''' <summary>
      ''' A local CasprCG server could not be loaded using the CasparExePath property
      ''' </summary>
      crLocalCasparCGCouldNotBeStarted
      ''' <summary>
      ''' This is not a local CasparCG and could not be started using the CasparExePath argument
      ''' </summary>
      crIsNotLocal
   End Enum

   ''' <summary>
   ''' The way channel and layer informations get added to template data.
   ''' </summary>
   Public Enum enumAddInfoFieldsType
      ''' <summary>
      ''' No channel and layer informations were added.
      ''' </summary>
      itNone
      ''' <summary>
      ''' The standard "channel" and "layer" variable names are used.
      ''' </summary>
      itStandard
      ''' <summary>
      ''' The Aveco automation specific variable names "astra_output" and "astra_layer" are used.
      ''' </summary>
      itAveco
   End Enum

   ''' <summary>
   ''' Class for use with delayed execution.
   ''' </summary>
   Public Class Retard
      Public Property Amount As Integer
      Public Sub New()
         Me.Amount = 0
      End Sub

      Public Sub New(ByVal Retard As Integer)
         Me.Amount = Retard
      End Sub
   End Class

   Public Class ChannelInfo

      Public Property Width As Integer
      Public Property Height As Integer
      Public Property Framerate As Single
      Public Property IsInterlaced As Boolean

      Public Sub New()
         Me.Width = 1920
         Me.Height = 1080
         Me.Framerate = 50
         Me.IsInterlaced = True
      End Sub

      Public Sub New(text As String)
         text = text.Replace("dci", "")
         Dim i As Integer = text.IndexOfAny("ip".ToCharArray)
         If i = -1 Then
            Select Case text
               Case "PAL"
                  Me.Height = 576
                  Me.Framerate = 50
                  Me.IsInterlaced = True
               Case "NTSC"
                  Me.Height = 486
                  Me.Framerate = 59.94
                  Me.IsInterlaced = True
            End Select
         Else
            Me.Height = Integer.Parse(text.Substring(0, i))
            Me.Framerate = Single.Parse(text.Substring(i + 1)) / 100
            Me.IsInterlaced = (text.Substring(i, 1) = "i")
         End If
         Me.Width = CInt(Me.Height * 1.7778)
      End Sub

      Public Sub New(ByVal Width As Integer, ByVal Height As Integer, ByVal Framerate As Integer, ByVal IsInterlaced As Boolean)
         Me.Width = Width
         Me.Height = Height
         Me.Framerate = Framerate
         Me.IsInterlaced = IsInterlaced
      End Sub

   End Class

#End Region

#Region "Vars"

   <NonSerialized()> Private _Caspar As Socket
   <NonSerialized()> Private _CasparProc As Process
   <NonSerialized()> Private _Scanner As Process
   <NonSerialized()> Private _Version As VersionInfo = New VersionInfo
   <NonSerialized()> Private _ServerPaths As Paths
   <NonSerialized()> Private _ChannelInfos As List(Of ChannelInfo) = Nothing
   <NonSerialized()> Private _DelayQueue As DelayQueue = New DelayQueue(Me)

   Private _ServerAdress As String = "localhost"
   Private _IsLocal As Boolean = True

#End Region

#Region "Properties"

   ''' <summary>
   ''' The name of the CasparCG Server instance
   ''' </summary>
   Public Property Name As String = "Local"

   ''' <summary>
   ''' The port number to connect to
   ''' </summary>
   ''' <remarks>Defaults to 5250</remarks>
   Public Property Port As Integer = 5250

   ''' <summary>
   ''' The default channel to output in
   ''' </summary>
   ''' <remarks>Defaults to 1</remarks>
   Public Property DefaultChannel As Integer = 1

   ''' <summary>
   ''' The default layer to output on
   ''' </summary>
   ''' <remarks>Defaults to 5</remarks>
   Public Property DefaultLayer As Integer = 5

   ''' <summary>
   ''' Inhibits exeptions on connection errors
   ''' </summary>
   ''' <remarks>Used in combination with CasparExePath</remarks>
   Public Property KeepQuiet As Boolean

   ''' <summary>
   ''' Full qualified path to Caspar's exe file. Must be local to the client.
   ''' </summary>
   ''' <remarks>Used in combination with Retries</remarks>
   Public Property CasparExePath As String = ""

   ''' <summary>
   ''' Number of retries for a connection, before giving up.
   ''' </summary>
   ''' <remarks>Used in combination with CasparExePath</remarks>
   Public Property Retries As Integer = 1

   ''' <summary>
   ''' IP-Address or computer name to connect to Caspar
   ''' </summary>
   Public Property ServerAdress As String
      Get
         Return _ServerAdress
      End Get
      Set(ByVal value As String)
         _ServerAdress = value.Trim
         _IsLocal = (_ServerAdress.ToLower = "localhost" Or _ServerAdress = "127.0.0.1")
      End Set
   End Property

   Public Property AddInfoFields As enumAddInfoFieldsType = enumAddInfoFieldsType.itStandard
   Public Property OverwriteInfoFields As Boolean = False

   ''' <summary>
   ''' True to replace backslashes, single and double quotes for HTML.
   ''' </summary>
   Public Property FormatTextsForHTML As Boolean = False

   ''' <summary>
   ''' If Caspar is on the local machine
   ''' </summary>
   ''' <returns>True if local, false if remote</returns>
   ReadOnly Property IsLocal As Boolean
      Get
         Return _IsLocal
      End Get
   End Property

   ''' <summary>
   ''' Points to a folder were pictres are stored
   ''' </summary>
   ''' <remarks>The path is local to the server</remarks>
   Public Property LocalPictureFolder As String

   ''' <summary>
   ''' Points to a folder were pictres are stored
   ''' </summary>
   ''' <remarks>The path is a network share seen from the client</remarks>
   Public Property RemotePictureFolder As String

   ''' <summary>
   ''' Indicate connection status
   ''' </summary>
   ''' <remarks>True if connected, false otherwise</remarks>
   Public ReadOnly Property Connected As Boolean
      Get
         Return _Caspar.Connected
      End Get
   End Property

   Public Class VersionInfo
      Public Property Major As Integer
      Public Property Minor As Integer
      Public Property Revision As Integer
   End Class

   ''' <summary>
   ''' Query the version of CasparCG server as a VersionInfo object
   ''' </summary>
   ''' <returns></returns>
   Public ReadOnly Property Version As VersionInfo
      Get
         If _Version Is Nothing OrElse _Version.Major = 0 Then

            Dim s As String = Execute("VERSION SERVER").Data
            If s <> "" Then

               Dim parts() As String = s.Split(CChar("."))

               If parts.Length >= 3 Then

                  _Version.Major = CInt(parts(0))
                  _Version.Minor = CInt(parts(1))

                  Dim inte As Integer = 0
                  If Integer.TryParse(parts(2), inte) Then
                     _Version.Revision = inte
                  Else
                     Dim subParts() As String = parts(2).Split(CChar(" "))
                     If Integer.TryParse(subParts(0), inte) Then
                        _Version.Revision = inte
                     Else
                        _Version.Revision = 0
                     End If
                  End If

               End If

            End If

         End If

         Return _Version

      End Get
   End Property

   ''' <summary>
   ''' Query the version of CasparCG server as a serialized integer:
   ''' Version.Major * 1000000 +
   ''' Version.Minor * 10000 +
   ''' Version.Revision
   ''' </summary>
   Public ReadOnly Property VersionSerialized As Integer
      Get
         Return (Version.Major * 1000000) + (Version.Minor * 10000) + Version.Revision
      End Get
   End Property

   Public ReadOnly Property IsPsdSupported As Boolean
      Get
         Return (Version.Major >= 2 And Version.Minor >= 1 And Version.Minor <> 2)
      End Get
   End Property

   ''' <summary>
   ''' Paths object
   ''' </summary>
   ''' <remarks>Used for the ServerPaths property</remarks>
   <XmlRoot("paths")>
   Public Class Paths

      ''' <summary>
      ''' Path to the media files (video, audio and stills)
      ''' </summary>
      <XmlElement(ElementName:="media-path")>
      Public Property MediaPath As String = "media\"

      ''' <summary>
      ''' Path to the log files
      ''' </summary>
      <XmlElement(ElementName:="log-path")>
      Public Property LogPath As String = "log\"

      ''' <summary>
      ''' Path to the data files
      ''' </summary>
      ''' <remarks>Data files are managed using the DATA commands</remarks>
      <XmlElement(ElementName:="data-path")>
      Public Property DataPath As String = "data\"

      ''' <summary>
      ''' Path to the templates
      ''' </summary>
      <XmlElement(ElementName:="template-path")>
      Public Property TemplatePath As String = "templates\"

      ''' <summary>
      ''' Path to the thumbnails
      ''' </summary>
      <XmlElement(ElementName:="thumbnail-path")>
      Public Property ThumbnailPath As String = ""

      ''' <summary>
      ''' Path to the fonts
      ''' </summary>
      <XmlElement(ElementName:="font-path")>
      Public Property FontPath As String = ""

      ''' <summary>
      ''' Initial path
      ''' </summary>
      ''' <value></value>
      <XmlElement(ElementName:="initial-path")>
      Public Property InitialPath As String = ""

   End Class

   ''' <summary>
   ''' Gets the paths settings of Caspar
   ''' </summary>
   ''' <remarks>Does not make sense for remote Caspar Servers</remarks>
   Public ReadOnly Property ServerPaths() As Paths
      Get

         If _ServerPaths Is Nothing OrElse _ServerPaths.InitialPath = "" Then

            Try
               Dim s As String = Execute("INFO PATHS").Data
               Dim b As Byte() = Encoding.UTF8.GetBytes(Left(s, InStrRev(s, "</paths>") + 7))
               Dim ser As XmlSerializer = New XmlSerializer(GetType(Paths))
               _ServerPaths = CType(ser.Deserialize(New System.IO.MemoryStream(b)), Paths)
            Catch ex As Exception
               _ServerPaths = New Paths
            End Try

         End If

         Return _ServerPaths

      End Get
   End Property


   Public ReadOnly Property NumberOfChannels As Integer
      Get

         If _ChannelInfos Is Nothing Then
            GetChannelInfos()
         End If

         Return _ChannelInfos.Count

      End Get
   End Property

   ''' <summary>
   ''' Sets the first layer, that will be cleared by Clearlayers
   ''' </summary>
   <XmlIgnore()>
   Public Property FirstClearlayer As Integer

   ''' <summary>
   ''' Sets the last layer, that will be cleared by Clearlayers
   ''' </summary>
   <XmlIgnore()>
   Public Property LastClearlayer As Integer

#End Region

#Region "Template Commands"

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Uses the template name inside the Template object. Defaults to DefaultChannel DefaultLayer and plays the template after loading</remarks>
   Public Function CG_Add(ByVal TemplateData As Template) As ReturnInfo
      Return CG_Add(DefaultChannel, DefaultLayer, TemplateData.Name, TemplateData, True)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Uses the template name inside the Template object. Defaults to DefaultChannel DefaultLayer and plays the template after loading</remarks>
   Public Function CG_Add(ByVal TemplateData As Template, ByVal Retard As Retard) As ReturnInfo
      Return CG_Add(DefaultChannel, DefaultLayer, TemplateData.Name, TemplateData, True, Retard)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer and plays the template after loading</remarks>
   Public Function CG_Add(ByVal TemplateName As String, ByVal TemplateData As Template) As ReturnInfo
      Return CG_Add(DefaultChannel, DefaultLayer, TemplateName, TemplateData, True)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer and plays the template after loading</remarks>
   Public Function CG_Add(ByVal TemplateName As String, ByVal TemplateData As Template, ByVal Retard As Retard) As ReturnInfo
      Return CG_Add(DefaultChannel, DefaultLayer, TemplateName, TemplateData, True, Retard)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel and plays the template after loading</remarks>
   Public Function CG_Add(ByVal Layer As Integer, ByVal TemplateName As String, ByVal TemplateData As Template) As ReturnInfo
      Return CG_Add(DefaultChannel, Layer, TemplateName, TemplateData, True)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel and plays the template after loading</remarks>
   Public Function CG_Add(ByVal Layer As Integer, ByVal TemplateName As String, ByVal TemplateData As Template, ByVal Retard As Retard) As ReturnInfo
      Return CG_Add(DefaultChannel, Layer, TemplateName, TemplateData, True, Retard)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="AutoPlay">Plays the template after loading if set to true</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Add(ByVal Layer As Integer, ByVal TemplateName As String, ByVal TemplateData As Template, ByVal AutoPlay As Boolean) As ReturnInfo
      Return CG_Add(DefaultChannel, Layer, TemplateName, TemplateData, AutoPlay)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="AutoPlay">Plays the template after loading if set to true</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Add(ByVal Layer As Integer, ByVal TemplateName As String, ByVal TemplateData As Template, ByVal AutoPlay As Boolean, ByVal Retard As Retard) As ReturnInfo
      Return CG_Add(DefaultChannel, Layer, TemplateName, TemplateData, AutoPlay, Retard)
   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="AutoPlay">Plays the template after loading if set to true</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Add(ByVal Channel As Integer, ByVal Layer As Integer, ByVal TemplateName As String, ByVal TemplateData As Template, ByVal AutoPlay As Boolean, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo

      If TemplateData IsNot Nothing Then

         If Not TemplateData.InfoFieldsAdded Or OverwriteInfoFields Then
            Select Case Me.AddInfoFields
               Case enumAddInfoFieldsType.itStandard
                  TemplateData.AddField("channel", Channel.ToString)
                  TemplateData.AddField("layer", Layer.ToString)
               Case enumAddInfoFieldsType.itAveco
                  TemplateData.AddField("astra_output", Channel.ToString)
                  TemplateData.AddField("astra_layer", Layer.ToString)
            End Select
            TemplateData.InfoFieldsAdded = True
         End If

         ''ToDo: Comment out for production
         'Data_Store("CasparData", TemplateData)

         Return Execute(String.Format("CG {1}-{2} ADD {6} {0}{4}{0} {5} {0}{3}{0}", ChrW(&H22), Channel, Layer, TemplateData.TemplateDataText(FormatTextsForHTML), TemplateName, IIf(AutoPlay, "1", "0"), FlashLayer))

      Else
         Return New ReturnInfo()
      End If

   End Function

   ''' <summary>
   ''' Adds the template in Caspar
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="AutoPlay">Plays the template after loading if set to true</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Add Function. Adds a channel and a layer TemplateField to the Fields collection. This is usefull to call back the server via TCP/IP.</remarks>
   Public Function CG_Add(ByVal Channel As Integer, ByVal Layer As Integer, ByVal TemplateName As String, ByVal TemplateData As Template, ByVal AutoPlay As Boolean, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo

      If TemplateData IsNot Nothing Then

         If Not TemplateData.InfoFieldsAdded Or OverwriteInfoFields Then
            Select Case Me.AddInfoFields
               Case enumAddInfoFieldsType.itStandard
                  TemplateData.AddField("channel", Channel.ToString)
                  TemplateData.AddField("layer", Layer.ToString)
               Case enumAddInfoFieldsType.itAveco
                  TemplateData.AddField("astra_output", Channel.ToString)
                  TemplateData.AddField("astra_layer", Layer.ToString)
            End Select
            TemplateData.InfoFieldsAdded = True
         End If

         ''ToDo: Comment out for production
         'Data_Store("CasparData", TemplateData)

         Return Execute(String.Format("CG {1}-{2} ADD {6} {0}{4}{0} {5} {0}{3}{0}", ChrW(&H22), Channel, Layer, TemplateData.TemplateDataText(FormatTextsForHTML), TemplateName, IIf(AutoPlay, "1", "0"), FlashLayer), Retard)

      Else
         Return New ReturnInfo()
      End If

   End Function


   ''' <summary>
   ''' Adds the template in Caspar, using a Dataset
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="DataSetName">The name of the DataSet already stored on the server</param>
   ''' <param name="AutoPlay">Plays the template after loading if set to true</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The CG_Add Function witch uses a DataSet insted on a Template object.</remarks>
   Public Function CG_Add(ByVal Channel As Integer, ByVal Layer As Integer, ByVal TemplateName As String, ByVal DataSetName As String, ByVal AutoPlay As Boolean, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} ADD {6} {0}{4}{0} {5} {0}{3}{0}", ChrW(&H22), Channel, Layer, DataSetName, TemplateName, IIf(AutoPlay, "1", "0"), FlashLayer))
   End Function

   ''' <summary>
   ''' Adds the template in Caspar, using a Dataset
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateName">The name of the template</param>
   ''' <param name="DataSetName">The name of the DataSet already stored on the server</param>
   ''' <param name="AutoPlay">Plays the template after loading if set to true</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The CG_Add Function witch uses a DataSet insted on a Template object.</remarks>
   Public Function CG_Add(ByVal Channel As Integer, ByVal Layer As Integer, ByVal TemplateName As String, ByVal DataSetName As String, ByVal AutoPlay As Boolean, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} ADD {6} {0}{4}{0} {5} {0}{3}{0}", ChrW(&H22), Channel, Layer, DataSetName, TemplateName, IIf(AutoPlay, "1", "0"), FlashLayer), Retard)
   End Function


   ''' <summary>
   ''' Plays the template in Caspar
   ''' </summary>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Play() As ReturnInfo
      Return CG_Play(DefaultChannel, DefaultLayer)
   End Function

   ''' <summary>
   ''' Plays the template in Caspar
   ''' </summary>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Play(ByVal Retard As Retard) As ReturnInfo
      Return CG_Play(DefaultChannel, DefaultLayer, Retard)
   End Function

   ''' <summary>
   ''' Plays the template in Caspar
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Play(ByVal Layer As Integer) As ReturnInfo
      Return CG_Play(DefaultChannel, Layer)
   End Function

   ''' <summary>
   ''' Plays the template in Caspar
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Play(ByVal Layer As Integer, ByVal Retard As Retard) As ReturnInfo
      Return CG_Play(DefaultChannel, Layer, Retard)
   End Function

   ''' <summary>
   ''' Plays the template in Caspar
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Play function.</remarks>
   Public Function CG_Play(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {0}-{1} PLAY {2}", Channel, Layer, FlashLayer))
   End Function

   ''' <summary>
   ''' Plays the template in Caspar
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Play function.</remarks>
   Public Function CG_Play(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {0}-{1} PLAY {2}", Channel, Layer, FlashLayer), Retard)
   End Function


   ''' <summary>
   ''' Updates the datafields of a template
   ''' </summary>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Update(ByVal TemplateData As Template) As ReturnInfo
      Return CG_Update(DefaultChannel, DefaultLayer, TemplateData)
   End Function

   ''' <summary>
   ''' Updates the datafields of a template
   ''' </summary>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Update(ByVal TemplateData As Template, ByVal Retard As Retard) As ReturnInfo
      Return CG_Update(DefaultChannel, DefaultLayer, TemplateData, Retard)
   End Function

   ''' <summary>
   ''' Updates the datafields of a template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Update(ByVal Layer As Integer, ByVal TemplateData As Template) As ReturnInfo
      Return CG_Update(DefaultChannel, Layer, TemplateData)
   End Function

   ''' <summary>
   ''' Updates the datafields of a template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Update(ByVal Layer As Integer, ByVal TemplateData As Template, ByVal Retard As Retard) As ReturnInfo
      Return CG_Update(DefaultChannel, Layer, TemplateData, Retard)
   End Function

   ''' <summary>
   ''' Updates the datafields of a template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Play function.</remarks>
   Public Function CG_Update(ByVal Channel As Integer, ByVal Layer As Integer, ByVal TemplateData As Template, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} UPDATE {4} {0}{3}{0}", ChrW(&H22), Channel, Layer, TemplateData.TemplateDataText(FormatTextsForHTML), FlashLayer))
   End Function

   ''' <summary>
   ''' Updates the datafields of a template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="TemplateData">The Template object to get fields from</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Play function.</remarks>
   Public Function CG_Update(ByVal Channel As Integer, ByVal Layer As Integer, ByVal TemplateData As Template, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} UPDATE {4} {0}{3}{0}", ChrW(&H22), Channel, Layer, TemplateData.TemplateDataText(FormatTextsForHTML), FlashLayer), Retard)
   End Function


   ''' <summary>
   ''' Updates the datafields of a template from a Dataset
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DataSetName">The name of the DataSet already stored on the server</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function CG_Update(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DataSetName As String, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} UPDATE {4} {0}{3}{0}", ChrW(&H22), Channel, Layer, DataSetName, FlashLayer))
   End Function

   ''' <summary>
   ''' Updates the datafields of a template from a Dataset
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DataSetName">The name of the DataSet already stored on the server</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function CG_Update(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DataSetName As String, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} UPDATE {4} {0}{3}{0}", ChrW(&H22), Channel, Layer, DataSetName, FlashLayer), Retard)
   End Function


   ''' <summary>
   ''' Stops the display of a template
   ''' </summary>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Stop() As ReturnInfo
      Return CG_Stop(DefaultChannel, DefaultLayer)
   End Function

   ''' <summary>
   ''' Stops the display of a template
   ''' </summary>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Stop(ByVal Retard As Retard) As ReturnInfo
      Return CG_Stop(DefaultChannel, DefaultLayer, Retard)
   End Function

   ''' <summary>
   ''' Stops the display of a template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Stop(ByVal Layer As Integer) As ReturnInfo
      Return CG_Stop(DefaultChannel, Layer, New Retard)
   End Function

   ''' <summary>
   ''' Stops the display of a template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Stop(ByVal Layer As Integer, ByVal Retard As Retard) As ReturnInfo
      Return CG_Stop(DefaultChannel, Layer, Retard)
   End Function

   ''' <summary>
   ''' Stops the display of a template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Stop function.</remarks>
   Public Function CG_Stop(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {0}-{1} STOP {2}", Channel, Layer, FlashLayer))
   End Function

   ''' <summary>
   ''' Stops the display of a template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Stop function.</remarks>
   Public Function CG_Stop(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {0}-{1} STOP {2}", Channel, Layer, FlashLayer), Retard)
   End Function


   ''' <summary>
   ''' Calls the next step inside a template
   ''' </summary>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Next() As ReturnInfo
      Return CG_Next(DefaultChannel, DefaultLayer, New Retard)
   End Function

   ''' <summary>
   ''' Calls the next step inside a template
   ''' </summary>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel DefaultLayer</remarks>
   Public Function CG_Next(ByVal Retard As Retard) As ReturnInfo
      Return CG_Next(DefaultChannel, DefaultLayer, Retard)
   End Function

   ''' <summary>
   ''' Calls the next step inside a template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Next(ByVal Layer As Integer) As ReturnInfo
      Return CG_Next(DefaultChannel, Layer)
   End Function

   ''' <summary>
   ''' Calls the next step inside a template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function CG_Next(ByVal Layer As Integer, ByVal Retard As Retard) As ReturnInfo
      Return CG_Next(DefaultChannel, Layer, Retard)
   End Function

   ''' <summary>
   ''' Calls the next step inside a template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Next function.</remarks>
   Public Function CG_Next(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {0}-{1} NEXT {2}", Channel, Layer, FlashLayer))
   End Function

   ''' <summary>
   ''' Calls the next step inside a template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Next function.</remarks>
   Public Function CG_Next(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {0}-{1} NEXT {2}", Channel, Layer, FlashLayer), Retard)
   End Function


   ''' <summary>
   ''' Invokes a parameterless procedure inside the template
   ''' </summary>
   ''' <param name="Func">The function to be executed</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function CG_Invoke(ByVal Func As String) As ReturnInfo
      Return CG_Invoke(DefaultChannel, DefaultLayer, Func)
   End Function

   ''' <summary>
   ''' Invokes a parameterless procedure inside the template
   ''' </summary>
   ''' <param name="Func">The function to be executed</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function CG_Invoke(ByVal Func As String, ByVal Retard As Retard) As ReturnInfo
      Return CG_Invoke(DefaultChannel, DefaultLayer, Func, Retard)
   End Function

   ''' <summary>
   ''' Invokes a parameterless procedure inside the template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Func">The function to be executed</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function CG_Invoke(ByVal Layer As Integer, ByVal Func As String) As ReturnInfo
      Return CG_Invoke(DefaultChannel, Layer, Func)
   End Function

   ''' <summary>
   ''' Invokes a parameterless procedure inside the template
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Func">The function to be executed</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function CG_Invoke(ByVal Layer As Integer, ByVal Func As String, ByVal Retard As Retard) As ReturnInfo
      Return CG_Invoke(DefaultChannel, Layer, Func, Retard)
   End Function

   ''' <summary>
   ''' Invokes a parameterless procedure inside the template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Func">The function to be executed</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Invoke function.</remarks>
   Public Function CG_Invoke(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Func As String, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} INVOKE {4} {0}{3}{0}", ChrW(&H22), Channel, Layer, Func, FlashLayer))
   End Function

   ''' <summary>
   ''' Invokes a parameterless procedure inside the template
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Func">The function to be executed</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="FlashLayer">Optional, use only when absolutely needed</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main CG_Invoke function.</remarks>
   Public Function CG_Invoke(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Func As String, ByVal Retard As Retard, ByVal Optional FlashLayer As Integer = 1) As ReturnInfo
      Return Execute(String.Format("CG {1}-{2} INVOKE {4} {0}{3}{0}", ChrW(&H22), Channel, Layer, Func, FlashLayer), Retard)
   End Function

#End Region

#Region "Data Commands"

   ''' <summary>
   ''' Stores templatedata as a dataset on the server.
   ''' </summary>
   ''' <param name="DataSetName"></param>
   ''' <param name="TemplateData"></param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main Data_Store Function.</remarks>
   Public Function Data_Store(ByVal DataSetName As String, ByVal TemplateData As Template) As ReturnInfo
      Return Execute(String.Format("DATA STORE {0}{1}{0} {0}{2}{0}", ChrW(&H22), DataSetName, TemplateData.TemplateDataText(FormatTextsForHTML)))
   End Function

   'Public Function Data_Retrieve(DataSetName As String) As Template
   '   Dim ri As ReturnInfo = Execute(String.Format("DATA RETRIEVE {0}{1}{0}", ChrW(&H22), DataSetName), True)
   '   Dim tmpl As Template = Template.Parse(ri.Data)
   '   Return tmpl
   'End Function

   ''' <summary>
   ''' Get the data contained in a named dataset
   ''' </summary>
   ''' <param name="DataSetName">Name of the dataset to retrieve</param>
   ''' <returns></returns>
   Public Function Data_Retrieve(ByVal DataSetName As String) As String
      Dim ri As ReturnInfo = Execute(String.Format("DATA RETRIEVE {0}{1}{0}", ChrW(&H22), DataSetName), True)
      If ri.Number >= 200 And ri.Number < 400 Then
         Dim s As String = ri.Data
         Dim i As Integer = s.IndexOf("<", 0)
         Return s.Substring(i).Replace(vbCrLf, "")
      Else
         Return ""
      End If
   End Function

   ''' <summary>
   ''' Remove the dataset
   ''' </summary>
   ''' <param name="DataSetName">Name of the dataset to remove</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Data_Remove(ByVal DataSetName As String) As ReturnInfo
      Return Execute(String.Format("DATA REMOVE {0}{1}{0}", ChrW(&H22), DataSetName), True)
   End Function

   ''' <summary>
   ''' Get a list of DataSets
   ''' </summary>
   ''' <returns>Names of the DataSets</returns>
   Public Function GetDataSets() As List(Of String)
      Return GetDataSets("", False)
   End Function

   ''' <summary>
   ''' Get a list of DataSets
   ''' </summary>
   ''' <param name="AddEmptyEntry">Add an empty entry in the list. Makes it easier to fill Comboboxes.</param>
   ''' <returns>Names of the DataSets</returns>
   Public Function GetDataSets(ByVal AddEmptyEntry As Boolean) As List(Of String)
      Return GetDataSets("", AddEmptyEntry)
   End Function

   ''' <summary>
   ''' Get a list of DataSets
   ''' </summary>
   ''' <param name="FolderName">Subfolder to list</param>
   ''' <param name="AddEmptyEntry">Add an empty entry in the list. Makes it easier to fill Comboboxes.</param>
   ''' <returns>Names of the DataSets</returns>
   Public Function GetDataSets(ByVal FolderName As String, ByVal AddEmptyEntry As Boolean) As List(Of String)

      Dim lst As List(Of String) = New List(Of String)
      If AddEmptyEntry Then
         lst.Add("")
      End If

      Dim dat As Object = Execute(String.Format("DATA LIST{0}", IIf(FolderName <> "", ChrW(&H22) + FolderName + ChrW(&H22), "")), True).Data
      If dat IsNot Nothing Then

         Dim Ret() As String = CType(dat, String).Split(CChar(Microsoft.VisualBasic.vbCrLf))
         Dim i As Integer = 0

         For Each s As String In Ret
            s = s.Trim.Replace(vbNullChar, "")
            If s.Contains(Chr(&H22)) Then
               i = s.IndexOf(Chr(&H22), 2)
               If i > -1 Then
                  Dim dataset As String = s.Substring(1, i - 1).Replace("\", "/").Replace(Chr(&H22), "")
                  lst.Add(dataset)
               End If
            Else
               If s <> "" Then
                  lst.Add(s.Replace("\", "/").Replace(Chr(&H22), ""))
               End If
            End If
         Next

      End If

      If lst.Count = 0 Then
         lst.Add(" ")
      End If

      Return lst

   End Function

#End Region

#Region "Video Commands"

   ''' <summary>
   ''' Simple Load Command
   ''' </summary>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <returns></returns>
   Public Function Load(ByVal MediaName As String) As ReturnInfo
      Return Load(DefaultChannel, DefaultLayer, MediaName)
   End Function

   ''' <summary>
   ''' Simple Load Command
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <returns></returns>
   Public Function Load(ByVal Layer As Integer, ByVal MediaName As String) As ReturnInfo
      Return Load(DefaultChannel, Layer, MediaName)
   End Function

   ''' <summary>
   ''' Simple Load Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   Public Function Load(ByVal Channel As Integer, ByVal Layer As Integer, ByVal MediaName As String) As ReturnInfo
      Return Execute(String.Format("Load {1}-{2} {0}{3}{0}", ChrW(&H22), Channel, Layer, MediaName))
   End Function

   ''' <summary>
   ''' Simple LoadBG Comand
   ''' </summary>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <returns></returns>
   Public Function LoadBG(ByVal MediaName As String) As ReturnInfo
      Return LoadBG(DefaultChannel, DefaultLayer, MediaName, False)
   End Function

   ''' <summary>
   ''' Simple LoadBG Command
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <returns></returns>
   Public Function LoadBG(ByVal Layer As Integer, ByVal MediaName As String) As ReturnInfo
      Return LoadBG(DefaultChannel, Layer, MediaName, False)
   End Function

   ''' <summary>
   ''' Simple LoadBG Command
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <param name="LoopPlayback">Loops if true</param>
   ''' <returns></returns>
   Public Function LoadBG(ByVal Layer As Integer, ByVal MediaName As String, ByVal LoopPlayback As Boolean) As ReturnInfo
      Return LoadBG(DefaultChannel, Layer, MediaName, LoopPlayback)
   End Function

   ''' <summary>
   ''' Simple LoadBG Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <param name="LoopPlayback">Loops if true</param>
   Public Function LoadBG(ByVal Channel As Integer, ByVal Layer As Integer, ByVal MediaName As String, ByVal LoopPlayback As Boolean) As ReturnInfo
      If MediaName.StartsWith("http") Then
         Return Execute(String.Format("LOADBG {1}-{2} [HTML] {0}{3}{0} {4}", ChrW(&H22), Channel, Layer, MediaName, IIf(LoopPlayback, "LOOP", "")))
      Else
         Return Execute(String.Format("LOADBG {1}-{2} {0}{3}{0} {4}", ChrW(&H22), Channel, Layer, MediaName, IIf(LoopPlayback, "LOOP", "")))
      End If
   End Function

   ''' <summary>
   ''' Complex LoadBG Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <param name="LoopPlayback">Loops if true</param>
   ''' <param name="EffectType">In animation effect type</param>
   ''' <param name="EffectDuration">Durration of the in effect in frames</param>
   ''' <param name="Direction">Direction of the in effect</param>
   ''' <param name="Tween">Tween to be used on in effect</param>
   ''' <param name="AutoPlay">Enable AutoPlay if true</param>
   Public Function LoadBG(ByVal Channel As Integer, ByVal Layer As Integer, ByVal MediaName As String, ByVal LoopPlayback As Boolean, ByVal EffectType As String, ByVal EffectDuration As Integer, ByVal Direction As String, ByVal Tween As String, ByVal AutoPlay As Boolean) As ReturnInfo
      If MediaName.StartsWith("http") Then
         Return Execute(String.Format("LOADBG {1}-{2} [HTML] {0}{3}{0} {4} {5} {6} {7}{8}{9}", ChrW(&H22), Channel, Layer, MediaName, EffectType, EffectDuration, Tween, Direction, IIf(LoopPlayback, " LOOP", ""), IIf(AutoPlay, " AUTO", "")))
      Else
         '                                                                              0           1        2      3          4           5               6      7          8                               9
         Return Execute(String.Format("LOADBG {1}-{2} {0}{3}{0} {4} {5} {6} {7}{8}{9}", ChrW(&H22), Channel, Layer, MediaName, EffectType, EffectDuration, Tween, Direction, IIf(LoopPlayback, " LOOP", ""), IIf(AutoPlay, " AUTO", "")))
      End If
   End Function

   ''' <summary>
   ''' Play after Load / LoadBG
   ''' </summary>
   Public Function Play() As ReturnInfo
      Return Play(DefaultChannel, DefaultLayer)
   End Function

   ''' <summary>
   ''' Play after Load / LoadBG
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   Public Function Play(ByVal Layer As Integer) As ReturnInfo
      Return Play(DefaultChannel, Layer)
   End Function

   ''' <summary>
   ''' Play after Load / LoadBG
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   Public Function Play(ByVal Channel As Integer, ByVal Layer As Integer) As ReturnInfo
      Return Execute(String.Format("PLAY {0}-{1}", Channel, Layer))
   End Function

   ''' <summary>
   ''' Simple Play Comand
   ''' </summary>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <returns></returns>
   Public Function Play(ByVal MediaName As String) As ReturnInfo
      Return Play(DefaultChannel, DefaultLayer, MediaName, False)
   End Function

   ''' <summary>
   ''' Simple Play Comand
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <returns></returns>
   Public Function Play(ByVal Layer As Integer, ByVal MediaName As String) As ReturnInfo
      Return Play(DefaultChannel, Layer, MediaName, False)
   End Function

   ''' <summary>
   ''' Simple Play Comand
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <param name="LoopPlayback">Loops if true</param>
   ''' <returns></returns>
   Public Function Play(ByVal Layer As Integer, ByVal MediaName As String, ByVal LoopPlayback As Boolean) As ReturnInfo
      Return Play(DefaultChannel, Layer, MediaName, LoopPlayback)
   End Function

   ''' <summary>
   ''' Simple Play Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <param name="LoopPlayback">Loops if true</param>
   Public Function Play(ByVal Channel As Integer, ByVal Layer As Integer, ByVal MediaName As String, ByVal LoopPlayback As Boolean) As ReturnInfo
      If MediaName.StartsWith("http") Then
         Return Execute(String.Format("PLAY {1}-{2} [HTML] {0}{3}{0} {4}", ChrW(&H22), Channel, Layer, MediaName, IIf(LoopPlayback, "LOOP", "")))
      Else
         Return Execute(String.Format("PLAY {1}-{2} {0}{3}{0} {4}", ChrW(&H22), Channel, Layer, MediaName, IIf(LoopPlayback, "LOOP", "")))
      End If
   End Function

   ''' <summary>
   ''' Complex Play Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="MediaName">The name of the clip or image</param>
   ''' <param name="LoopPlayback">Loops if true</param>
   ''' <param name="EffectType">In animation effect type</param>
   ''' <param name="EffectDuration">Durration of the in effect in frames</param>
   ''' <param name="Direction">Direction of the in effect</param>
   ''' <param name="Tween">Tween to be used aon in effect</param>
   Public Function Play(ByVal Channel As Integer, ByVal Layer As Integer, ByVal MediaName As String, ByVal LoopPlayback As Boolean, ByVal EffectType As String, ByVal EffectDuration As Integer, ByVal Direction As String, ByVal Tween As String) As ReturnInfo
      If MediaName.StartsWith("http") Then
         Return Execute(String.Format("PLAY {1}-{2} [HTML] {0}{3}{0} {8} {4} {5} {6} {7}", ChrW(&H22), Channel, Layer, MediaName, EffectType, EffectDuration, Tween, Direction, IIf(LoopPlayback, "LOOP", "")))
      Else
         Return Execute(String.Format("PLAY {1}-{2} {0}{3}{0} {8} {4} {5} {6} {7}", ChrW(&H22), Channel, Layer, MediaName, EffectType, EffectDuration, Tween, Direction, IIf(LoopPlayback, "LOOP", "")))
      End If
   End Function

   ''' <summary>
   ''' Plays a local UDP stream from port 1234
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   Public Function PlayUDP(ByVal Channel As Integer, ByVal Layer As Integer) As ReturnInfo
      Return PlayUDP(Channel, Layer, "127.0.0.1", 1234)
   End Function

   ''' <summary>
   ''' Plays a local UDP stream from the given port
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Port">Port of the stream</param>
   Public Function PlayUDP(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Port As Integer) As ReturnInfo
      Return PlayUDP(Channel, Layer, "127.0.0.1", Port)
   End Function

   ''' <summary>
   ''' Plays a UDP stream
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="IPAddress">IP Address of the stream</param>
   ''' <param name="Port">Port of the stream</param>
   Public Function PlayUDP(ByVal Channel As Integer, ByVal Layer As Integer, ByVal IPAddress As String, ByVal Port As Integer) As ReturnInfo
      Return Execute(String.Format("PLAY {0}-{1} udp://{2}:{3}", Channel, Layer, IPAddress, Port))
   End Function

   ''' <summary>
   ''' Plays a stream
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Videolink">The video link to play</param>
   ''' <returns></returns>
   Public Function PlayStream(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Videolink As String) As ReturnInfo
      Return Execute(String.Format("PLAY {0}-{1} {2}", Channel, Layer, Videolink))
   End Function


   ''' <summary>
   ''' Stop Command
   ''' </summary>
   Public Function Stopp() As ReturnInfo
      Return Stopp(DefaultChannel, DefaultLayer)
   End Function

   ''' <summary>
   ''' Stop Command
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   Public Function Stopp(ByVal Layer As Integer) As ReturnInfo
      Return Stopp(DefaultChannel, Layer)
   End Function

   ''' <summary>
   ''' Stop Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   Public Function Stopp(ByVal Channel As Integer, ByVal Layer As Integer) As ReturnInfo
      Return Execute(String.Format("STOP {0}-{1}", Channel, Layer))
   End Function

#End Region

#Region "Mixer Commands"

   Private _MixerWidth As Integer = 1920
   Private _MixerHeight As Integer = 1080

   ''' <summary>
   ''' Used to set the channels resolution for mixer effects like Mixer_Fill etc.
   ''' </summary>
   ''' <param name="width">Channels width, defaults to 1920</param>
   ''' <param name="height">Channels height, defaults to 1080</param>
   ''' <remarks></remarks>
   Public Sub SetMixerResolution(ByVal width As Integer, ByVal height As Integer)
      _MixerWidth = width
      _MixerHeight = height
   End Sub

   ''' <summary>
   ''' Fades a layer in
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to fade, in frames</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function Mixer_FadeIn(ByVal Layer As Integer, ByVal Duration As Integer) As ReturnInfo
      Return Mixer_FadeIn(DefaultChannel, Layer, Duration, False)
   End Function

   ''' <summary>
   ''' Fades a layer in
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to fade, in frames</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main Mixer_FadeIn function</remarks>
   Public Function Mixer_FadeIn(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer) As ReturnInfo
      Return Mixer_FadeIn(Channel, Layer, Duration, False)
   End Function

   ''' <summary>
   ''' Fades a layer in
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to fade, in frames</param>
   ''' <param name="WithAudio">Also fades audio if true</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main Mixer_FadeIn function</remarks>
   Public Function Mixer_FadeIn(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, WithAudio As Boolean) As ReturnInfo
      If WithAudio Then
         Mixer_Volume(Channel, Layer, 1, Duration)
      End If
      Return Execute(String.Format("MIXER {0}-{1} OPACITY 1 {2}", Channel, Layer, Duration))
   End Function

   ''' <summary>
   ''' Fades a layer out
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to fade, in frames</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function Mixer_FadeOut(ByVal Layer As Integer, ByVal Duration As Integer) As ReturnInfo
      Return Mixer_FadeOut(DefaultChannel, Layer, Duration, False)
   End Function

   ''' <summary>
   ''' Fades a layer out
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to fade, in frames</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Mixer_FadeOut(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer) As ReturnInfo
      Return Mixer_FadeOut(Channel, Layer, Duration, False)
   End Function

   ''' <summary>
   ''' Fades a layer out
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to fade, in frames</param>
   ''' <param name="WithAudio">Also fades audio if true</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>The main Mixer_FadeOut function</remarks>
   Public Function Mixer_FadeOut(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, WithAudio As Boolean) As ReturnInfo
      If WithAudio Then
         Mixer_Volume(Channel, Layer, 0, Duration)
      End If
      Return Execute(String.Format("MIXER {0}-{1} OPACITY 0 {2}", Channel, Layer, Duration))
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="X">X position of the rectanngle, in pixels</param>
   ''' <param name="Y">X position of the rectanngle, in pixels</param>
   ''' <param name="Width">Width of the rectangle, in pixels</param>
   ''' <param name="Height">Height of the rectangle, in pixels</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Arguments as Integer</remarks> 
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer) As ReturnInfo
      Return Mixer_Fill(Channel, Layer, 0, New Drawing.Rectangle(X, Y, Width, Height))
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="X">X position of the rectanngle, in pixels</param>
   ''' <param name="Y">X position of the rectanngle, in pixels</param>
   ''' <param name="Width">Width of the rectangle, in pixels</param>
   ''' <param name="Height">Height of the rectangle, in pixels</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Arguments as Doubles</remarks> 
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal X As Double, ByVal Y As Double, ByVal Width As Double, ByVal Height As Double) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} FILL {2:0.00000}  {3:0.00000} {4:0.00000} {5:0.00000}", Channel, Layer, X, Y, Width, Height))
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="rect">A Drawing.Rectangle that defines the position, in pixels</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function Mixer_Fill(ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle) As ReturnInfo
      Return Mixer_Fill(DefaultChannel, Layer, Duration, rect, "linear", False)
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="rect">A Drawing.Rectangle that defines the position, in pixels</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle) As ReturnInfo
      Return Mixer_Fill(Channel, Layer, Duration, rect, "linear", False)
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="rect">A Drawing.Rectangle that defines the position, in pixels</param>
   ''' <param name="tween">A string containing the tween. See AMCP doc for details</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle, ByVal tween As String) As ReturnInfo
      Return Mixer_Fill(Channel, Layer, Duration, rect, tween, False)
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="rect">A Drawing.Rectangle that defines the position, in pixels</param>
   ''' <param name="tween">A string containing the tween. See AMCP doc for details</param>
   ''' <param name="Defer">Does not execute the command, wait for a Mixer_Commit</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} FILL {2:0.00000}  {3:0.00000} {4:0.00000} {5:0.00000} {6} {7}{8}", Channel, Layer, rect.X / _MixerWidth, rect.Y / _MixerHeight, rect.Width / _MixerWidth, rect.Height / _MixerHeight, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="X">X position of the rectanngle, in pixels</param>
   ''' <param name="Y">X position of the rectanngle, in pixels</param>
   ''' <param name="Width">Width of the rectangle, in pixels</param>
   ''' <param name="Height">Height of the rectangle, in pixels</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="tween">A string containing the tween. See AMCP doc for details</param>
   ''' <param name="Defer">Does not execute the command, wait for a Mixer_Commit</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Arguments as Doubles</remarks> 
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal X As Double, ByVal Y As Double, ByVal Width As Double, ByVal Height As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} FILL {2:0.00000}  {3:0.00000} {4:0.00000} {5:0.00000} {6} {7}{8}", Channel, Layer, X, Y, Width, Height, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function


   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DVEffect">As string formated according to the values in the AMCP protocoll of the MIXER_FILL effect</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Old version</remarks>  
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DVEffect As String) As ReturnInfo
      Return Mixer_Fill(Channel, Layer, DVEffect, 0, False)
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DVEffect">As string formated according to the values in the AMCP protocoll of the MIXER_FILL effect</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Old version</remarks>  
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DVEffect As String, ByVal Duration As Integer) As ReturnInfo
      Return Mixer_Fill(Channel, Layer, DVEffect, Duration, False)
   End Function

   ''' <summary>
   ''' Scales a layer in a given rectanngle
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DVEffect">As string formated according to the values in the AMCP protocoll of the MIXER_FILL effect</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="Defer">Does not execute the command, wait for a Mixer_Commit</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Old version</remarks>  
   Public Function Mixer_Fill(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DVEffect As String, ByVal Duration As Integer, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format("MIXER {0}-{1} FILL {2} {3}{4}", Channel, Layer, DVEffect.Replace(",", "."), Duration, IIf(Defer, " DEFER", "")))
   End Function


   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="rect">A Drawing.Rectangle that defines the clipping, in pixels</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Defaults to DefaultChannel</remarks>
   Public Function Mixer_Clip(ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle) As ReturnInfo
      Return Mixer_Clip(DefaultChannel, Layer, Duration, rect, "linear", False)
   End Function

   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Channel"></param>
   ''' <param name="Layer"></param>
   ''' <param name="Duration"></param>
   ''' <param name="rect"></param>
   ''' <returns></returns>
   Public Function Mixer_Clip(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle) As ReturnInfo
      Return Mixer_Clip(Channel, Layer, Duration, rect, "linear", False)
   End Function

   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="rect">A Drawing.Rectangle that defines the position, in pixels</param>
   ''' <param name="tween">A string containing the tween. See AMCP doc for details</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Mixer_Clip(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle, ByVal tween As String) As ReturnInfo
      Return Mixer_Clip(Channel, Layer, Duration, rect, tween, False)
   End Function

   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="rect">A Drawing.Rectangle that defines the position, in pixels</param>
   ''' <param name="tween">A string containing the tween. See AMCP doc for details</param>
   ''' <param name="Defer">Does not execute the command, wait for a Mixer_Commit</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Mixer_Clip(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Duration As Integer, ByVal rect As Drawing.Rectangle, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} CLIP {2:0.00000} {3:0.00000} {4:0.00000} {5:0.00000} {6} {7}{8}", Channel, Layer, rect.X / _MixerWidth, rect.Y / _MixerHeight, rect.Width / _MixerWidth, rect.Height / _MixerHeight, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="X">X position of the rectanngle, in pixels</param>
   ''' <param name="Y">X position of the rectanngle, in pixels</param>
   ''' <param name="Width">Width of the rectangle, in pixels</param>
   ''' <param name="Height">Height of the rectangle, in pixels</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="tween">A string containing the tween. See AMCP doc for details</param>
   ''' <param name="Defer">Does not execute the command, wait for a Mixer_Commit</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Arguments as Doubles</remarks> 
   Public Function Mixer_Clip(ByVal Channel As Integer, ByVal Layer As Integer, ByVal X As Double, ByVal Y As Double, ByVal Width As Double, ByVal Height As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} CLIP {2:0.00000} {3:0.00000} {4:0.00000} {5:0.00000} {6} {7}{8}", Channel, Layer, X, Y, Width, Height, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DVEffect">As string formated according to the values in the AMCP protocoll of the MIXER_FILL effect</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Old version</remarks>  
   Public Function Mixer_Clip(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DVEffect As String) As ReturnInfo
      Return Mixer_Clip(Channel, Layer, DVEffect, 0, False)
   End Function

   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DVEffect">As string formated according to the values in the AMCP protocoll of the MIXER_FILL effect</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Old version</remarks>  
   Public Function Mixer_Clip(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DVEffect As String, ByVal Duration As Integer) As ReturnInfo
      Return Mixer_Clip(Channel, Layer, DVEffect, Duration, False)
   End Function

   ''' <summary>
   ''' Clips a layer around it's edges
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="DVEffect">As string formated according to the values in the AMCP protocoll of the MIXER_FILL effect</param>
   ''' <param name="Duration">The time to move into the given position, in frames</param>
   ''' <param name="Defer">Does not execute the command, wait for a Mixer_Commit</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Old version</remarks>  
   Public Function Mixer_Clip(ByVal Channel As Integer, ByVal Layer As Integer, ByVal DVEffect As String, ByVal Duration As Integer, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format("MIXER {0}-{1} CLIP {2} {3}{4}", Channel, Layer, DVEffect, Duration, IIf(Defer, " DEFER", "")))
   End Function

   ''' <summary>
   ''' Executes all defered Mixer commands of the given channel.
   ''' </summary>
   ''' <param name="Channel"></param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Mixer_Commit(ByVal Channel As Integer) As ReturnInfo
      Return Execute(String.Format("MIXER {0} COMMIT", Channel))
   End Function


   Public Function Mixer_Opacity(ByVal Layer As Integer, ByVal Opacity As Single) As ReturnInfo
      Return Mixer_Opacity(1, Layer, Opacity, 1)
   End Function

   Public Function Mixer_Opacity(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Opacity As Single) As ReturnInfo
      Return Mixer_Opacity(Channel, Layer, Opacity, 1)
   End Function

   Public Function Mixer_Opacity(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Opacity As Single, ByVal Duration As Integer) As ReturnInfo
      Return Execute(String.Format("MIXER {0}-{1} OPACITY {2} {3}", Channel, Layer, Opacity, Duration))
   End Function

   Public Function Mixer_Opacity(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Opacity As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} OPACITY {2:0.00000} {3} {4}{5}", Channel, Layer, Opacity, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function


   Public Function Mixer_Clear(ByVal Channel As Integer) As ReturnInfo
      Return Execute(String.Format("MIXER {0} CLEAR", Channel))
   End Function

   'Simple commands, whitout a lot of overlaods

   Public Function LoadBG_Input(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Input As Integer, ByVal BroadcastFormat As String) As ReturnInfo
      Return Execute(String.Format("LOADBG {0}-{1} DECKLINK DEVICE {2} FORMAT {3}", Channel, Layer, Input, BroadcastFormat))
   End Function

   Public Function Play_Input(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Input As Integer, ByVal BroadcastFormat As String) As ReturnInfo
      Return Execute(String.Format("PLAY {0}-{1} DECKLINK DEVICE {2} FORMAT {3}", Channel, Layer, Input, BroadcastFormat))
   End Function


   Public Function Mixer_Anchor(ByVal Channel As Integer, ByVal Layer As Integer, ByVal X As Double, ByVal Y As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} ANCHOR {2:0.00000}  {3:0.00000} {4} {5}{6}", Channel, Layer, X, Y, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Blend(ByVal Channel As Integer, ByVal Layer As Integer, ByVal BlendMode As String) As ReturnInfo
      Return Execute(String.Format("MIXER {0}-{1} BLEND {2}", Channel, Layer, BlendMode))
   End Function

   Public Function Mixer_Brightness(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Brightness As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} BRIGHTNESS {2:0.00000} {3} {4}{5}", Channel, Layer, Brightness, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Contrast(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Contrast As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} CONTRAST {2:0.00000} {3} {4}{5}", Channel, Layer, Contrast, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Crop(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Left As Double, ByVal Top As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} CROP {2:0.00000}  {3:0.00000} {4:0.00000} {5:0.00000} {6} {7}{8}", Channel, Layer, Left, Top, Right, Bottom, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Perspective(ByVal Channel As Integer, ByVal Layer As Integer, ByVal TopLeftX As Double, ByVal TopLeftY As Double, ByVal TopRightX As Double, ByVal TopRightY As Double, ByVal BottomRightX As Double, ByVal BottomRightY As Double, ByVal BottomLeftX As Double, ByVal BottomLeftY As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} PERSPECTIVE {2:0.00000}  {3:0.00000} {4:0.00000} {5:0.00000} {6:0.00000} {7:0.00000} {8:0.00000} {9} {10}{11}{12}", Channel, Layer, TopLeftX, TopLeftY, TopRightX, TopRightY, BottomRightX, BottomRightY, BottomLeftX, BottomLeftY, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Levels(ByVal Channel As Integer, ByVal Layer As Integer, ByVal MinInput As Double, ByVal MaxInput As Double, ByVal Gamma As Double, ByVal MinOutput As Double, ByVal MaxOutput As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} LEVELS {2:0.00000}  {3:0.00000} {4:0.00000} {5:0.00000} {6:0.00000} {7} {8}{9}", Channel, Layer, MinInput, MaxInput, Gamma, MinOutput, MaxOutput, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Keyer(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Activate As Boolean, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format("MIXER {0}-{1} KEYER {2}{3}", Channel, Layer, IIf(Activate, "1", "0"), IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Rotation(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Angle As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} ROTATION {2: 0.00000} {3} {4}{5}", Channel, Layer, Angle, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Saturation(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Saturation As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} SATURATION {2:0.00000} {3} {4}{5}", Channel, Layer, Saturation, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

   Public Function Mixer_Volume(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Volume As Double) As ReturnInfo
      Return Mixer_Volume(Channel, Layer, Volume, 0, "linear", False)
   End Function

   Public Function Mixer_Volume(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Volume As Double, ByVal Duration As Integer) As ReturnInfo
      Return Mixer_Volume(Channel, Layer, Volume, Duration, "linear", False)
   End Function

   Public Function Mixer_Volume(ByVal Channel As Integer, ByVal Layer As Integer, ByVal Volume As Double, ByVal Duration As Integer, ByVal tween As String, ByVal Defer As Boolean) As ReturnInfo
      Return Execute(String.Format(CultureInfo.InvariantCulture, "MIXER {0}-{1} VOLUME {2:0.00000} {3} {4}{5}", Channel, Layer, Volume, Duration, tween, IIf(Defer, " DEFER", "")))
   End Function

#End Region

#Region "Clear Commands"

   Public Function ClearChannel(ByVal Channel As Integer) As ReturnInfo
      Return Execute(String.Format("CLEAR {0}", Channel))
   End Function

   ''' <summary>
   ''' Clear Command
   ''' </summary>
   Public Function Clear() As ReturnInfo
      Return Clear(DefaultChannel, DefaultLayer)
   End Function

   ''' <summary>
   ''' Clear Command
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   Public Function Clear(ByVal Layer As Integer) As ReturnInfo
      Return Clear(DefaultChannel, Layer)
   End Function

   ''' <summary>
   ''' Clear Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   Public Function Clear(ByVal Channel As Integer, ByVal Layer As Integer) As ReturnInfo
      Return Execute(String.Format("CLEAR {0}-{1}", Channel, Layer))
   End Function

   ''' <summary>
   ''' Clears all layers defined by FirstClearlayer and LastClearlayer properties.
   ''' </summary>
   Public Sub Clearlayers()
      Clearlayers(1)
   End Sub

   ''' <summary>
   ''' Clears all layers defined by FirstClearlayer and LastClearlayer properties.
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   Public Sub Clearlayers(ByVal Channel As Integer)
      For i As Integer = FirstClearlayer To LastClearlayer
         Execute(String.Format("CLEAR {0}-{1}", Channel, i))
      Next
   End Sub

#End Region

#Region "Misc Commands"

   ''' <summary>
   ''' Swap Command
   ''' </summary>
   ''' <param name="FirstLayer">The fisrt layer to swap</param>
   ''' <param name="SecondLayer">The second layer to swap</param>
   Public Function Swap(ByVal FirstLayer As Integer, ByVal SecondLayer As Integer) As ReturnInfo
      Return Swap(1, FirstLayer, SecondLayer)
   End Function

   ''' <summary>
   ''' Swap Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="FirstLayer">The fisrt layer to swap</param>
   ''' <param name="SecondLayer">The second layer to swap</param>
   Public Function Swap(ByVal Channel As Integer, ByVal FirstLayer As Integer, ByVal SecondLayer As Integer) As ReturnInfo
      Return Execute(String.Format("SWAP {0}-{1} {0}-{2}", Channel, FirstLayer, SecondLayer))
   End Function

#End Region

#Region "Info Commands"

   'Helper
   Private Sub GetChannelInfos()

      _ChannelInfos = New List(Of ChannelInfo)
      Try

         Dim s As String = Execute("INFO").Data
         If s <> "" Then
            Dim arr() As String = s.Split(vbCr.ToCharArray)
            For Each t As String In arr

               Dim parts() As String = t.Replace(vbLf, "").Split(" ".ToCharArray)
               If parts.Count >= 2 AndAlso IsNumeric(parts(0)) Then
                  _ChannelInfos.Add(New ChannelInfo(parts(1)))
               End If
            Next

         End If

      Catch ex As Exception
         'Ignore
      End Try

      If _ChannelInfos.Count = 0 Then
         _ChannelInfos.Add(New ChannelInfo())
      End If

   End Sub

   ''' <summary>
   ''' Get the horizontal resolution of a channel
   ''' </summary>
   ''' <param name="Channel">Channel to query</param>
   ''' <returns>Horizontal resolution</returns>
   Public Function ChannelHorizontalResolution(ByVal Channel As Integer) As Integer

      If _ChannelInfos Is Nothing Then
         GetChannelInfos()
      End If

      If Channel - 1 < _ChannelInfos.Count Then
         Return _ChannelInfos(Channel - 1).Width
      Else
         Return 0
      End If

   End Function

   ''' <summary>
   ''' Get the vertical resolution of a channel
   ''' </summary>
   ''' <param name="Channel">Channel to query</param>
   ''' <returns>Vertical resolution</returns>
   Public Function ChannelVerticalResolution(ByVal Channel As Integer) As Integer

      If _ChannelInfos Is Nothing Then
         GetChannelInfos()
      End If

      If Channel - 1 < _ChannelInfos.Count Then
         Return _ChannelInfos(Channel - 1).Height
      Else
         Return 0
      End If

   End Function

   ''' <summary>
   ''' Get the frame-rate of a channel
   ''' </summary>
   ''' <param name="Channel">Channel to query</param>
   ''' <returns>Frame-rate in fps</returns>
   Public Function ChannelFramerate(ByVal Channel As Integer) As Single

      If _ChannelInfos Is Nothing Then
         GetChannelInfos()
      End If

      If Channel - 1 < _ChannelInfos.Count Then
         Return _ChannelInfos(Channel - 1).Framerate
      Else
         Return 0
      End If

   End Function

   ''' <summary>
   ''' Query if a channel is interlaced or nor
   ''' </summary>
   ''' <param name="Channel">Channel to query</param>
   ''' <returns>True if the channel is interlaced</returns>
   Public Function ChannelIsInterlaced(ByVal Channel As Integer) As Boolean

      If _ChannelInfos Is Nothing Then
         GetChannelInfos()
      End If

      If Channel - 1 < _ChannelInfos.Count Then
         Return _ChannelInfos(Channel - 1).IsInterlaced
      Else
         Return True
      End If

   End Function

#End Region

#Region "NDI Commands"

   Public Class NDISource
      Public Property ComputerName As String = ""
      Public Property ChannelName As String = ""

      Public Sub New(data As String)

         Dim start As Integer = data.IndexOf(ChrW(&H22), 0)
         Dim ende As Integer = data.IndexOf(ChrW(&H22), start + 1)
         If ende = -1 Then ende = data.Length
         Dim bracket As Integer = data.IndexOf("(", start + 1)

         ComputerName = data.Substring(start + 1, bracket - start - 2)
         ChannelName = data.Substring(bracket + 1, ende - bracket - 2)

      End Sub
      Public Sub New(computer As String, channel As String)
         ComputerName = computer
         ChannelName = channel
      End Sub

      Public Overrides Function ToString() As String
         Return String.Format("{0} ({1})", ComputerName, ChannelName)
      End Function

   End Class

   Public Function GetNDISources(ByVal AddEmptyEntry As Boolean) As List(Of String)

      Dim lst As List(Of String) = New List(Of String)
      If AddEmptyEntry Then
         lst.Add("")
      End If

      If Me.Version.Major >= 2 And Me.Version.Minor >= 2 Then

         Try

            Dim ri As ReturnInfo = Execute("NDI LIST", True)
            If ri IsNot Nothing AndAlso ri.Data IsNot Nothing Then

               Dim Ret() As String = ri.Data.Replace(vbLf, "").Split(CChar(vbCr))
               Dim i As Integer = 0

               For Each s As String In Ret
                  If s.Contains(ChrW(&H22)) Then
                     Dim nds As NDISource = New NDISource(s.Trim)
                     lst.Add(nds.ToString)
                  End If
               Next

            End If

         Catch ex As Exception
            'Ignore
         End Try

      End If

      If lst.Count = 0 Then
         lst.Add(" ")
      End If

      Return lst

   End Function

   ''' <summary>
   ''' Checks if a NDI Source is available
   ''' </summary>
   ''' <param name="ComputerName">The name of the Computer sending NDI</param>
   ''' <param name="ChannelName">The NDI chnnel</param>
   ''' <returns>True if it is available</returns>
   Public Function IsNDISourceAvailable(ByVal ComputerName As String, ByVal ChannelName As String) As Boolean
      Return IsNDISourceAvailable(New NDISource(ComputerName, ChannelName).ToString)
   End Function

   ''' <summary>
   ''' Checks if a NDI Source is available
   ''' </summary>
   ''' <param name="ComputerAndChannelName">Combined Computer- and ChannelName</param>
   ''' <returns>True if it is available</returns>
   Public Function IsNDISourceAvailable(ByVal ComputerAndChannelName As String) As Boolean

      Dim lst As List(Of String) = GetNDISources(False)
      Dim retval As Boolean = False

      For Each entry As String In lst
         If entry = ComputerAndChannelName Then
            retval = True
            Exit For
         End If
      Next

      Return retval

   End Function

   ''' <summary>
   ''' NDI Play Comand
   ''' </summary>
   ''' <param name="ComputerName">The name of the Computer sending NDI</param>
   ''' <param name="ChannelName">The NDI chnnel</param>
   Public Function NDIPlay(ByVal ComputerName As String, ByVal ChannelName As String) As ReturnInfo
      Return NDIPlay(DefaultChannel, DefaultLayer, ComputerName, ChannelName)
   End Function

   ''' <summary>
   ''' NDI Play Comand
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="ComputerName">The name of the Computer sending NDI</param>
   ''' <param name="ChannelName">The NDI chnnel</param>
   Public Function NDIPlay(ByVal Layer As Integer, ByVal ComputerName As String, ByVal ChannelName As String) As ReturnInfo
      Return NDIPlay(DefaultChannel, Layer, ComputerName, ChannelName)
   End Function

   ''' <summary>
   ''' NDI Play Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="ComputerName">The name of the Computer sending NDI</param>
   ''' <param name="ChannelName">The NDI chnnel</param>
   Public Function NDIPlay(ByVal Channel As Integer, ByVal Layer As Integer, ByVal ComputerName As String, ByVal ChannelName As String) As ReturnInfo
      Return Execute(String.Format("PLAY {1}-{2} {0}ndi://{3}/{4}{0}", ChrW(&H22), Channel, Layer, ComputerName, ChannelName))
   End Function

   ''' <summary>
   ''' NDI Play Command
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <param name="ComputerAndChannelName">Combined Computer- and ChannelName</param>
   Public Function NDIPlay(ByVal Channel As Integer, ByVal Layer As Integer, ByVal ComputerAndChannelName As String) As ReturnInfo
      Try
         Dim nds As NDISource = New NDISource(ComputerAndChannelName)
         Return NDIPlay(Channel, Layer, nds.ComputerName, nds.ChannelName)
      Catch ex As Exception
         Return New ReturnInfo
      End Try
   End Function

#End Region

#Region "Grab Commands"

   Private _grabChannel As Integer
   Private _grabLayer As Integer
   Private _grabPicPath As String
   Private _grabImgNum As Integer = 1
   Private _grabRi As ReturnInfo
   Private _grabUniqueId As String
   Private _returnBitmap As Boolean = False

   Private WithEvents _grabTimer As System.Windows.Forms.Timer = New System.Windows.Forms.Timer
   Private _grabTimerMode As Integer = 0

   Public Class GrabFinishEventArgs
      Inherits EventArgs

      Public Property Image As Drawing.Bitmap

      Public Property UniqueId As String

      Public Property GrabFilename As String

      Public Sub New(ByVal Image As Drawing.Bitmap, ByVal UniqueId As String)
         MyBase.New()
         Me.Image = Image
         Me.UniqueId = UniqueId
         Me.GrabFilename = ""
      End Sub

      Public Sub New(ByVal Image As Drawing.Bitmap, ByVal UniqueId As String, ByVal GrabFilename As String)
         MyBase.New()
         Me.Image = Image
         Me.UniqueId = UniqueId
         Me.GrabFilename = GrabFilename
      End Sub

   End Class

   Public Delegate Sub GrabFinishEventHandler(ByVal sender As Object, ByVal e As GrabFinishEventArgs)

   Public Event GrabFinish As GrabFinishEventHandler

   Public Sub StartGrab(ByVal channel As Integer, ByVal uniqueId As String)
      StartGrab(channel, -1, uniqueId, 4000, True)
   End Sub

   Public Sub StartGrab(ByVal channel As Integer, ByVal layer As Integer, ByVal uniqueId As String)
      StartGrab(channel, layer, uniqueId, 4000, True)
   End Sub

   Public Sub StartGrab(ByVal channel As Integer, ByVal uniqueId As String, ByVal initalDelay As Integer)
      StartGrab(channel, -1, uniqueId, initalDelay, True)
   End Sub

   Public Sub StartGrab(ByVal channel As Integer, ByVal uniqueId As String, ByVal initalDelay As Integer, returnBitmap As Boolean)
      StartGrab(channel, -1, uniqueId, initalDelay, returnBitmap)
   End Sub

   Public Sub StartGrab(ByVal channel As Integer, ByVal layer As Integer, ByVal uniqueId As String, ByVal initalDelay As Integer)
      StartGrab(channel, layer, uniqueId, initalDelay, True)
   End Sub
   Public Sub StartGrab(ByVal channel As Integer, ByVal layer As Integer, ByVal uniqueId As String, ByVal initalDelay As Integer, returnBitmap As Boolean)

      If Me.ServerAdress.Trim.ToLower <> "localhost" And Me.ServerAdress.Trim <> "127.0.0.1" Then
         Throw New Exception("The Grab-Function only works for local CasparCG-Server (localhost)")
         Exit Sub
      End If

      _grabChannel = channel
      _grabLayer = layer
      _grabUniqueId = uniqueId
      _returnBitmap = returnBitmap

      _grabPicPath = ServerPaths.MediaPath

      If _grabPicPath.EndsWith("\") Then
         _grabPicPath = _grabPicPath.Substring(0, _grabPicPath.Length - 1)
      End If

      If Not _grabPicPath.Substring(1, 1) = ":" Then
         Throw New Exception("The Grab-Function needs to have absolute paths configured in CasparCG.config.")
         Exit Sub
      End If

      'Delete old files once a day
      Dim delFn As String = String.Format("~{0:yyyyMMdd}.del", Date.Now)
      If Not IO.File.Exists(IO.Path.Combine(_grabPicPath, delFn)) Then

         Dim fn() As String = IO.Directory.GetFiles(_grabPicPath, "~*.del")

         Try
            For c As Integer = 0 To fn.Length - 1
               IO.File.Delete(IO.Path.Combine(_grabPicPath, fn(c)))
            Next
         Catch ex As Exception
            Debug.Print(ex.Message)
         End Try

         fn = IO.Directory.GetFiles(_grabPicPath, "~GRAB_*.png")
         For c As Integer = 0 To fn.Length - 1

            Try
               IO.File.Delete(IO.Path.Combine(_grabPicPath, fn(c)))
            Catch ex As Exception
               Debug.Print(ex.Message)
            End Try
         Next

         Dim fs As FileStream = File.Create(IO.Path.Combine(_grabPicPath, delFn))
         fs.Close()

      End If

      _grabTimerMode = 1
      _grabTimer.Interval = initalDelay
      _grabTimer.Start()

   End Sub

   Private Sub _grabTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles _grabTimer.Tick

      _grabTimer.Stop()

      Select Case _grabTimerMode
         Case 1
            Do
               If Not IO.File.Exists(IO.Path.Combine(_grabPicPath, String.Format("~GRAB_{0:00000}.png", _grabImgNum))) Then Exit Do
               _grabImgNum += 1
            Loop

            If _grabLayer = -1 Then
               _grabRi = Execute(String.Format("ADD {0} IMAGE ~GRAB_{1:00000}", _grabChannel, _grabImgNum))
            Else
               _grabRi = Execute(String.Format("ADD {0}-{1} IMAGE ~GRAB_{2:00000}", _grabChannel, _grabLayer, _grabImgNum))
            End If

            _grabTimerMode = 2
            _grabTimer.Interval = 1000
            _grabTimer.Start()

         Case 2
            If _grabRi.Number = 202 Then

               If _grabLayer = -1 Then
                  Execute(String.Format("REMOVE {0} IMAGE", _grabChannel))
               Else
                  Execute(String.Format("REMOVE {0}-{1} IMAGE", _grabChannel, _grabLayer))
               End If

               _grabTimerMode = 3
               _grabTimer.Interval = 1000
               _grabTimer.Start()

            Else
               RaiseEvent GrabFinish(Me, New GrabFinishEventArgs(Nothing, _grabUniqueId))
            End If

         Case 3
            Dim fn As String = IO.Path.Combine(_grabPicPath, String.Format("~GRAB_{0:00000}.png", _grabImgNum))
            If IO.File.Exists(fn) Then
               _grabTimerMode = 4
               _grabTimer.Interval = 1500
               _grabTimer.Start()
            Else
               _grabTimerMode = 3
               _grabTimer.Interval = 500
               _grabTimer.Start()
            End If

         Case 4
            Dim fn As String = String.Format("~GRAB_{0:00000}.png", _grabImgNum)

            Try
               If _returnBitmap Then
                  Dim bm As Drawing.Bitmap = New Drawing.Bitmap(IO.Path.Combine(_grabPicPath, fn))
                  RaiseEvent GrabFinish(Me, New GrabFinishEventArgs(bm, _grabUniqueId, fn))
               Else
                  RaiseEvent GrabFinish(Me, New GrabFinishEventArgs(Nothing, _grabUniqueId, fn))
               End If
            Catch ex As Exception
               'Ignore
            End Try

            _grabTimerMode = 1

         Case Else
            _grabTimerMode = 1

      End Select

   End Sub



   ''' <summary>
   ''' Grab the output to a bitmap
   ''' </summary>
   ''' <param name="channel">The channel in Caspar</param>
   ''' <returns>A Drawing.Bitmap object</returns>
   ''' <remarks>Defaults to channel 1</remarks>
   Public Function Grab(ByVal channel As Integer) As Drawing.Bitmap
      Return Grab(channel, -1)
   End Function


   ''' <summary>
   ''' Grab the output to a bitmap
   ''' </summary>
   ''' <param name="channel">The channel in Caspar</param>
   ''' <param name="layer">The layer number</param>
   ''' <returns>A Drawing.Bitmap object</returns>
   Public Function Grab(ByVal channel As Integer, ByVal layer As Integer) As Drawing.Bitmap

      If Me.ServerAdress.ToLower <> "localhost" Then
         Throw New Exception("The Grab-Function only works for local CasparCG-Server (localhost)")
         Return Nothing
         Exit Function
      End If

      If VersionSerialized >= 2000004 Then   'Server 2.0.4 Stable: Breaking change of PRINT command, use ADD IMAGE instead.

         Dim PicPath As String = ServerPaths.MediaPath

         If PicPath.EndsWith("\") Then
            PicPath = PicPath.Substring(0, PicPath.Length - 1)
         End If

         If Not PicPath.Substring(1, 1) = ":" Then
            Throw New Exception("The Grab-Function needs to have absolute paths configured in CasparCG.config.")
            Return Nothing
            Exit Function
         End If

         'Delete old files once a day
         Dim delFn As String = String.Format("~{0:yyyyMMdd}.del", Date.Now)
         If Not IO.File.Exists(IO.Path.Combine(PicPath, delFn)) Then

            Dim fn() As String = IO.Directory.GetFiles(PicPath, "~*.del")
            For c As Integer = 0 To fn.Length - 1

               Try
                  IO.File.Delete(IO.Path.Combine(PicPath, fn(c)))
               Catch ex As Exception
                  Debug.Print(ex.Message)
               End Try

            Next

            fn = IO.Directory.GetFiles(PicPath, "~GRAB_*.png")
            For c As Integer = 0 To fn.Length - 1

               Try
                  IO.File.Delete(IO.Path.Combine(PicPath, fn(c)))
               Catch ex As Exception
                  Debug.Print(ex.Message)
               End Try

            Next

            Dim fs As FileStream = File.Create(IO.Path.Combine(_grabPicPath, delFn))
            fs.Close()

         End If

         Threading.Thread.Sleep(2000)

         Dim imgNum As Integer = 1
         Do
            If Not IO.File.Exists(IO.Path.Combine(PicPath, String.Format("~GRAB_{0:00000}.png", imgNum))) Then Exit Do
            imgNum += 1
         Loop

         Dim ri As ReturnInfo
         If layer = -1 Then
            ri = Execute(String.Format("ADD {0} IMAGE ~GRAB_{1:00000}", channel, imgNum))
         Else
            ri = Execute(String.Format("ADD {0}-{1} IMAGE ~GRAB_{2:00000}", channel, layer, imgNum))
         End If

         Threading.Thread.Sleep(1000)

         If ri.Number = 202 Then

            If layer = -1 Then
               Execute(String.Format("REMOVE {0} IMAGE", channel))
            Else
               Execute(String.Format("REMOVE {0}-{1} IMAGE", channel, layer))
            End If

            Dim fn As String = IO.Path.Combine(PicPath, String.Format("~GRAB_{0:00000}.png", imgNum))
            Do
               Threading.Thread.Sleep(1000)
               If IO.File.Exists(fn) Then
                  Exit Do
               End If
            Loop

            Dim bm As Drawing.Bitmap = New Drawing.Bitmap(fn)

            Return bm

         Else
            Return Nothing
         End If


      Else

         Dim PicPath As String = ServerPaths.DataPath

         If PicPath.EndsWith("\") Then
            PicPath = PicPath.Substring(0, PicPath.Length - 1)
         End If

         If Not PicPath.Substring(2, 1) = ":" Then
            Throw New Exception("The Grab-Function needs to have absolute paths configured in CasparCG.config.")
            Return Nothing
            Exit Function
         End If

         Dim fn() As String
         fn = IO.Directory.GetFiles(PicPath, "*.png")
         For c As Integer = 0 To fn.Length - 1
            IO.File.Delete(fn(c))
         Next

         Threading.Thread.Sleep(1000)

         Dim ri As ReturnInfo
         If layer = -1 Then
            ri = Execute(String.Format("PRINT {0}", channel))
         Else
            ri = Execute(String.Format("PRINT {0}-{1}", channel, layer))
         End If

         If ri.Number = 202 Then

            Do
               Threading.Thread.Sleep(1000)
               fn = IO.Directory.GetFiles(PicPath, "*.png")
               If fn.Length > 0 Then
                  Exit Do
               End If
            Loop

            Dim bm As Drawing.Bitmap = New Drawing.Bitmap(fn(0))

            Return bm

         Else
            Return Nothing
         End If
      End If

   End Function

   ''' <summary>
   ''' Simple Grab function also for remote servers
   ''' </summary>
   ''' <param name="Filename">The filename to store the grabbed image</param>
   ''' <param name="channel">The channel in Caspar</param>
   ''' <param name="layer">The layer number</param>
   Public Sub SimpleGrab(ByVal Filename As String, ByVal channel As Integer, ByVal layer As Integer)

      Execute(String.Format("ADD {0}-{1} IMAGE {2}", channel, layer, Filename))

      Thread.Sleep(1000)

      Execute(String.Format("REMOVE {0}-{1} IMAGE", channel, layer))

   End Sub

#End Region

#Region "Info and query commands"

   ''' <summary>
   ''' Check if something is playing on the default layer and channel
   ''' </summary>
   ''' <returns></returns>
   Public Function IsPlaying() As Boolean
      Return IsPlaying(DefaultChannel, DefaultLayer)
   End Function

   ''' <summary>
   ''' Check if something is playing on the given layer on the default channel
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <returns></returns>
   Public Function IsPlaying(ByVal Layer As Integer) As Boolean
      Return IsPlaying(DefaultChannel, Layer)
   End Function

   ''' <summary>
   ''' Check if something is playing on the given layer on the given channel
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <returns></returns>
   Public Function IsPlaying(ByVal Channel As Integer, ByVal Layer As Integer) As Boolean

      Dim ret As Boolean = False

      Try
         Dim s As String = Execute(String.Format("INFO {0}-{1}", Channel, Layer)).Data
         Dim doc As XmlDocument = New XmlDocument()
         doc.LoadXml(Left(s, InStrRev(s, ">")))
         Dim ndLayer As XmlNode = doc.ChildNodes(1)
         For Each nd As XmlNode In ndLayer.ChildNodes
            If nd.Name = "frame-age" Then
               ret = (nd.InnerText <> "0")
               Exit For
            End If
         Next
      Catch ex As Exception
         ret = True
      End Try

      Return ret

   End Function

   Public Function GetTemplateParameters(ByVal TemplateName As String) As List(Of TemplateParameter)

      Dim lst As List(Of TemplateParameter) = New List(Of TemplateParameter)

      Try

         Dim s As String = Execute(String.Format("INFO TEMPLATE {0}", TemplateName)).Data

         Dim doc As XmlDocument = New XmlDocument()
         doc.LoadXml(Left(s, s.IndexOf("</template>") + 11))

         Dim nl As XmlNodeList = doc.SelectNodes("/template/parameters/parameter")

         For Each nd As XmlNode In nl
            Dim tp As TemplateParameter = New TemplateParameter(nd.Attributes("id").InnerText, nd.Attributes("type").InnerText, nd.Attributes("info").InnerText)
            lst.Add(tp)
         Next

      Catch ex As Exception
         Debug.Print(ex.Message)
      End Try

      Return lst

   End Function


   ''' <summary>
   ''' Get all media clip names
   ''' </summary>
   ''' <returns>List of media clip names</returns>
   Public Function GetMediaClipsNames() As List(Of String)
      Return GetMediaClipsNames("", MediaTypes.All, False)
   End Function

   ''' <summary>
   ''' Get specified media clip names
   ''' </summary>
   ''' <param name="ShowMedia">Media clip types</param>
   ''' <returns>List of media clip names</returns>
   Public Function GetMediaClipsNames(ByVal ShowMedia As MediaTypes) As List(Of String)
      Return GetMediaClipsNames("", ShowMedia, False)
   End Function

   ''' <summary>
   ''' Get specified media clip names
   ''' </summary>
   ''' <param name="ShowMedia">Media clip types</param>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <returns>List of media clip names</returns>
   Public Function GetMediaClipsNames(ByVal ShowMedia As MediaTypes, ByVal AddEmptyEntry As Boolean) As List(Of String)
      Return GetMediaClipsNames("", ShowMedia, AddEmptyEntry, True)
   End Function

   ''' <summary>
   ''' Get specified media clip names
   ''' </summary>
   ''' <param name="FolderName">The name of the subfolder, that is used to list the entries</param>
   ''' <param name="ShowMedia">Media clip types</param>
   ''' <returns>List of media clip names</returns>
   Public Function GetMediaClipsNames(ByVal FolderName As String, ByVal ShowMedia As MediaTypes) As List(Of String)
      Return GetMediaClipsNames(FolderName, ShowMedia, False, True)
   End Function

   ''' <summary>
   ''' Get specified media clip names
   ''' </summary>
   ''' <param name="FolderName">The name of the subfolder, that is used to list the entries</param>
   ''' <param name="ShowMedia">Media clip types</param>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <returns>List of media clip names</returns>
   Public Function GetMediaClipsNames(ByVal FolderName As String, ByVal ShowMedia As MediaTypes, ByVal AddEmptyEntry As Boolean) As List(Of String)
      Return GetMediaClipsNames(FolderName, ShowMedia, False, True)
   End Function

   ''' <summary>
   ''' Get specified media clip names
   ''' </summary>
   ''' <param name="FolderName">The name of the subfolder, that is used to list the entries</param>
   ''' <param name="ShowMedia">Media clip types</param>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <param name="SuppressFolderName">If True only the name of the template is returned, if False the foldername and the templatename is returned</param>
   ''' <returns></returns>
   Public Function GetMediaClipsNames(ByVal FolderName As String, ByVal ShowMedia As MediaTypes, ByVal AddEmptyEntry As Boolean, ByVal SuppressFolderName As Boolean) As List(Of String)

      Dim lst As List(Of String) = New List(Of String)
      If AddEmptyEntry Then
         lst.Add("")
      End If

      Dim dat As Object = Execute("CLS", True).Data
      If dat IsNot Nothing Then

         Dim Ret() As String = CType(dat, String).Replace(vbLf, "").Split(CChar(vbCr))
         Dim c As Integer = 1
         Dim d As Integer = 0

         Thread.Sleep(200)

         FolderName = FolderName.ToUpper

         For Each s As String In Ret
            If s.Contains(Chr(&H22)) Then

               c = s.IndexOf(Chr(&H22))
               If c > -1 Then

                  d = s.IndexOf(Chr(&H22), c + 1)
                  If d > -1 Then

                     Dim clip As String = s.Substring(c + 1, d - (c + 1)).Replace("\", "/").Trim

                     If ShowMedia = MediaTypes.All Then

                        If FolderName <> "" Then
                           If FolderName = "_ROOT" Then
                              If Not clip.Contains("/") Then
                                 lst.Add(clip)
                              End If
                           Else
                              If clip.StartsWith(FolderName) Then
                                 If SuppressFolderName Then
                                    lst.Add(clip.Substring(FolderName.Length + 1))
                                 Else
                                    lst.Add(clip)
                                 End If
                              End If
                           End If
                        Else
                           lst.Add(clip)
                        End If

                     Else
                        If s.Length >= (d + 8) Then

                           If ShowMedia = MediaTypes.Movie And s.Substring(d + 3, 5) = "MOVIE" Then
                              If FolderName <> "" Then
                                 If FolderName = "_ROOT" Then
                                    If Not clip.Contains("/") Then
                                       lst.Add(clip)
                                    End If
                                 Else
                                    If clip.StartsWith(FolderName) Then
                                       If SuppressFolderName Then
                                          lst.Add(clip.Substring(FolderName.Length + 1))
                                       Else
                                          lst.Add(clip)
                                       End If
                                    End If
                                 End If
                              Else
                                 lst.Add(clip)
                              End If
                           End If

                           If ShowMedia = MediaTypes.Still And s.Substring(d + 3, 5) = "STILL" Then
                              If FolderName <> "" Then
                                 If FolderName = "_ROOT" Then
                                    If Not clip.Contains("/") Then
                                       lst.Add(clip)
                                    End If
                                 Else
                                    If clip.StartsWith(FolderName) Then
                                       If SuppressFolderName Then
                                          lst.Add(clip.Substring(FolderName.Length + 1))
                                       Else
                                          lst.Add(clip)
                                       End If
                                    End If
                                 End If
                              Else
                                 lst.Add(clip)
                              End If
                           End If

                           If ShowMedia = MediaTypes.Audio And s.Substring(d + 3, 5) = "AUDIO" Then
                              If FolderName <> "" Then
                                 If FolderName = "_ROOT" Then
                                    If Not clip.Contains("/") Then
                                       lst.Add(clip)
                                    End If
                                 Else
                                    If clip.StartsWith(FolderName) Then
                                       If SuppressFolderName Then
                                          lst.Add(clip.Substring(FolderName.Length + 1))
                                       Else
                                          lst.Add(clip)
                                       End If
                                    End If
                                 End If
                              Else
                                 lst.Add(clip)
                              End If
                           End If

                        End If
                     End If
                  End If
               End If

            End If
         Next
      End If

      If lst.Count = 0 Then
         lst.Add(" ")
      End If

      Return lst

   End Function

   ''' <summary>
   ''' Get a list of media folder names
   ''' </summary>
   ''' <returns></returns>
   Public Function GetMediaFolderNames() As List(Of String)
      Return GetMediaFolderNames(False)
   End Function

   ''' <summary>
   ''' Get a list of media folder names
   ''' </summary>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <returns></returns>
   Public Function GetMediaFolderNames(ByVal AddEmptyEntry As Boolean) As List(Of String)

      Dim lst As List(Of String) = New List(Of String)
      If AddEmptyEntry Then
         lst.Add("")
      End If

      lst.Add("_ROOT")

      Dim clips As List(Of String) = GetMediaClipsNames()
      Dim hash As HashSet(Of String) = New HashSet(Of String)

      Dim folder As String = ""
      Dim ind As Integer = -1
      For Each cn As String In clips

         ind = cn.LastIndexOf("/")
         If ind > -1 Then

            folder = cn.Substring(0, ind)

            If Not hash.Contains(folder) Then
               hash.Add(folder)
               lst.Add(folder)
            End If

         End If

      Next

      Return lst

   End Function


   ''' <summary>
   ''' Get a list of template names
   ''' </summary>
   ''' <returns>List of template names</returns>
   Public Function GetTemplateNames() As List(Of String)
      Return GetTemplateNames("", False)
   End Function

   ''' <summary>
   ''' Get a list of template names
   ''' </summary>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <returns>List of template names</returns>
   Public Function GetTemplateNames(ByVal AddEmptyEntry As Boolean) As List(Of String)
      Return GetTemplateNames("", AddEmptyEntry)
   End Function

   ''' <summary>
   ''' Get a list of template names
   ''' </summary>
   ''' <param name="FolderName">The name of the subfolder, that is used to list the entries</param>
   ''' <returns>List of template names</returns>
   Public Function GetTemplateNames(ByVal FolderName As String) As List(Of String)
      Return GetTemplateNames(FolderName, False)
   End Function

   ''' <summary>
   ''' Get a list of template names
   ''' </summary>
   ''' <param name="FolderName">The name of the subfolder, that is used to list the entries</param>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <returns>List of template names</returns>
   Public Function GetTemplateNames(ByVal FolderName As String, ByVal AddEmptyEntry As Boolean) As List(Of String)
      Return GetTemplateNames(FolderName, AddEmptyEntry, True)
   End Function

   ''' <summary>
   ''' Get a list of template names
   ''' </summary>
   ''' <param name="FolderName">The name of the subfolder, that is used to list the entries</param>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <param name="SuppressFolderName">If True only the name of the template is returned, if False the foldername and the templatename is returned</param>
   ''' <returns>List of template names</returns>
   Public Function GetTemplateNames(ByVal FolderName As String, ByVal AddEmptyEntry As Boolean, ByVal SuppressFolderName As Boolean) As List(Of String)

      Dim lst As List(Of String) = New List(Of String)
      If AddEmptyEntry Then
         lst.Add("")
      End If

      Try

         Dim ri As ReturnInfo = Execute("TLS", True)
         If ri IsNot Nothing AndAlso ri.Data IsNot Nothing Then

            Dim Ret() As String = ri.Data.Replace(vbLf, "").Split(CChar(vbCr))
            Dim i As Integer = 0

            If Me.Version.Major >= 2 And Me.Version.Minor >= 2 Then

               For Each s As String In Ret

                  If FolderName <> "" Then
                     If FolderName = "_ROOT" Then
                        If Not s.Contains("/") Then
                           lst.Add(s)
                        End If
                     Else
                        If s.StartsWith(FolderName) Then
                           If SuppressFolderName Then
                              lst.Add(s.Substring(FolderName.Length + 1))
                           Else
                              lst.Add(s)
                           End If
                        End If
                     End If
                  Else
                     lst.Add(s)
                  End If

               Next

            Else

               For Each s As String In Ret
                  If s.Contains(Chr(&H22)) Then
                     i = s.IndexOf(Chr(&H22), 2)
                     If i > -1 Then
                        Dim tmpl As String = s.Substring(1, i - 1).Replace("\", "/").Replace(Chr(&H22), "")

                        If FolderName <> "" Then
                           If FolderName = "_ROOT" Then
                              If Not tmpl.Contains("/") Then
                                 lst.Add(tmpl)
                              End If
                           Else
                              If tmpl.StartsWith(FolderName) Then
                                 If SuppressFolderName Then
                                    lst.Add(tmpl.Substring(FolderName.Length + 1))
                                 Else
                                    lst.Add(tmpl)
                                 End If
                              End If
                           End If
                        Else
                           lst.Add(tmpl)
                        End If

                     End If
                  Else
                     Exit For
                  End If
               Next

            End If

         End If

      Catch ex As Exception
         'Ignore
      End Try

      If lst.Count = 0 Then
         lst.Add(" ")
      End If

      Return lst

   End Function

   ''' <summary>
   ''' Get a list of template folder names
   ''' </summary>
   ''' <returns>List of folder names</returns>
   Public Function GetTemplateFolderNames() As List(Of String)
      Return GetTemplateFolderNames(False)
   End Function

   ''' <summary>
   ''' Get a list of template folder names
   ''' </summary>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <returns>List of folder names</returns>
   Public Function GetTemplateFolderNames(ByVal AddEmptyEntry As Boolean) As List(Of String)

      Dim lst As List(Of String) = New List(Of String)
      If AddEmptyEntry Then
         lst.Add("")
      End If

      lst.Add("_ROOT")

      Dim templates As List(Of String) = GetTemplateNames()
      Dim hash As HashSet(Of String) = New HashSet(Of String)

      Dim folder As String = ""
      Dim ind As Integer = -1
      For Each tn As String In templates

         ind = tn.LastIndexOf("/")
         If ind > -1 Then

            folder = tn.Substring(0, ind)

            If Not hash.Contains(folder) Then
               hash.Add(folder)
               lst.Add(folder)
            End If

         End If

      Next

      Return lst

   End Function

   ''' <summary>
   ''' Get a list of template folder names, that contains also the sub paths.
   ''' </summary>
   ''' <returns>List of folder names</returns>
   Public Function GetAllTemplateFolderNames() As List(Of String)
      Return GetAllTemplateFolderNames(False)
   End Function

   ''' <summary>
   ''' Get a list of template folder names, that contains also the sub paths.
   ''' </summary>
   ''' <param name="AddEmptyEntry">True if the first entry is empty. Usefull to be able to set a selection (combobox) to nothing.</param>
   ''' <returns>List of folder names</returns>
   Public Function GetAllTemplateFolderNames(ByVal AddEmptyEntry As Boolean) As List(Of String)

      Dim fld As List(Of String) = GetTemplateFolderNames(AddEmptyEntry)
      Dim hash As HashSet(Of String) = New HashSet(Of String)
      Dim parts() As String
      Dim folder As String

      For Each pat As String In fld

         folder = ""
         parts = pat.Split(CChar("/"))

         For Each s As String In parts
            folder += s + "/"
            If Not hash.Contains(folder.Substring(0, folder.Length - 1)) Then
               hash.Add(folder.Substring(0, folder.Length - 1))
            End If
         Next

      Next

      Return hash.ToList

   End Function

   ''' <summary>
   ''' Get a Template object by providing a name
   ''' </summary>
   ''' <param name="Name">The name of the template</param>
   ''' <returns>A Template object parsed out of informations queried by Caspar</returns>
   ''' <remarks></remarks>
   Public Function GetTemplate(ByVal Name As String) As Template

      Dim info As String = Execute(String.Format("INFO TEMPLATE {0}", Name)).Data
      Dim tmpl As Template = Nothing

      If info IsNot Nothing Then
         tmpl = Template.Parse(info)
      Else
         tmpl = New Template
      End If

      tmpl.Name = Name
      Return tmpl

   End Function

   ''' <summary>
   ''' Get a list of Template objects
   ''' </summary>
   ''' <returns>List of Template objects parsed out of informations queried by Caspar</returns>
   Public Function GetTemplates() As List(Of Template)

      Dim tl As List(Of Template) = New List(Of Template)
      Dim lst As List(Of String) = Me.GetTemplateNames

      For Each s As String In lst
         If s <> "" Then
            tl.Add(Me.GetTemplate(s))
         End If
      Next

      Return tl

   End Function

   Public Function GetNumberOfChannels() As Integer

      Dim ret As ReturnInfo = Execute("INFO SERVER")
      Dim count As Integer = 0

      Dim pos As Integer = -1
      Do
         pos = ret.Data.IndexOf("<channel>", pos + 1)
         If pos > 0 Then
            count += 1
         Else
            Exit Do
         End If
      Loop

      Return count

   End Function

   ''' <summary>
   ''' LayerStatus object
   ''' </summary>
   ''' <remarks>Used for GetLayerStatus</remarks>
   Public Class LayerStatus

      ''' <summary>
      ''' The name of the file playing
      ''' </summary>
      Public Property FilePlaying As String
      ''' <summary>
      ''' The layer status
      ''' </summary>
      Public Property Status As String
      ''' <summary>
      ''' The current position in frames
      ''' </summary>
      Public Property FrameNumber As Long
      ''' <summary>
      ''' The length of the clip in frames
      ''' </summary>
      Public Property TotalFrames As Long
      ''' <summary>
      ''' Determines if a file has been loaded for seamless playback
      ''' </summary>
      Public Property BackGroundIsEmpty As Boolean = False
   End Class

   ''' <summary>
   ''' Get the status of a layer
   ''' </summary>
   ''' <param name="Layer">The layer number</param>
   ''' <returns>A LayerStatus object</returns>
   Public Function GetLayerStatus(ByVal Layer As Integer) As LayerStatus
      Return GetLayerStatus(DefaultChannel, Layer)
   End Function

   ''' <summary>
   ''' Get the status of a layer
   ''' </summary>
   ''' <param name="Channel">The channel in Caspar</param>
   ''' <param name="Layer">The layer number</param>
   ''' <returns>A LayerStatus object</returns>
   Public Function GetLayerStatus(ByVal Channel As Integer, ByVal Layer As Integer) As LayerStatus

      Dim doc As XmlDocument = New XmlDocument()
      Dim ls As LayerStatus = New LayerStatus()

      Dim ri As ReturnInfo = Execute(String.Format("INFO {0}-{1}", Channel, Layer))
      Try
         doc.LoadXml(ri.Data.Substring(0, ri.Data.LastIndexOf("</layer>") + 8))

         Dim nl As XmlNodeList = doc.SelectNodes("/layer/foreground/producer/filename")
         If nl.Count > 0 Then
            ls.FilePlaying = nl(0).FirstChild.Value
         End If

         nl = doc.SelectNodes("/layer/status")
         ls.Status = nl(0).FirstChild.Value

         nl = doc.SelectNodes("/layer/frame-number")
         ls.FrameNumber = Long.Parse(nl(0).FirstChild.Value)

         nl = doc.SelectNodes("/layer/nb_frames")
         ls.TotalFrames = Long.Parse(nl(0).FirstChild.Value)

         nl = doc.SelectNodes("/layer/background/destination/producer/type")
         If nl.Count = 0 Then
            ls.BackGroundIsEmpty = True
         Else
            ls.BackGroundIsEmpty = (nl(0).FirstChild.Value = "empty-producer")
         End If

      Catch ex As Exception
      End Try

      Return ls

   End Function

#End Region

#Region "Other Methods, functions and shared"

   Private Function AssembleReturnInfo(ByVal s As String) As ReturnInfo

      Dim ri As ReturnInfo = New ReturnInfo()

      If IsNumeric(Left(s, 3)) Then
         ri.Number = Integer.Parse(Left(s, 3))
         Dim c As Integer = s.IndexOf(Microsoft.VisualBasic.vbCrLf)
         If c > 2 Then
            ri.Message = s.Substring(4, c - 2).Trim
            Dim d As Integer = s.IndexOf(Microsoft.VisualBasic.vbCrLf, c + 1)
            If d > 0 Then
               ri.Data = s.Substring(c + 2).Trim
               'ri.Data = s.Substring(c + 2, d - c - 2).Trim
            End If
         Else
            ri.Message = "No usefull information returned"
         End If
         Return ri
      Else
         ri.Number = 0
         ri.Message = "Data returned"
         ri.Data = s
      End If

      Return ri
   End Function

   ''' <summary>
   ''' Sends a command to Caspar for execution
   ''' </summary>
   ''' <param name="Command">The command-string</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Execute(ByVal Command As String) As ReturnInfo
      Return Execute(Command, New Retard, False)
   End Function

   ''' <summary>
   ''' Sends a command to Caspar for execution
   ''' </summary>
   ''' <param name="Command">The command-string</param>
   ''' <param name="useDelay">Use a delay while waiting for a server response (list commands)</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Execute(ByVal Command As String, ByVal useDelay As Boolean) As ReturnInfo
      Return Execute(Command, New Retard, useDelay)
   End Function

   ''' <summary>
   ''' Sends a command to Caspar for execution
   ''' </summary>
   ''' <param name="Command">The command-string</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <returns>A ReturnInfo object</returns>
   Public Function Execute(ByVal Command As String, ByVal Retard As Retard) As ReturnInfo
      Return Execute(Command, Retard, False)
   End Function

   ''' <summary>
   ''' Sends a command to Caspar for execution
   ''' </summary>
   ''' <param name="Command">The command-string</param>
   ''' <param name="Retard">Delay the execution for this amount of milliseconds</param>
   ''' <param name="useDelay">Use a delay while waiting for a server response (list commands)</param>
   ''' <returns>A ReturnInfo object</returns>
   ''' <remarks>Can be used for any command. See CasparCG's wiki for valid commands</remarks>
   Public Function Execute(ByVal Command As String, ByVal Retard As Retard, ByVal useDelay As Boolean) As ReturnInfo

      If _Caspar.Connected Then

         'Debug.Print(Command)

         If Retard.Amount > 0 Then
            _DelayQueue.Add(Retard.Amount, Command)
         Else

            Dim cmd As Byte() = Encoding.UTF8.GetBytes(Command + Microsoft.VisualBasic.vbCrLf)
            Dim bytes(128) As Byte
            Dim s As String = ""
            Try
               ' Blocks until send returns.
               Dim i As Integer = _Caspar.Send(cmd)

               If useDelay Then
                  Thread.Sleep(250)
               End If

               ' Get reply from the server.
               Do
                  i = _Caspar.Receive(bytes, bytes.Length, SocketFlags.None)
                  s += Encoding.UTF8.GetString(bytes)
                  If i < bytes.Length Then Exit Do
                  'If s.Length > 64000 Then Exit Do 'ToDo: Fix this
               Loop

            Catch ex As SocketException
               If s.Length > 0 Then
                  Return AssembleReturnInfo(s)
               Else
                  Return New ReturnInfo(0, String.Format("{0} Error code: {1}.", ex.Message, ex.ErrorCode), "")
               End If
            End Try

            Return AssembleReturnInfo(s)

         End If

      Else

         _Caspar = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
         Connect()
         If _Caspar.Connected Then
            Return Execute(Command)
         Else
            Return New ReturnInfo(0, "Not connected to CasparCG", "")
         End If

      End If

      Return New ReturnInfo(0, "Still not connected to CasparCG", "")

   End Function

   ''' <summary>
   ''' Connects to CasparCG
   ''' </summary>
   ''' <param name="serverAdress">IP-Address or computer name to connect to Caspar</param>
   ''' <remarks>Sets the ServerAdress property</remarks>
   ''' <returns>A enumConnectResult denoting the result</returns>
   Public Function Connect(ByVal serverAdress As String) As enumConnectResult
      _ServerAdress = serverAdress
      Return Me.Connect()
   End Function

   ''' <summary>
   ''' Connects to CasparCG
   ''' </summary>
   ''' <param name="serverAdress">IP-Address or computer name to connect to Caspar</param>
   ''' <param name="Port">The port number to connect to</param>
   ''' <remarks>Sets the ServerAdress and Port properties</remarks>
   ''' <returns>A enumConnectResult denoting the result</returns>
   Public Function Connect(ByVal serverAdress As String, ByVal Port As Integer) As enumConnectResult
      _ServerAdress = serverAdress
      _Port = Port
      Return Me.Connect()
   End Function

   ''' <summary>
   ''' Connects to CasparCG
   ''' </summary>
   ''' <param name="ServerAdress">IP-Address or computer name to connect to Caspar</param>
   ''' <param name="KeepQuiet">Inhibits exeptions on connection errors</param>
   ''' <remarks>Sets the ServerAdress and KeepQuiet properties</remarks>
   ''' <returns>A enumConnectResult denoting the result</returns>
   Public Function Connect(ByVal ServerAdress As String, ByVal KeepQuiet As Boolean) As enumConnectResult
      _ServerAdress = ServerAdress
      _KeepQuiet = KeepQuiet
      Return Me.Connect()
   End Function

   ''' <summary>
   ''' Connects to CasparCG
   ''' </summary>
   ''' <param name="Retries">Number of retries for a connection, before giving up.</param>
   ''' <param name="CasparExePath">Full qualified path to Caspar's exe file. Must be local to the client.</param>
   ''' <remarks>Sets the Retries and CasparExePath properties</remarks>
   ''' <returns>A enumConnectResult denoting the result</returns>
   Public Function Connect(ByVal Retries As Integer, ByVal CasparExePath As String) As enumConnectResult
      If Me.ServerAdress.ToLower = "localhost" Or Me.ServerAdress = "127.0.0.1" Then
         _CasparExePath = CasparExePath
         _Retries = Retries
         Return Me.Connect()
      Else
         Return enumConnectResult.crIsNotLocal
      End If
   End Function

   ''' <summary>
   ''' Connects to CasparCG
   ''' </summary>
   ''' <remarks>The main Connect function. Uses properties</remarks>
   ''' <returns>A enumConnectResult denoting the result</returns>
   Public Function Connect() As enumConnectResult

      Dim retVal As enumConnectResult = enumConnectResult.crSuccessfull

      Dim ping As Ping = New Ping

      Dim pr As PingReply = Nothing
      Try
         pr = ping.Send(_ServerAdress, 1500)
      Catch ex As Exception
         'handled futher down
      End Try

      If pr.Status = IPStatus.Success Then

         Try
            _Caspar.Connect(_ServerAdress, _Port)

         Catch ex As Exception

            If _Retries > 5 And IO.File.Exists(_CasparExePath) Then

               _CasparProc = New Process
               _CasparProc.StartInfo.FileName = _CasparExePath
               _CasparProc.StartInfo.WorkingDirectory = IO.Path.GetDirectoryName(_CasparExePath)
               _CasparProc.StartInfo.CreateNoWindow = False
               _CasparProc.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
               _CasparProc.Start()

               Dim t As Integer = 0
               Do
                  If Not _Caspar.Connected Then

                     Try
                        _Caspar.Connect(_ServerAdress, _Port)
                     Catch exp As Exception
                        Debug.Print(exp.Message)
                     End Try

                  End If

                  If _Caspar.Connected Then

                     'Check for Scanner.exe
                     Dim scannerFilename As String = IO.Path.Combine(IO.Path.GetDirectoryName(_CasparExePath), "Scanner.exe")
                     If IO.File.Exists(scannerFilename) Then

                        _Scanner = New Process
                        _Scanner.StartInfo.FileName = IO.Path.Combine(IO.Path.GetDirectoryName(_CasparExePath), "Scanner.exe")
                        _Scanner.StartInfo.WorkingDirectory = IO.Path.GetDirectoryName(_CasparExePath)
                        _Scanner.StartInfo.CreateNoWindow = False
                        _Scanner.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                        _Scanner.Start()

                     End If

                     Exit Do

                  Else

                     If t > _Retries Then
                        retVal = enumConnectResult.crLocalCasparCGCouldNotBeStarted
                        Exit Do
                     Else
                        System.Threading.Thread.Sleep(1000)
                        t += 1
                     End If

                  End If

               Loop

            Else
               retVal = enumConnectResult.crCasparCGNotStarted
            End If

         End Try

      Else
         retVal = enumConnectResult.crMachineNotAvailable
      End If

      Return retVal

   End Function

   ''' <summary>
   ''' Disconnects from Caspar
   ''' </summary>
   ''' <remarks>The object can reconnect on version 2.0.7</remarks>
   Public Sub Disconnect()
      If _Caspar.Connected Then
         Try
            _Caspar.Disconnect(False)
            _Caspar = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
         Catch ex As Exception
            'ignore
         End Try
      End If
   End Sub

   Public Sub Shutdown()

      If _Scanner IsNot Nothing Then
         Try
            _Scanner.Kill()
         Catch ex As Exception
         Finally
            _Scanner.Dispose()
         End Try
      End If

      If _CasparProc IsNot Nothing Then
         Try
            _CasparProc.Kill()
         Catch ex As Exception
         Finally
            _CasparProc.Dispose()
         End Try
      Else
         If Version.Major = 2 And Version.Minor = 0 Then
            Execute("KILL")
         End If
      End If

   End Sub

   Public Sub Restart()
      If IO.File.Exists(CasparExePath) Then

         If _Scanner IsNot Nothing Then
            Try
               _Scanner.Kill()
            Catch ex As Exception
            Finally
               _Scanner.Dispose()
            End Try
         End If

         If _CasparProc IsNot Nothing Then
            Try
               _CasparProc.Kill()
            Catch ex As Exception
            Finally
               _CasparProc.Dispose()
            End Try
         Else
            If Version.Major = 2 And Version.Minor = 0 Then
               Execute("KILL")
            End If
         End If

         _Caspar = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
         Threading.Thread.Sleep(2000)
         Connect()

      End If
   End Sub

   Public Sub SynchPictureFolder(ByVal ClientsPictureFolder As String)
      CopyAll(New DirectoryInfo(ClientsPictureFolder), New DirectoryInfo(Me.RemotePictureFolder))
   End Sub

   Private Sub CopyAll(ByVal source As DirectoryInfo, ByVal target As DirectoryInfo)

      Try

         If Not Directory.Exists(target.FullName) Then
            Directory.CreateDirectory(target.FullName)
         End If

         For Each fi As FileInfo In source.GetFiles()
            fi.CopyTo(Path.Combine(target.ToString(), fi.Name), True)
         Next

         For Each diSourceDir As DirectoryInfo In source.GetDirectories()
            Dim nextTargetDir As DirectoryInfo = target.CreateSubdirectory(diSourceDir.Name)
            CopyAll(diSourceDir, nextTargetDir)
         Next

      Catch ex As Exception
         MsgBox(ex.Message)
      End Try

   End Sub

   'Public Function PicNameToLocalPicNameAsURI(SourceFilename As String) As String
   '   Return PicNameToLocalPicNameAsURI(SourceFilename, ClientsPictureFolder, Me.LocalPictureFolder, UsePathAdapter)
   'End Function

   Public Function PicNameToLocalPicNameAsURI(ByVal SourceFilename As String, ByVal ClientsPictureFolder As String, ByVal UsePathAdapter As Boolean) As String
      Return PicNameToLocalPicNameAsURI(SourceFilename, ClientsPictureFolder, Me.LocalPictureFolder, UsePathAdapter)
   End Function

   Public Shared Function PicNameToLocalPicNameAsURI(ByVal SourceFilename As String, ByVal ClientsPictureFolder As String, ByVal CasparsLocalPictureFolder As String, ByVal UsePathAdapter As Boolean) As String
      If SourceFilename <> "" Then
         If UsePathAdapter Then
            Dim subPath As String = SourceFilename.Substring(ClientsPictureFolder.Length)
            If subPath.StartsWith("\") Then
               subPath = subPath.Substring(1)
            End If
            Return New System.Uri(Path.Combine(CasparsLocalPictureFolder, subPath)).ToString
         Else
            Return New System.Uri(SourceFilename).ToString
         End If
      Else
         Return ""
      End If
   End Function

   ''' <summary>
   ''' Serialize the whole CasparCG object to a string.
   ''' </summary>
   ''' <returns>A XML formated string</returns>
   ''' <remarks>Used to store the settings</remarks>
   Public Function SerializeToString() As String

      Dim ser As XmlSerializer = New XmlSerializer(GetType(CasparCG))
      Dim sw As IO.StringWriter = New IO.StringWriter()
      ser.Serialize(sw, Me)

      Return sw.ToString()

   End Function

   Public Overrides Function ToString() As String
      Return String.Format("{0} Channel {1}", Name, DefaultChannel)
   End Function

   ''' <summary>
   ''' Checks if the current threads country uses NTSC framerate 29.97, 30, 59.94 or 60fps
   ''' </summary>
   ''' <returns>True/False</returns>
   Public Shared Function CurrentCountryIsNTSC() As Boolean

      'Contains a list of NTSC country two letter ISO codes
      Dim iso As String() = {"AN", "BS", "BB", "BM", "BO", "CA", "CL", "CO", "CR", "CU", "CW", "DO", "EC", "SV", "GL", "GU", "GT", "HN", "JM", "JP", "KR", "MX", "NI", "PA", "PE", "PH", "PR", "LK", "SR", "TW", "TT", "US", "VE", "VN", "VI"}
      Dim isoHash As HashSet(Of String) = New HashSet(Of String)(iso)
      Dim ci As CultureInfo = Thread.CurrentThread.CurrentCulture
      Dim ri As RegionInfo = New RegionInfo(ci.LCID)

      Return isoHash.Contains(ri.TwoLetterISORegionName.ToUpper)

   End Function

   ''' <summary>
   ''' Write a caspar.config file based on a template
   ''' </summary>
   ''' <param name="TemplateXMLFilename">Template file to use. {paths} will be replaced by the paths to the assets, {videomode} will be replased with the 1080i and the current framerate.</param>
   ''' <param name="CasparExePath">Path to the CasparCG exe file</param>
   ''' <param name="AssetsPath">Base path for all assets. If the folders do not exist, they will be created.</param>
   ''' <param name="modePAL">Optional, videomode string for PAL countries</param>
   ''' <param name="modeNTSC">Optional, videomode string for NTSC countries</param>
   Public Shared Sub WriteCasparConfig(ByVal TemplateXMLFilename As String, ByVal CasparExePath As String, ByVal AssetsPath As String, Optional ByVal modePAL As String = "1080i5000", Optional ByVal modeNTSC As String = "1080i5994")

      If Not IO.File.Exists(Path.Combine(CasparExePath, "casparcg.config")) Then

         If IO.File.Exists(TemplateXMLFilename) AndAlso IO.Directory.Exists(CasparExePath) Then

            If Not IO.Directory.Exists(AssetsPath) Then

               Dim parts() As String = AssetsPath.Split("\".ToCharArray)
               Dim folder As String = ""

               For Each prt As String In parts
                  folder += "\" + prt
                  If Not IO.Directory.Exists(folder) Then
                     Try
                        IO.Directory.CreateDirectory(folder)
                     Catch ex As Exception
                        Exit Sub
                     End Try
                  End If
               Next

            End If

            Dim tmpPath As String = IO.Path.Combine(AssetsPath, "media")
            If Not IO.Directory.Exists(tmpPath) Then
               IO.Directory.CreateDirectory(tmpPath)
            End If

            tmpPath = IO.Path.Combine(AssetsPath, "log")
            If Not IO.Directory.Exists(tmpPath) Then
               IO.Directory.CreateDirectory(tmpPath)
            End If

            tmpPath = IO.Path.Combine(AssetsPath, "data")
            If Not IO.Directory.Exists(tmpPath) Then
               IO.Directory.CreateDirectory(tmpPath)
            End If

            tmpPath = IO.Path.Combine(AssetsPath, "templates")
            If Not IO.Directory.Exists(tmpPath) Then
               IO.Directory.CreateDirectory(tmpPath)
            End If

            tmpPath = IO.Path.Combine(AssetsPath, "thumbnails")
            If Not IO.Directory.Exists(tmpPath) Then
               IO.Directory.CreateDirectory(tmpPath)
            End If

            Using reader As StreamReader = New StreamReader(TemplateXMLFilename)
               Using writer As StreamWriter = New StreamWriter(IO.Path.Combine(CasparExePath, "casparcg.config"))

                  Dim line As String
                  Dim head As String
                  Do
                     If reader.EndOfStream Then Exit Do

                     line = reader.ReadLine()
                     If line.Trim.ToLower = "{paths}" Then

                        head = line.ToLower.Replace("{paths}", "")

                        writer.WriteLine(String.Format("{1}<media-path>{0}\media\</media-path>", AssetsPath, head))
                        writer.WriteLine(String.Format("{1}<log-path>{0}\log\</log-path>", AssetsPath, head))
                        writer.WriteLine(String.Format("{1}<data-path>{0}\data\</data-path>", AssetsPath, head))
                        writer.WriteLine(String.Format("{1}<template-path>{0}\templates\</template-path>", AssetsPath, head))
                        writer.WriteLine(String.Format("{1}<thumbnails-path>{0}\thumbnails\</thumbnails-path>", AssetsPath, head))
                        writer.WriteLine(String.Format("{1}<font-path>{0}</font-path>", System.Environment.GetFolderPath(Environment.SpecialFolder.Fonts), head))

                     ElseIf line.Trim.ToLower = "{videomode}" Then

                        head = line.ToLower.Replace("{videomode}", "")

                        If CurrentCountryIsNTSC() Then
                           writer.WriteLine(String.Format("{1}<video-mode>{0}</video-mode>", modeNTSC, head))
                        Else
                           writer.WriteLine(String.Format("{1}<video-mode>{0}</video-mode>", modePAL, head))
                        End If

                     Else
                        writer.WriteLine(line)
                     End If

                  Loop

                  writer.Flush()
                  writer.Close()

               End Using

               reader.Close()

            End Using

         End If

      End If

   End Sub

#End Region

#Region "Contructors"

   ''' <summary>
   ''' Parameter less constructor
   ''' </summary>
   Public Sub New()

      _Caspar = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
      _Caspar.ReceiveTimeout = 5000

   End Sub

   ''' <summary>
   ''' Deserializing Constructor
   ''' </summary>
   ''' <param name="Serialized">A XML formated string of a CasprCG object</param>
   ''' <remarks>Used to recall the settings</remarks>
   Public Sub New(Serialized As String)

      _Caspar = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
      _Caspar.ReceiveTimeout = 5000

      Dim b As Byte() = Encoding.Unicode.GetBytes(Serialized)
      Dim ser As XmlSerializer = New XmlSerializer(GetType(CasparCG))
      Dim ccg As CasparCG = CType(ser.Deserialize(New System.IO.MemoryStream(b)), CasparCG)

      Me.Name = ccg.Name
      Me.ServerAdress = ccg.ServerAdress
      Me.Port = ccg.Port
      Me.KeepQuiet = ccg.KeepQuiet
      Me.CasparExePath = ccg.CasparExePath
      Me.Retries = ccg.Retries
      Me.AddInfoFields = ccg.AddInfoFields
      Me.LocalPictureFolder = ccg.LocalPictureFolder
      Me.RemotePictureFolder = ccg.RemotePictureFolder

   End Sub

   ''' <summary>
   ''' ICloneable implaementation
   ''' </summary>
   ''' <returns>A clone of this object</returns>
   ''' <remarks>To reuse the socket connection, create a clone of the object</remarks>
   Public Function Clone() As Object Implements System.ICloneable.Clone
      Dim s As String = Me.SerializeToString()
      Return New CasparCG(s)
   End Function

#End Region

End Class
