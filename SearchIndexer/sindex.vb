Imports System.IO
Imports System.Threading
Imports System.IO.Path
Imports System.Xml


'Data Source=(LocalDB)\v11.0;AttachDbFilename="R:\Nelethill\Documents\Visual Studio 2012\Projects\Backup and Match\Backup and Match\SampleDatabase.mdf";Integrated Security=True;Connect Timeout=30
Public Class sindex
    Public path As String
    Dim timer1 As Timer
    Public ErrLog As New List(Of String)
    Dim addin As New addin
    Public sitem As String
    Dim FileList As New List(Of String)
    Dim DriveList As New List(Of String)
    Dim FileIndex As New DataTable
    Dim _gFileName As String
    Public Drive As String = "C:\"
    Public FileCounter As UInteger = 0
    Dim Worker As Thread
    Dim go As Date
    Dim goa As Date

    Private Sub sindex_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DriveList.AddRange(Drive_add)
        Arrayout(DriveList.ToArray)
    End Sub

    Private Function Drive_add() As List(Of String)
        Dim lst As New List(Of String)
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim d As DriveInfo
        Dim dt As String
        For Each d In allDrives
            dt = addin._getDriveType(d.DriveType)
            If dt = "Fixed" Then
                lst.Add(d.Name)
            End If
        Next

        Return lst
    End Function
    Private Sub Arrayout(ByVal out() As String)
        For Each e In out
            ListBox1.Items.Add("[" & DateAndTime.Now & "] " & "Detected Drive: " & e)
        Next
    End Sub
    Private Sub ReadFiles()
        goa = System.DateTime.Now
        Dim list As New List(Of String)
        DriveList.RemoveAt(0)
        For Each i In DriveList
            go = System.DateTime.Now
            SetListBox1("[" & DateAndTime.Now & "] " & "Start Collecting Files on Drive " & i)
            SetProBarStyle(2)
            list.AddRange(GetFilesRecursive(i))
            SetLabel1(list.Count)
            FileCounter = FileCounter + list.Count
            SetListBox1("[" & Now & "] " & "Time To read in " & FileCounter & " Files: " & System.DateTime.Now.Subtract(go).TotalMilliseconds & " ms")
            go = System.DateTime.Now
            SetProBarStyle(1)
            XMLWrite(list, drive & GenFileName(i) & ".xml", FileCounter, i)
            list.Clear()
            Reset_all()
            SetListBox1("[" & DateAndTime.Now & "] " & "Time write XMLFile: " & System.DateTime.Now.Subtract(go).TotalMilliseconds & " ms")
            FileCounter = 0
        Next
        SetListBox1("[" & DateAndTime.Now & "] " & "Time to do the Job: " & System.DateTime.Now.Subtract(go).Minutes & " m")
    End Sub
    Private Sub ReadPath()
        goa = System.DateTime.Now
        'Blocks     ProgressBar1.Style = 0
        'Continuous ProgressBar1.Style = 1
        'Marquee    ProgressBar1.Style = 2
        Dim list As New List(Of String)

        go = System.DateTime.Now
        SetListBox1("[" & DateAndTime.Now & "] " & "Start Collecting Files on Drive " & path)
        SetProBarStyle(2)
        list.AddRange(GetFilesRecursive(path))
        SetLabel1(list.Count)
        FileCounter = FileCounter + list.Count
        SetListBox1("[" & DateAndTime.Now & "] " & "Time To read in " & FileCounter & " Files: " & System.DateTime.Now.Subtract(go).TotalMilliseconds & " ms")
        go = System.DateTime.Now
        SetProBarStyle(1)
        XMLWrite(list, drive & GenFileName(path) & ".xml", FileCounter, path)
        list.Clear()
        Reset_all()
        SetListBox1("[" & DateAndTime.Now & "] " & "Time write XMLFile: " & System.DateTime.Now.Subtract(go).TotalMilliseconds & " ms")
        FileCounter = 0

        SetListBox1("[" & DateAndTime.Now & "] " & "Time to do the Job: " & System.DateTime.Now.Subtract(go).Minutes & " m")
    End Sub
    Public Function GenFileName(ByVal dr As String) As String
        Dim FileName As String
        FileName = CStr(DateAndTime.Now)
        FileName = FileName.Replace(".", "")
        FileName = My.Computer.Name & "." & dr & "." & FileName
        FileName = FileName.Replace("\", "")
        FileName = FileName.Replace("/", "")
        FileName = FileName.Replace(" ", "")
        FileName = FileName.Replace(":", "")
        Return FileName
    End Function
#Region "Delegate&invoke"
    Delegate Sub SetProMaxCallback(ByVal [max] As Integer)
    Delegate Sub SetProMinCallback(ByVal [max] As Integer)
    Delegate Sub SetValueCallback(ByVal [value] As Integer)
    Delegate Sub SetProbarStyleCallback(ByVal [inter] As Integer)
    Delegate Sub SetLabel4Callback(ByVal [text] As String)
    Delegate Sub SetLabel5Callback(ByVal [text] As String)
    Delegate Sub SetLabel6Callback(ByVal [text] As String)
    Delegate Sub SetLabel1Callback(ByVal [text] As String)
    Delegate Sub SetListBox1Callback(ByVal [text] As String)
    Delegate Sub ListBoxScrollDown()

    '''<summary>
    ''' InvokeRequired required compares the thread ID of the
    ''' calling thread to the thread ID of the creating thread.
    ''' If these threads are different, it returns true.
    '''</summary>
    '''
    Public Sub ListBoxScDown()
        If Me.ListBox1.InvokeRequired Then
            Dim d As New ListBoxScrollDown(AddressOf ListBoxScDown)
            Me.Invoke(d, New Object())
        Else
            Me.ListBox1.SelectedIndex = Me.ListBox1.Items.Count - 1
            Me.ListBox1.ClearSelected()
        End If
    End Sub
    Public Sub SetMax(ByVal [max] As Integer)
        If Me.ProgressBar1.InvokeRequired Then
            Dim d As New SetProMaxCallback(AddressOf SetMax)
            Me.Invoke(d, New Object() {[max]})
        Else
            Me.ProgressBar1.Maximum = [max]
        End If
    End Sub
    Public Sub SetMin(ByVal [max] As Integer)
        If Me.ProgressBar1.InvokeRequired Then
            Dim d As New SetProMaxCallback(AddressOf SetMax)
            Me.Invoke(d, New Object() {[max]})
        Else
            Me.ProgressBar1.Minimum = [max]
        End If
    End Sub
    Public Sub SetValue(ByVal [value] As Integer)
        If Me.ProgressBar1.InvokeRequired Then
            Dim d As New SetValueCallback(AddressOf SetValue)
            Me.Invoke(d, New Object() {[value]})
        Else
            Me.ProgressBar1.Value = [value]
        End If

    End Sub
    Public Sub SetLabel4(ByVal [text] As String)
        If Me.Label4.InvokeRequired Then
            Dim d As New SetLabel4Callback(AddressOf SetLabel4)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label4.Text = [text]
        End If
    End Sub
    Public Sub SetLabel5(ByVal [text] As String)
        If Me.Label5.InvokeRequired Then
            Dim d As New SetLabel5Callback(AddressOf SetLabel5)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label5.Text = [text]
        End If
    End Sub
    Public Sub SetLabel6(ByVal [text] As String)
        If Me.Label6.InvokeRequired Then
            Dim d As New SetLabel5Callback(AddressOf SetLabel6)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label6.Text = [text]
        End If
    End Sub
    Public Sub SetLabel1(ByVal [text] As String)
        If Me.Label1.InvokeRequired Then
            Dim d As New SetLabel5Callback(AddressOf SetLabel1)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label1.Text = [text]
        End If
    End Sub
    Public Sub SetListBox1(ByVal [text] As String)
        If Me.ListBox1.InvokeRequired Then
            Dim d As New SetListBox1Callback(AddressOf SetListBox1)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.ListBox1.Items.Add([text])
            ListBoxScDown()
        End If
    End Sub

    Public Sub SetProBarStyle(ByVal [inter] As Int16)
        If Me.ProgressBar1.InvokeRequired Then
            Dim d As New SetProbarStyleCallback(AddressOf SetProBarStyle)
            Me.Invoke(d, New Object() {[inter]})
        Else
            Me.ProgressBar1.Style = [inter]
        End If
    End Sub
#End Region
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Worker = New Thread(AddressOf ReadFiles) With {
            .IsBackground = True
        }
        Worker.Start()
        ProgressBar1.Value = 0
    End Sub
    Private Sub Reset_all()
        SetValue(0)
        SetMin(0)
        SetMax(1000)
        SetLabel1("")
        SetLabel4("")
        SetLabel5("")
        SetLabel6("")
    End Sub
    ''' <summary>
    ''' This method starts at the specified directory, and traverses all subdirectories.
    ''' It returns a List of those directories.
    ''' </summary>
    Public Function GetFilesRecursive(ByVal initial As String) As List(Of String)

        ' This list stores the results.
        Dim result As New List(Of String)

        ' This stack stores the directories to process.
        Dim stack As New Stack(Of String)

        ' Add the initial directory
        stack.Push(initial)


        ' Continue processing for each stacked directory
        Do While (stack.Count > 0)
            ' Get top directory string
            Dim dir As String = stack.Pop
            Try
                ' Add all immediate file paths
                result.AddRange(Directory.GetFiles(dir, "*.*"))
                Me.SetMax(result.Count)
                Me.SetValue(Directory.GetFiles(dir, "*.*").Length)
                ' Loop through all subdirectories and add them to the stack.
                Dim directoryName As String
                For Each directoryName In Directory.GetDirectories(dir)
                    stack.Push(directoryName)
                Next

            Catch ex As Exception
                ErrLog.Add(String.Format("[{0}] {1}", DateAndTime.Now, ex.Message))
                SetListBox1(String.Format("[{0}] {1}", DateAndTime.Now, ex.Message))
                File.WriteAllLines(String.Format("{0}ErrorLog.txt", Drive), ErrLog.ToArray)
                'GeWis = 6
                Me.SetLabel6("Working ... With " & ErrLog.Count.ToString & " Errors")
            End Try


            Me.SetLabel4(result.Count.ToString)
            Me.SetLabel5(CStr(ProgressBar1.Maximum))
            Me.SetLabel6(stack.Count.ToString)
            If stack.Count > 0 Then
                Me.SetLabel1("True")
            Else
                Me.SetLabel1("False")
            End If

        Loop
        ' Return the list
        Return result

    End Function

    Sub XMLWrite(ByVal liste As List(Of String), ByVal FileN As String, ByVal count As UInteger, ByVal DriveLabel As String)
        Try
            Dim XMLName As String = GetDriveLabel(DriveLabel)
            XMLName = XMLName.Replace(" ", "")
            XMLName = XMLName.Replace("(VM)", "")
            XMLName = XMLName.Replace("(TMP)", "")
            SetListBox1("[" & DateAndTime.Now & "] " & "Write XML to " & FileN)
            Reset_all()
            Dim FileXMLList((count - 1)) As flcm
            Dim FileID As Integer = 0
            SetMax(count)
            For Each el In liste
                FileXMLList(FileID) = New flcm(CUInt(FileID), GetFileName(el), GetFullPath(el), GetExtension(el), CULng(My.Computer.FileSystem.GetFileInfo(el).Length))
                FileID = FileID + 1
                SetValue(FileID)
                SetLabel1("Processing File: " & FileID & " of " & count)
            Next
            ' Create XmlWriterSettings.
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True
            settings.Async = True
            ' Create XmlWriter.
            Using writer As XmlWriter = XmlWriter.Create(FileN, settings)
                ' Begin writing.
                writer.WriteStartDocument()
                writer.WriteStartElement(XMLName) ' Root.
                writer.WriteElementString("File_Count", liste.Count.ToString)
                ' Loop over files in array.
                Dim _flcm As flcm
                For Each _flcm In FileXMLList
                    writer.WriteStartElement("File")
                    writer.WriteElementString("ID", _flcm._id.ToString)
                    writer.WriteElementString("File_Name", _flcm._filename)
                    writer.WriteElementString("Full_Path", _flcm._fullfilepath)
                    writer.WriteElementString("File_Extension", _flcm._Extension)
                    writer.WriteElementString("Size", _flcm._filesize.ToString)
                    writer.WriteEndElement()
                Next
                ' End document.

                writer.WriteEndElement()
                writer.WriteEndDocument()

            End Using
        Catch ex As Exception
            ErrLog.Add("[" & DateAndTime.Now & "] " & ex.Message)
            SetListBox1("[" & DateAndTime.Now & "] " & ex.Message)
            File.WriteAllLines(drive & "ErrorLog.txt", ErrLog.ToArray)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim folderdialog As New FolderBrowserDialog
        With folderdialog
            .RootFolder = Environment.SpecialFolder.Desktop
            .ShowDialog()
            path = .SelectedPath
        End With
        SetListBox1("[" & DateAndTime.Now & "] " & "Path to scan: " & path)
        Worker = New Thread(AddressOf ReadPath)
        Worker.IsBackground = True
        Worker.Start()
        ProgressBar1.Value = 0
    End Sub
End Class
''' <summary>
''' Employee type.
''' </summary>
Class flcm
    Public Sub New(ByVal id As UInteger, ByVal filename As String, ByVal fullfilepath As String, ByVal ext As String, ByVal filesize As UInt64)
        ' Set fields.
        Me._id = id
        Me._filename = filename
        Me._fullfilepath = fullfilepath
        Me._filesize = filesize
        Me._Extension = ext
    End Sub

    ' Storage of File data.
    Public _filesize As UInt64
    Public _id As UInteger
    Public _filename As String
    Public _fullfilepath As String
    Public _Extension As String
End Class
