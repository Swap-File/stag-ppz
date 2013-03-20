Public Class Form1

    Inherits System.Windows.Forms.Form
    Private m_CommPort As New Rs232
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Friend WithEvents LoadBufferButton As System.Windows.Forms.Button
    Friend WithEvents SaveBufferButton As System.Windows.Forms.Button
    Friend WithEvents SendBufferButton As System.Windows.Forms.Button
    Friend WithEvents RecieveBufferButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.LoadBufferButton = New System.Windows.Forms.Button
        Me.SaveBufferButton = New System.Windows.Forms.Button
        Me.SendBufferButton = New System.Windows.Forms.Button
        Me.RecieveBufferButton = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.AddExtension = False
        Me.OpenFileDialog1.Filter = "Binary (*.bin)|*.bin|All Files (*.*)|*.*"
        Me.OpenFileDialog1.ReadOnlyChecked = True
        '
        'LoadBufferButton
        '
        Me.LoadBufferButton.Location = New System.Drawing.Point(47, 54)
        Me.LoadBufferButton.Name = "LoadBufferButton"
        Me.LoadBufferButton.Size = New System.Drawing.Size(85, 39)
        Me.LoadBufferButton.TabIndex = 0
        Me.LoadBufferButton.Text = "Load Buffer From File"
        Me.LoadBufferButton.UseVisualStyleBackColor = True
        '
        'SaveBufferButton
        '
        Me.SaveBufferButton.Location = New System.Drawing.Point(189, 54)
        Me.SaveBufferButton.Name = "SaveBufferButton"
        Me.SaveBufferButton.Size = New System.Drawing.Size(85, 39)
        Me.SaveBufferButton.TabIndex = 1
        Me.SaveBufferButton.Text = "Save Buffer To File"
        Me.SaveBufferButton.UseVisualStyleBackColor = True
        '
        'SendBufferButton
        '
        Me.SendBufferButton.Location = New System.Drawing.Point(189, 254)
        Me.SendBufferButton.Name = "SendBufferButton"
        Me.SendBufferButton.Size = New System.Drawing.Size(85, 39)
        Me.SendBufferButton.TabIndex = 2
        Me.SendBufferButton.Text = "Send Data From Buffer"
        Me.SendBufferButton.UseVisualStyleBackColor = True
        '
        'RecieveBufferButton
        '
        Me.RecieveBufferButton.Location = New System.Drawing.Point(47, 254)
        Me.RecieveBufferButton.Name = "RecieveBufferButton"
        Me.RecieveBufferButton.Size = New System.Drawing.Size(85, 39)
        Me.RecieveBufferButton.TabIndex = 3
        Me.RecieveBufferButton.Text = "Recieve Data To Buffer"
        Me.RecieveBufferButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(61, 140)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Buffer Status:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(138, 140)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Empty"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "bin"
        Me.SaveFileDialog1.Filter = "Binary (*.bin)|*.bin|All Files (*.*)|*.*"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(61, 163)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Current File:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(138, 163)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(33, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "None"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(61, 186)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Checksum:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(138, 186)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(33, 13)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "None"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(52, 320)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(108, 13)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Expected Checksum:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(166, 320)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(33, 13)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "None"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(47, 366)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(227, 20)
        Me.ProgressBar1.TabIndex = 12
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(52, 349)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(99, 13)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "Serial Port Settings:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(166, 349)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(33, 13)
        Me.Label10.TabIndex = 14
        Me.Label10.Text = "None"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(328, 408)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.RecieveBufferButton)
        Me.Controls.Add(Me.SendBufferButton)
        Me.Controls.Add(Me.SaveBufferButton)
        Me.Controls.Add(Me.LoadBufferButton)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "Stag PPZ"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    'Storage buffer
    Public databuffer As Byte()

    'Connection Settings
    Public baud As Long
    Public databits As Long
    Public ComPort As Long
    Public stopbits As Long
    Public parity As Long


    Private Sub LoadBufferButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadBufferButton.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            databuffer = My.Computer.FileSystem.ReadAllBytes(OpenFileDialog1.FileName)

            Label2.Text = (databuffer.Length - 1).ToString + " Bytes Loaded"

            Label4.Text = OpenFileDialog1.FileName.Remove(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) 'keep only file name

            'calculate checksum
            Dim i As Long
            Dim checksum As Long
            For i = 0 To databuffer.Length - 1
                checksum = checksum + databuffer(i)
            Next i
            Label6.Text = Hex$(checksum Mod 256) 'keep only first 4 characters

        End If
    End Sub

    Private Sub SaveBufferButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBufferButton.Click
        If Label2.Text = "Empty" Then
            MsgBox("Buffer Empty - No Data To Save", 0, "Buffer Empty")
        Else
            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                My.Computer.FileSystem.WriteAllBytes(SaveFileDialog1.FileName, databuffer, 0)
                Label2.Text = (databuffer.Length - 1).ToString + " Bytes Saved"
                Label4.Text = SaveFileDialog1.FileName.Remove(0, SaveFileDialog1.FileName.LastIndexOf("\") + 1)
            End If
        End If
    End Sub


    Private Sub RecieveBufferButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecieveBufferButton.Click
        Dim ammountread As Integer
        Dim size As Integer

        'open connection
        Try
            m_CommPort.Open(ComPort, baud, databits, parity, stopbits, 512)
        Catch
            MsgBox("Invalid Port - Check Settings In PPZ.INI", 0, "Invalid Port")
            GoTo done
        End Try


        'get header
        Try
            m_CommPort.Read(6)
        Catch
            MsgBox("No Data - Stag PPZ Not Ready", 0, "No Data")
            GoTo done
        End Try
        'header format:
        '01 00 (Always)
        '-- Size (least significant bits of size)
        '-- Size (most significant bits of size)
        '-- -- Offset (Ignored)

        'convert paypload size hex data to int
        size = (CShort(Val("&H" & Hex(m_CommPort.InputStream(3)))) * 256) + CShort(Val("&H" & Hex(m_CommPort.InputStream(2))))

        'special case of largest size, sending 0x10000 is sent as 0000
        If size = 0 Then
            size = 65536
        End If

        'read in the actual data
        ammountread = 0

        databuffer = New Byte(size - 1) {}

        Do While (ammountread < size)
            Try
                m_CommPort.Read(1)
            Catch
                MsgBox("Connection Lost - Data Transfer Stopped", 0, "Connection Lost")
                GoTo done
            End Try

            databuffer(ammountread) = m_CommPort.InputStream(0)
            ammountread = ammountread + 1
            ProgressBar1.Value = ammountread / size * 100  'Very wasteful, eats CPU, I dont care.
        Loop

        'update status
        Label2.Text = (databuffer.Length - 1).ToString + " Bytes Recieved"
        Label4.Text = "None"

        'update checksum
        Dim checksum As Long
        Dim i As Long
        For i = 0 To databuffer.Length - 1
            checksum = checksum + databuffer(i)
        Next i
        Label6.Text = Hex$(checksum Mod 256) 'keep only first 2 characters 

        'read footer
        Try
            m_CommPort.Read(6)
        Catch
            MsgBox("Connection Lost - Data Transfer Stopped", 0, "Connection Lost")
            GoTo done
        End Try
        'display expected checksum
        Label8.Text = Hex$(m_CommPort.InputStream(0))
        'display error message if bad checksum
        If Label8.Text <> Label6.Text Then
            MsgBox("Checksum Error - Try A Slower Speed", 0, "Checksum Error")
        End If

done:   m_CommPort.Close()

        ProgressBar1.Value = 0
    End Sub

   
    Private Sub SendBufferButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendBufferButton.Click

        If Label2.Text = "Empty" Then
            MsgBox("Buffer Empty - No Data To Send", 0, "Buffer Empty")
        Else

            Dim ammountsent As Long
            Dim Size As Long

            'open connection
            Try
                m_CommPort.Open(ComPort, baud, databits, parity, stopbits, 512)
            Catch
                MsgBox("Invalid Port - Check Settings In PPZ.INI", 0, "Invalid Port")

                m_CommPort.Close()
                Return
            End Try


            'send header (0100)
            m_CommPort.Write(Chr(1) & Chr(0))

            'send size of data
            Size = databuffer.Length

            If Size = 65536 Then
                'Sending 0xffff bytes is 0x10000, so it is a special case. 
                m_CommPort.Write(Chr(0) & Chr(0))
            Else
                'convert length of payload to hex and send.  Least significant first
                'grabs only 2 lsb
                m_CommPort.Write(Chr(Size Mod 256))
                'grabs only 2 msb
                m_CommPort.Write(Chr(Math.Truncate(Size / 256)))
            End If
            


            'send start address (0000)
            m_CommPort.Write(Chr(0) & Chr(0))


            'send the actual data
            ammountsent = 0

            Do While (ammountsent < Size)
                m_CommPort.Write(Chr(databuffer(ammountsent)))
                ammountsent = ammountsent + 1
                ProgressBar1.Value = ammountsent / Size * 100  'Very wasteful, eats CPU, I dont care.
            Loop

            'send checksum
            m_CommPort.Write(Chr(Convert.ToInt16(Label6.Text, 16)))

            'send end code 01 00 06 00 1A I have no idea why this code is important
            'I just am mirroring the code that the programmer sent to me when I started dumping data
            m_CommPort.Write(Chr(1) & Chr(0) & Chr(6) & Chr(0) & Chr(26))

            m_CommPort.Close()

            ProgressBar1.Value = 0
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim objIniFile As New IniFile(CurDir() + "\PPZ.ini")


        'check if ini file exists.  If not, create it with default settings.
        Try
            GetAttr(CurDir() + "\PPZ.ini")
        Catch

            'My Stag PPZ can send to my computer at 4800, and recieve from my computer at 9600
            'More than this causes Checksum Errors.
            objIniFile.WriteString("Settings", "ComPort", 2)
            objIniFile.WriteString("Settings", "baud", 4800)
            objIniFile.WriteString("Settings", "databits", 8)
            objIniFile.WriteString("Settings", "stopbits", 1)
            objIniFile.WriteString("Settings", "parity", 0)

            MsgBox("Defaults Loaded - New PPZ.ini Created", 0, "Defaults Loaded")
        End Try

        ComPort = objIniFile.GetInteger("Settings", "ComPort", -1)
        baud = objIniFile.GetInteger("Settings", "baud", -1)
        databits = objIniFile.GetInteger("Settings", "databits", -1)
        stopbits = objIniFile.GetInteger("Settings", "stopbits", -1)
        parity = objIniFile.GetInteger("Settings", "parity", -1)

        If (ComPort = -1 Or baud = -1 Or databits = -1 Or stopbits = -1 Or parity = -1) Then
            MsgBox("Invalid INI - Delete PPZ.ini To Recreate It With Defaults", 0, "Invalid INI")
            End
        End If

        Dim temp As String
        Select Case parity
            Case 0
                temp = "N"
            Case 1
                temp = "O"
            Case 2
                temp = "E"
            Case 3
                temp = "M"
            Case Else
                temp = "Broken"
        End Select

        Label10.Text = "COM" + ComPort.ToString + "," + baud.ToString + "," + databits.ToString + "," + stopbits.ToString + "," + temp

    End Sub
End Class
