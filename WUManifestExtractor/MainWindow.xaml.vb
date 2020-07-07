Imports System.IO
Imports System.Windows.Forms
Imports System.Windows.Window
Imports System.Xml
Class MainWindow
    Dim InputDirectory As String
    Dim OutputDiectory As String
    Dim ResourceDirectory As String
    Dim IsKeepStructure As Boolean
    Dim EmptyList As New List(Of String)
    Dim MessageList As New List(Of String)
    Sub RefreshMessageList()
        lstMessage.ItemsSource = EmptyList
        lstMessage.ItemsSource = MessageList
        DoEvents()
    End Sub
    Sub AddMessage(MessageText As String)
        MessageList.Add(MessageText)
        RefreshMessageList()
        lstMessage.SelectedIndex = lstMessage.Items.Count - 1
        lstMessage.ScrollIntoView(lstMessage.SelectedItem)
    End Sub
    Sub LockUI()
        txtInputDir.IsEnabled = False
        txtResourceDir.IsEnabled = False
        txtOutputDir.IsEnabled = False
        btnBrowseInput.IsEnabled = False
        btnBrowseResource.IsEnabled = False
        btnBrowseOutput.IsEnabled = False
        btnStart.IsEnabled = False
        chkKeepStructure.IsEnabled = False
    End Sub
    Sub UnlockUI()
        txtInputDir.IsEnabled = True
        txtResourceDir.IsEnabled = True
        txtOutputDir.IsEnabled = True
        btnBrowseInput.IsEnabled = True
        btnBrowseResource.IsEnabled = True
        btnBrowseOutput.IsEnabled = True
        btnStart.IsEnabled = True
        chkKeepStructure.IsEnabled = True
    End Sub
    Private Sub SetTaskbarProgess(MaxValue As Integer, MinValue As Integer, CurrentValue As Integer, Optional State As Shell.TaskbarItemProgressState = Shell.TaskbarItemProgressState.Normal)
        If MaxValue <= MinValue Or CurrentValue < MinValue Or CurrentValue > MaxValue Then
            Exit Sub
        End If
        TaskbarItem.ProgressValue = (CurrentValue - MinValue) / (MaxValue - MinValue)
        TaskbarItem.ProgressState = State
    End Sub
    Function GetPathFromFile(FilePath As String) As String
        If FilePath.Trim = "" Then
            Return ""
        End If
        If FilePath(FilePath.Length - 1) = "\" Then
            Return FilePath
        End If
        Try
            Return FilePath.Substring(0, FilePath.LastIndexOf("\"))
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetNameFromFullPath(FullPath As String) As String
        If FullPath.Trim = "" Then
            Return ""
        End If
        If FullPath(FullPath.Length - 1) = "\" Then
            Return ""
        End If
        Try
            Return FullPath.Substring(FullPath.LastIndexOf("\") + 1, FullPath.LastIndexOf(".") - FullPath.LastIndexOf("\") - 1)
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetFullNameFromFullPath(FullPath As String) As String
        If FullPath.Trim = "" Then
            Return ""
        End If
        If FullPath(FullPath.Length - 1) = "\" Then
            Return ""
        End If
        Try
            Return FullPath.Substring(FullPath.LastIndexOf("\") + 1)
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Private Sub btnBrowseInput_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseInput.Click
        Dim FolderBrowser As New FolderBrowserDialog
        With FolderBrowser
            .Description = "请指定 Manifest 文件的位置，然后单击""确定""按钮。"
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            InputDirectory = FolderBrowser.SelectedPath
            If InputDirectory(InputDirectory.Length - 1) <> "\" Then
                InputDirectory = InputDirectory & "\"
            End If
            txtInputDir.Text = InputDirectory
        End If
    End Sub
    Private Sub btnBrowseResource_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseResource.Click
        Dim FolderBrowser As New FolderBrowserDialog
        With FolderBrowser
            .Description = "请指定用于抽取文件的目录的位置，然后单击""确定""按钮。"
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            ResourceDirectory = FolderBrowser.SelectedPath
            If ResourceDirectory(ResourceDirectory.Length - 1) <> "\" Then
                ResourceDirectory = ResourceDirectory & "\"
            End If
            txtResourceDir.Text = ResourceDirectory
        End If
    End Sub

    Private Sub btnBrowseOutput_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseOutput.Click
        Dim FolderBrowser As New FolderBrowserDialog
        With FolderBrowser
            .Description = "请指定重建完成的目录结构要输出的位置，然后单击""确定""按钮。"
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            OutputDiectory = FolderBrowser.SelectedPath
            If OutputDiectory(OutputDiectory.Length - 1) <> "\" Then
                OutputDiectory = OutputDiectory & "\"
            End If
            txtOutputDir.Text = OutputDiectory
        End If
    End Sub

    Private Sub btnStart_Click(sender As Object, e As RoutedEventArgs) Handles btnStart.Click
        LockUI()
        If txtInputDir.Text.Trim = "" Then
            MessageBox.Show("CAB 输入路径不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockUI()
            Exit Sub
        End If
        If txtResourceDir.Text.Trim = "" Then
            MessageBox.Show("文件抽取源路径不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockUI()
            Exit Sub
        End If
        If txtOutputDir.Text.Trim = "" Then
            MessageBox.Show("输出路径不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockUI()
            Exit Sub
        End If
        If Not Directory.Exists(OutputDiectory) Then
            Try
                Directory.CreateDirectory(OutputDiectory)
            Catch ex As Exception
                MessageBox.Show("试图创建输出目录""" & OutputDiectory & """时发生错误: " & vbCrLf & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockUI()
                Exit Sub
            End Try
        End If
        With prgProgress
            .Minimum = 0
            .Maximum = 100
            .Value = 0
        End With
        MessageList.Clear()
        RefreshMessageList()

        AddMessage("正在确定 Manifest 文件总数。")
        Dim nManifestFileCount As Integer = Directory.GetFiles(InputDirectory, "*.manifest", SearchOption.TopDirectoryOnly).Length
        If nManifestFileCount = 0 Then
            MessageBox.Show("输入目录""" & InputDirectory & """中不包含任何 Manifest 文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            AddMessage("输入目录""" & InputDirectory & """中不包含任何 Manifest 文件。")
            AddMessage("发生错误，取消操作。")
            UnlockUI()
            Exit Sub
        End If
        AddMessage("计算完毕，共有 " & nManifestFileCount.ToString & " 个 Manifest 文件。")
        With prgProgress
            .Minimum = 0
            .Maximum = nManifestFileCount
            .Value = 0
        End With
        SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
        Dim nSuccess As UInteger = 0
        Dim nFail As UInteger = 0
        Dim nIgnored As UInteger = 0
        Dim IsErrorOccurred As Boolean = False

        For Each ManifestFilePath In Directory.EnumerateFiles(InputDirectory, "*.manifest", SearchOption.TopDirectoryOnly)
            Dim ManifestFileName As String = GetNameFromFullPath(ManifestFilePath)
            Dim UpdateInfoFile As New XmlDocument
            AddMessage("正在打开描述文件""" & ManifestFilePath & """。")
            RefreshMessageList()
            Try
                UpdateInfoFile.Load(ManifestFilePath)
            Catch ex As Exception
                AddMessage("无法打开描述文件""" & ManifestFilePath & """，发生错误: " & ex.Message)
                nFail += 1
                prgProgress.Value += 1
                SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
                Continue For
            End Try
            AddMessage("成功打开描述文件""" & ManifestFilePath & """。")
            Dim nsMgr As New XmlNamespaceManager(UpdateInfoFile.NameTable)
            nsMgr.AddNamespace("ns", "urn:schemas-microsoft-com:asm.v3")
            Dim AssemblyNode As XmlNode = UpdateInfoFile.SelectSingleNode("/ns:assembly", nsMgr)
            AddMessage("正在定位 XML 节点""/assembly""。")
            If IsNothing(AssemblyNode) Then
                AddMessage("XML 节点定位失败。")
                nFail += 1
                prgProgress.Value += 1
                SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
                Continue For
            End If
            AddMessage("XML 节点""/assembly""定位成功，共有 " & AssemblyNode.ChildNodes.Count & " 条记录。")
            Dim TempFileInfo As New WindowsUpdatePackageFileNodeProperties
            Dim FileList As XmlNodeList = AssemblyNode.ChildNodes
            For Each FileNode As XmlNode In FileList
                Dim FileElement As XmlElement = FileNode
                If FileElement.Name <> "file" Then
                    AddMessage("已忽略一个节点，因为它的类型是""" & FileElement.Name & """而不是""file""。")
                    Continue For
                End If
                Try
                    With TempFileInfo
                        .DestinationPath = FileElement.GetAttribute("destinationPath").ToString
                        .Name = FileElement.GetAttribute("name").ToString
                    End With
                    If Not TempFileInfo.DestinationPath.StartsWith("$(runtime.system32)\") And Not TempFileInfo.DestinationPath.StartsWith("$(runtime.bootdrive)\") And Not TempFileInfo.DestinationPath.StartsWith("$(runtime.drivers)\") Then
                        If File.Exists(InputDirectory & TempFileInfo.DestinationPath) Then
                            TempFileInfo.DestinationPath = InputDirectory & TempFileInfo.DestinationPath
                        ElseIf File.Exists(ResourceDirectory & TempFileInfo.DestinationPath) Then
                            TempFileInfo.DestinationPath = ResourceDirectory & TempFileInfo.DestinationPath
                        ElseIf File.Exists(ResourceDirectory & "Windows\WinSxS\Manifests\" & TempFileInfo.DestinationPath) Then
                            TempFileInfo.DestinationPath = ResourceDirectory & "Windows\WinSxS\Manifests\" & TempFileInfo.DestinationPath
                        Else
                            AddMessage("已忽略一个文件节点，因为它没有描述文件复制信息。")
                            Continue For
                        End If
                    End If
                    With TempFileInfo
                        If .DestinationPath.StartsWith("$(runtime.system32)\") Then
                            .DestinationPath = .DestinationPath.Replace("$(runtime.system32)\", ResourceDirectory & "Windows\System32\")
                        End If
                        If .DestinationPath.StartsWith("$(runtime.bootdrive)\") Then
                            .DestinationPath = .DestinationPath.Replace("$(runtime.bootdrive)\", ResourceDirectory)
                        End If
                        If .DestinationPath.StartsWith("$(runtime.drivers)\") Then
                            .DestinationPath = .DestinationPath.Replace("$(runtime.drivers)\", ResourceDirectory & "Windows\System32\Drivers\")
                        End If
                        .Name = OutputDiectory & ManifestFileName & "\" & .Name
                    End With
                    With TempFileInfo
                        If .DestinationPath.Last() <> "\" Then
                            .DestinationPath = .DestinationPath & "\"
                        End If
                        .DestinationPath = .DestinationPath & GetFullNameFromFullPath(.Name)
                    End With
                    Dim CopyDest As String
                    If IsKeepStructure Then
                        CopyDest = TempFileInfo.DestinationPath
                        If CopyDest.StartsWith(InputDirectory) Then
                            CopyDest = CopyDest.Replace(InputDirectory, OutputDiectory & "CopiedDirectlyFromInputDirectory\")
                        Else
                            CopyDest = CopyDest.Replace(ResourceDirectory, OutputDiectory)
                        End If
                    Else
                        CopyDest = TempFileInfo.Name
                    End If
                    Dim CopyDestDir As String = GetPathFromFile(CopyDest)
                    If Not Directory.Exists(CopyDestDir) Then
                        Directory.CreateDirectory(CopyDestDir)
                    End If
                    If File.Exists(CopyDest) Then
                        File.Delete(CopyDest)
                    End If
                    File.Copy(TempFileInfo.DestinationPath, CopyDest)
                    AddMessage("已成功从""" & TempFileInfo.DestinationPath & """复制文件到""" & CopyDest & """。")
                    DoEvents()
                Catch ex As Exception
                    AddMessage("已忽略一个文件节点，因为发生错误: " & ex.Message)
                    Continue For
                End Try
            Next

            Dim RegistryKeysNodeTest As XmlNode = UpdateInfoFile.SelectSingleNode("/ns:assembly/ns:registryKeys", nsMgr)
            AddMessage("正在检查 Manifest 文件是否包含注册表信息。")
            If Not IsNothing(RegistryKeysNodeTest) Then
                AddMessage("已找到注册表信息节点。")
                AddMessage("正在创建注册表文件""" & OutputDiectory & ManifestFileName & ".reg""。")
                Try
                    Dim OutputFileStream As New IO.StreamWriter(OutputDiectory & ManifestFileName & ".reg", False)
                    OutputFileStream.WriteLine("Windows Registry Editor Version 5.00")
                    OutputFileStream.WriteLine()
                    AddMessage("已成功创建注册表文件""" & OutputDiectory & ManifestFileName & ".reg""。")
                    For Each RegistryKeysNode As XmlNode In AssemblyNode
                        If RegistryKeysNode.Name <> "registryKeys" Then
                            Continue For
                        End If
                        Dim RegistryKeyList As XmlNodeList = RegistryKeysNode.ChildNodes
                        For Each RegistryKeyNode As XmlNode In RegistryKeyList
                            If RegistryKeyNode.Name <> "registryKey" Then
                                AddMessage("已忽略一个节点，因为它的类型是""" & RegistryKeyNode.Name & """而不是""registryKey""。")
                                Continue For
                            End If
                            Dim RegistryValueList As XmlNodeList = RegistryKeyNode.ChildNodes
                            AddMessage("已读取到注册表键信息""" & RegistryKeyNode.Attributes("keyName").Value & """。")
                            Try
                                OutputFileStream.WriteLine("[" & RegistryKeyNode.Attributes("keyName").Value & "]")
                            Catch ex As Exception
                                AddMessage("已忽略注册表键结点""" & RegistryKeyNode.Attributes("keyName").Value & """，因为发生错误: " & ex.Message)
                                Continue For
                            End Try
                            Dim RegistryValueInfo As New WindowsUpdatePackageRegistryValueNoteProperties
                            For Each RegistryValueNode As XmlElement In RegistryValueList
                                If RegistryValueNode.Name <> "registryValue" Then
                                    AddMessage("已忽略一个节点，因为它的类型是""" & RegistryValueNode.Name & """而不是""registryValue""。")
                                    Continue For
                                End If
                                With RegistryValueInfo
                                    .Name = RegistryValueNode.GetAttribute("name")
                                    .Value = RegistryValueNode.GetAttribute("value")
                                    .ValueType = RegistryValueNode.GetAttribute("valueType")
                                End With
                                AddMessage("已读取到类型为 " & RegistryValueInfo.ValueType & " 的注册表值。")
                                Try
                                    With RegistryValueInfo
                                        Select Case .ValueType.ToUpper()
                                            Case "REG_NONE" 'hex(0)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(0):" & SplitContinuousBinaryStirng(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(0):" & SplitContinuousBinaryStirng(.Value))
                                                End If
                                            Case "REG_SZ" 'hex(1)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=""" & ManifestRegistryStringValueToRegFileStringValue(.Value) & """")
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=""" & ManifestRegistryStringValueToRegFileStringValue(.Value) & """")
                                                End If
                                            Case "REG_EXPAND_SZ" 'hex(2)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=" & ManifestRegistryExpendableStringValueToRegFileExpendableStringValue(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=" & ManifestRegistryExpendableStringValueToRegFileExpendableStringValue(.Value))
                                                End If
                                            Case "REG_DWORD_LITTLE_ENDIAN" 'hex(3)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=dword:" & .Value.Substring(2, 8))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=dword:" & .Value.Substring(2, 8))
                                                End If
                                            Case "REG_DWORD" 'hex(3)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=dword:" & .Value.Substring(2, 8))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=dword:" & .Value.Substring(2, 8))
                                                End If
                                            Case "REG_DWORD_BIG_ENDIAN" 'hex(4)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(5):" & SplitContinuousBinaryStirng(.Value.Substring(2, 8)))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(5):" & SplitContinuousBinaryStirng(.Value.Substring(2, 8)))
                                                End If
                                            Case "REG_BINARY" 'hex(5)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=" & ManifestRegistryBinaryValueToRegFileBinaryValue(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=" & ManifestRegistryBinaryValueToRegFileBinaryValue(.Value))
                                                End If
                                            Case "REG_LINK" 'hex(6)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(6):" & SplitContinuousBinaryStirng(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(6):" & SplitContinuousBinaryStirng(.Value))
                                                End If
                                            Case "REG_MULTI_SZ" 'hex(7)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=" & ManifestRegistryMultiStringValueToRegFileMultiStringValue(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=" & ManifestRegistryMultiStringValueToRegFileMultiStringValue(.Value))
                                                End If
                                            Case "REG_RESOURCE_LIST" 'hex(8)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(8):" & SplitContinuousBinaryStirng(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(8):" & SplitContinuousBinaryStirng(.Value))
                                                End If
                                            Case "REG_FULL_RESOURCE_DESCRIPTOR" 'hex(9)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(9):" & SplitContinuousBinaryStirng(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(9):" & SplitContinuousBinaryStirng(.Value))
                                                End If
                                            Case "REG_RESOURCE_REQUIREMENT_LIST" 'hex(10)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(10):" & SplitContinuousBinaryStirng(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(10):" & SplitContinuousBinaryStirng(.Value))
                                                End If
                                            Case "REG_QWORD" 'hex(b)
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(b):" & SplitContinuousBinaryStirng(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(b):" & SplitContinuousBinaryStirng(.Value))
                                                End If
                                            Case Else 'use hex(ValueType.ToLower())
                                                If .Name = "" Then
                                                    OutputFileStream.WriteLine("@=hex(" & .ValueType.ToLower() & "):" & SplitContinuousBinaryStirng(.Value))
                                                Else
                                                    OutputFileStream.WriteLine("""" & .Name & """=hex(" & .ValueType.ToLower() & "):" & SplitContinuousBinaryStirng(.Value))
                                                End If
                                        End Select
                                    End With
                                Catch ex As Exception
                                    AddMessage("已忽略一个注册表值节点，因为发生错误: " & ex.Message)
                                    Continue For
                                End Try
                            Next
                            OutputFileStream.WriteLine()
                        Next
                    Next
                    OutputFileStream.Flush()
                    OutputFileStream.Close()
                Catch ex As Exception
                    AddMessage("无法创建或处理注册表文件""" & OutputDiectory & ManifestFileName & ".reg""，因为发生错误: " & ex.Message)
                End Try
            Else
                AddMessage("注册表信息节点定位失败。可能是因为这个 Manifest 文件不包含注册表信息。")
            End If

            Try
                File.Copy(ManifestFilePath, OutputDiectory & ManifestFileName & ".manifest")
            Catch ex As Exception
                AddMessage("无法将 Manifest 文件""" & ManifestFilePath & """复制到""" & OutputDiectory & ManifestFileName & ".manifest""，因为发生错误: " & ex.Message)
                nFail += 1
                prgProgress.Value += 1
                SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
                Continue For
            End Try
            AddMessage("已成功从""" & ManifestFilePath & """复制文件到""" & OutputDiectory & ManifestFileName & ".manifest""。")

            AddMessage("对描述文件""" & ManifestFilePath & """的操作成功完成。")
            nSuccess += 1
            prgProgress.Value += 1
            SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
        Next

        MessageBox.Show("操作完成，共有 " & nSuccess.ToString & "个 Manifest 文件被处理，有 " & nIgnored.ToString & " 个 Manifest 文件被忽略，处理 " & nFail.ToString & " 个 Manifest 文件时出错。", "大功告成!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        UnlockUI()
        With prgProgress
            .Minimum = 0
            .Maximum = 100
            .Value = 0
        End With
        SetTaskbarProgess(100, 0, 0)
    End Sub

    Private Sub chkKeepStructure_Click(sender As Object, e As RoutedEventArgs) Handles chkKeepStructure.Click
        IsKeepStructure = chkKeepStructure.IsChecked
    End Sub
End Class
