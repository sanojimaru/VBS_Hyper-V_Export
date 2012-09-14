option explicit  
 
Dim objWMIService 
Dim managementService 
Dim fileSystem
Dim shell
  
const JobStartIng = 3 
const JobRunnIng = 4 
const JobCompleted = 7 
const wmiStarted = 4096 
const wmiSuccessful = 0 
  
MaIn() 
  
  
'----------------------------------------------------------------- 
' MaIn 
'----------------------------------------------------------------- 
Sub MaIn() 
  
    Dim computer, objArgs, vmName, vm, vmDirectory, exportDirectory 
     
    Set objArgs = WScript.Arguments.Named 
    If WScript.Arguments.Count = 3 Then 
        If objArgs.Exists("VMName") Then 
           vmName = objArgs.Item("VMName") 
        Else 
           WScript.Echo "VMName argument is not provided, Please refer the Usage Section For InFormation on the Arguments" 
           Usage 
        End If

        If objArgs.Exists("VMDirectory") Then 
           vmDirectory = objArgs.Item("VMDirectory") 
        Else 
           WScript.Echo "VMDirectory argument is not provided, Please refer the Usage Section For InFormation on the Arguments" 
           Usage 
        End If  

        If objArgs.Exists("ExportDirectory") Then 
           exportDirectory = objArgs.Item("ExportDirectory") 
        Else 
           WScript.Echo "ExportDirectory argument is not provided, Please refer the Usage Section For InFormation on the Arguments" 
        End If 
    Else 
        Usage  
    End If 
 
    Set fileSystem = Wscript.CreateObject("ScriptIng.FileSystemObject") 
    if fileSystem.FolderExists(exportDirectory) then
        fileSystem.DeleteFolder(exportDirectory)
    end if

    WScript.sleep(5000)

    if Not fileSystem.FolderExists(exportDirectory) then
        fileSystem.CreateFolder(exportDirectory)
    end if
  
    Call ChangeVMState(vmName, 32769)
    ExecCommand("XCOPY " & vmDirectory & " " & exportDirectory & "/e /y")
    Call ChangeVMState(vmName, 2)
End Sub

'起動
'RequestStateChange(2)
'停止
'RequestStateChange(3)
'保存
'RequestStateChange(32769)
Sub ChangeVMState(vmName, state)
    Dim objWMIService, VMs, timeInterval

    Set objWMIService = GetObject("winmgmts:\\.\root\virtualization")
    Set VMs = objWMIService.ExecQuery("SELECT * FROM Msvm_ComputerSystem WHERE ElementName='" & vmName & "'")
    VMs.ItemIndex(0).RequestStateChange(state)

    timeInterval = 1000
    Do while timeInterval <> 0
        WScript.Sleep(timeInterval)
        Set VMs = objWMIService.ExecQuery("SELECT * FROM Msvm_ComputerSystem WHERE ElementName='" & vmName & "'")
        
        Select Case VMs.ItemIndex(0).EnabledState
            Case state
                timeInterval = 0
            Case 32773
                timeInterval = 5000
            Case 32774
                timeInterval = 5000
            Case Else
                timeInterval = 1000
        End Select
    Loop

    Set objWMIService = Nothing
    Set VMs = Nothing
    Set timeInterval = Nothing
End Sub

' 関数名：ExecCommand
' 目　的：DOS コマンドの実行結果を取得します。
Sub ExecCommand(sCommand)
    Dim objWshShell, objExecCmd

    Set objWshShell = WScript.CreateObject("WScript.Shell")
    Set objExecCmd = objWshShell.Exec("%ComSpec% /c " & sCommand)

    Do While objExecCmd.Status = 0
        WScript.Sleep(1000)
    Loop

    Set objExecCmd = Nothing
    Set objWshShell = Nothing
End Sub

'------------------------------------------------------------------------------ 
' The Usage function to convey how to call the script. 
'------------------------------------------------------------------------------ 
Sub Usage() 
    WScript.Echo "Usage: cscript ExportVM.vbs /VMName:vmName /ExportDirectory:exportDirectoryName" 
    WScript.Echo "/VMName: Name of the Virtual MachIne that needs to be Exported." 
    WScript.Echo "/ExportDirectory: Directory to export the Virtual Machine." 
    WScript.Echo "/VMDirectory: Directory of the Virtual Machines." 
    WScript.Quit(1) 
End Sub 
