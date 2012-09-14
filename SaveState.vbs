VMName = "仮想マシン名"
Set WMIService = GetObject("winmgmts:\\.\root\virtualization")
Set VMList = WMIService.ExecQuery("SELECT * FROM Msvm_ComputerSystem WHERE ElementName='" & VMName & "'")

'起動
'RequestStateChange(2)
'停止
'RequestStateChange(3)
'保存
'RequestStateChange(32769)

For Each VM In VMList
VM.RequestStateChange(32769)
Next