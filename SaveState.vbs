VMName = "���z�}�V����"
Set WMIService = GetObject("winmgmts:\\.\root\virtualization")
Set VMList = WMIService.ExecQuery("SELECT * FROM Msvm_ComputerSystem WHERE ElementName='" & VMName & "'")

'�N��
'RequestStateChange(2)
'��~
'RequestStateChange(3)
'�ۑ�
'RequestStateChange(32769)

For Each VM In VMList
VM.RequestStateChange(32769)
Next