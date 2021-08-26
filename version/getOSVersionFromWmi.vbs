Dim vWmiCimv2
Dim vSWbemObjectSet
Set vWmiCimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set vSWbemObjectSet = wmiCimv2.ExecQuery("SELECT * FROM Win32_OperatingSystem")
WScript.Echo vSWbemObjectSet.ItemIndex(0).Version
