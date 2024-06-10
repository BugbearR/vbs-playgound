' WMI task: registry https://docs.microsoft.com/ja-jp/windows/win32/wmisdk/wmi-tasks--registry
' StdRegProv methods https://docs.microsoft.com/ja-jp/previous-versions/windows/desktop/regprov/stdregprov


Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001

Function Registry_GetRegProv()
    Set Registry_GetRegProv = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
End Function

Function Registry_CreateKey(hkey, keyPath)
    Dim regProv
    Set regProv = Registry_GetRegProv()
    regProv.CreateKey hkey, keyPath
End Function

Function Registry_SetExpandedStringValue(hkey, keyPath, name, value)
    Dim regProv
    Set regProv = Registry_GetRegProv()
    regProv.SetExpandedStringValue hkey, keyPath, name, value
End Function

Registry_CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\HelloWorldExampleKey"
Registry_SetExpandStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\HelloWorldExampleKey", "hello", "world"

WScript.Echo "Done"
