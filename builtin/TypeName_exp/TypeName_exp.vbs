Dim vEmpty
Dim vNothing
Dim vNull
Dim vInteger
Dim vDouble
Dim vString
Dim vEmptyString
Dim vArray0
Dim vArray1
Dim vArray1Str
Dim vCurrency

Set vNothing = Nothing
vNull = Null
vInteger = 123
vDouble = 1.23
vDouble = 1.23
vString = "Hello, world!"
vEmptyString = ""
vArray0 = Array()
vArray1 = Array(1)
vArray1Str = Array("a")
vCurrency = CCur("922337203685477.5807")
' Compile error
'vCurrency = 922337203685477.5807@

WScript.Echo TypeName(vEmpty)
WScript.Echo TypeName(vNothing)
WScript.Echo TypeName(vNull)
WScript.Echo TypeName(vInteger)
WScript.Echo TypeName(vDouble)
WScript.Echo TypeName(vString)
WScript.Echo TypeName(vEmptyString)
WScript.Echo TypeName(vArray0)
WScript.Echo TypeName(vArray1)
WScript.Echo TypeName(vArray1Str)
WScript.Echo TypeName(vCurrency)
