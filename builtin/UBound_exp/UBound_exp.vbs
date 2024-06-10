On Error Resume Next

Dim a0(0)
Dim a1(1)
Dim a2(2)

WScript.Echo UBound(a0)
WScript.Echo UBound(a1)
WScript.Echo UBound(a2)

WScript.Echo "----"
WScript.Echo UBound(Array())
WScript.Echo UBound(Array(2))
WScript.Echo UBound(Array(2,3))
WScript.Echo UBound(Array(2,3,5))

WScript.Echo "----"
Err.Number = 0
a0(0) = 123
WScript.Echo Err.Number
Err.Number = 0
WScript.Echo a0(0)
WScript.Echo Err.Number
Err.Number = 0
a0(1) = 1234
WScript.Echo Err.Number
Err.Number = 0
WScript.Echo a0(1)
WScript.Echo Err.Number

WScript.Echo "----"
Err.Number = 0
a1(0) = 123
WScript.Echo Err.Number
Err.Number = 0
WScript.Echo a1(0)
WScript.Echo Err.Number
Err.Number = 0
a1(1) = 1234
WScript.Echo Err.Number
Err.Number = 0
WScript.Echo a1(1)
WScript.Echo Err.Number
Err.Number = 0
a1(2) = 12345
WScript.Echo Err.Number
Err.Number = 0
WScript.Echo a1(2)
WScript.Echo Err.Number
