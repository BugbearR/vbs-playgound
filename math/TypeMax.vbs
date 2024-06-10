Dim Byte_Min
Byte_Min = 0
Dim Byte_Max
Byte_Max = 255

Dim Integer_Min
Integer_Min = CInt(-32767) - CInt(1)
Dim Integer_Max
Integer_Max = CInt(32767)

Dim Long_Min
Long_Min = -2147483647 - 1
Dim Long_Max
Long_Max = 2147483647

Dim Single_Min
Single_Min = CSng(-3.402823E+38)
Dim Single_Max
Single_Max = CSng(3.402823E+38)

Dim Double_Min
Double_Min = -1.79769313486231E+308
Dim Double_Max
Double_Max = 1.79769313486231E+308

Dim Currency_Min
Currency_Min = CCur("-922337203685477.5808")
Dim Currency_Max
Currency_Max = CCur("922337203685477.5807")

WScript.Echo CDec(1)
Dim Decimal_Min
Decimal_Min = CDec(1.1)
'Decimal_Min = CDec("-79228162514264337593543950334")
'Dim Decimal_Max
'Decimal_Max = CDec("79228162514264337593543950334")

WScript.Echo Decimal_Min
WScript.Echo TypeName(Decimal_Min)

' Function TypeMax(v)
'     Select TypeName(v)
'     Case "Byte"
'         TypeMax = CByte(255)
'         Exit Function
'     Case "Integer"
'         TypeMax = CInt(32767)
'         Exit Function
'     Case "Currency"
'         TypeMax = CCur(32767)
'         Exit Function
'     End Select
' End Function
