REM Collections For LibreOffice Basic
REM Copyright (C) 2022 J. Férard <https://github.com/jferard>
REM
REM Trying to create a sane API for collections in Basic. This is a work in progress.
REM
Option Explicit


REM
REM Helpers
REM

Sub Raise(Optional message As String)
	MsgBox message
	Error(1004)
End Sub


Sub AssertEq(actual As Variant, expected As Variant)
	If actual <> expected Then Raise("Value " & actual & " is different from expected : " & expected)
End Sub

Sub AssertArrayEq(actual() As Variant, expected() As Variant)
	Dim i As Long

	If LBound(actual) <> LBound(expected) Or UBound(actual) <> UBound(expected) Then
		Raise("Arrays have different sizes")
	Else
		For i=LBound(actual) To UBound(actual)
			If actual(i) <> expected(i) Then Raise("Arrays differ at index " & i & ": " & actual(i) & " is different from expected : " & expected(i))
		Next i
	End If
End Sub

Sub AssertRaises(funcName As String, parameters() As Variant)
	Dim script As Object

	On Error GoTo ok
	script = ThisComponent.scriptProvider.getScript("vnd.sun.star.script:Standard.Collections." & funcName & "?language=Basic&location=document")
	script.invoke(parameters, Array(), Array())
	On Error GoTo 0
	Raise("Call " & funcName & "(" & ArrayToString(parameters) & ") did not raise any error")
ok:
End Sub

Sub AssertTests()
	AssertEq("ABC", UCase("abc"))
	AssertArrayEq(Array(), Array())
	AssertArrayEq(Array(1, 2), Array(1, 2))
	AssertRaises("Arrai", Array(1, 2))
End Sub

REM
REM Enums Helpers
REM

''
'' Make an Array out of an XEnumeration
''
Function EnumToArray(e As com.sun.star.container.XEnumeration) As Variant
	Dim arr(8) As Variant
	Dim i As Long

	If Not e.hasMoreElements() Then
		EnumToArray = Array()
		Exit Function
	End If
	
	i = 0
	Do While e.hasMoreElements()
		arr(i) = e.nextElement()
		i = i + 1
		If i > UBound(arr) Then
			ReDim Preserve arr(i * 2)
		End If
	Loop
	
	
	ReDim Preserve arr(i - 1)
	
	EnumToArray = arr
End Function

''
'' Find the size of an XEnumeration
''
Function GetEnumSize(e As com.sun.star.container.XEnumeration) As Long
	Dim i As Long

	i = 0
	Do While e.hasMoreElements()
		e.nextElement()
		i = i + 1
	Loop
	
	GetEnumSize = i
End Function

REM 
REM Tests
REM

Sub ArrayTests
	AssertEq(ArrayToString(Array(4, False, "x", ThisComponent)), "4, False, ""x"", <obj>")
	AssertArrayEq(ReversedArray(Array()), Array())
	
	AssertArrayEq(ReversedArray(Array(4, 5, 6)), Array(6, 5, 4))
	AssertArrayEq(ReversedArray(Array(3, 4, 5, 6)), Array(6, 5, 4, 3))

	' Sorted
	AssertArrayEq(SortedArray(Array(5, 7, 9, 5, 4, 2, 1)), Array(1, 2, 4, 5, 5, 7, 9))
	AssertArrayEq(SortedArray(Array(5, 7, 9, 5, 4, 2, 1, 10, 18, 16)), Array(1, 2, 4, 5, 5, 7, 9, 10, 16, 18))
	AssertArrayEq(SortedArray(Array(80, 16, 9, 14, 68, 0, 46, 98, 74, 37, 18, 58, 69, 28, 62, 53, 76, 2, 57, 20, 11, 72, 84, 86, 50, 78, 39, 40, 27, 94, 81, 67, 61, 26, 12, 96, 19, 71, 92, 47, 75, 6, 42, 55, 54, 17, 21, 66, 8, 59, 63, 45, 88, 44, 49, 41, 4, 83, 22, 31, 82, 99, 5, 48, 79, 1, 73, 77, 65, 38, 90, 30, 91, 32, 43, 25, 33, 35, 85, 60, 87, 51, 36, 70, 7, 29, 56, 93, 24, 15, 89, 52, 13, 95, 10, 34, 64, 3, 23, 97)),_
		Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99))
	AssertArrayEq(SortedArray(Array("80", "16", "9", "14", "68", "0", "46", "98", "74", "37", "18", "58", "69", "28", "62", "53", "76", "2", "57", "20", "11", "72", "84", "86", "50", "78", "39", "40", "27", "94", "81", "67", "61", "26", "12", "96", "19", "71", "92", "47", "75", "6", "42", "55", "54", "17", "21", "66", "8", "59", "63", "45", "88", "44", "49", "41", "4", "83", "22", "31", "82", "99", "5", "48", "79", "1", "73", "77", "65", "38", "90", "30", "91", "32", "43", "25", "33", "35", "85", "60", "87", "51", "36", "70", "7", "29", "56", "93", "24", "15", "89", "52", "13", "95", "10", "34", "64", "3", "23", "97")),_
		Array("0", "1", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "2", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "3", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "4", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "5", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "6", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "7", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "8", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "9", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99"))

	' SortInPlace
	Dim arr As Variant
	arr = Array(5, 7, 9, 5, 4, 2, 1)
	SortArrayInPlace(arr)
	AssertArrayEq(arr, Array(1, 2, 4, 5, 5, 7, 9))

	arr = Array(5, 7, 9, 5, 4, 2, 1, 10, 18, 16)
	SortArrayInPlace(arr)
	AssertArrayEq(arr, Array(1, 2, 4, 5, 5, 7, 9, 10, 16, 18))
	
	arr = Array(80, 16, 9, 14, 68, 0, 46, 98, 74, 37, 18, 58, 69, 28, 62, 53, 76, 2, 57, 20, 11, 72, 84, 86, 50, 78, 39, 40, 27, 94, 81, 67, 61, 26, 12, 96, 19, 71, 92, 47, 75, 6, 42, 55, 54, 17, 21, 66, 8, 59, 63, 45, 88, 44, 49, 41, 4, 83, 22, 31, 82, 99, 5, 48, 79, 1, 73, 77, 65, 38, 90, 30, 91, 32, 43, 25, 33, 35, 85, 60, 87, 51, 36, 70, 7, 29, 56, 93, 24, 15, 89, 52, 13, 95, 10, 34, 64, 3, 23, 97)
	SortArrayInPlace(arr)
	AssertArrayEq(arr,_
		Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99))

	arr = Array("80", "16", "9", "14", "68", "0", "46", "98", "74", "37", "18", "58", "69", "28", "62", "53", "76", "2", "57", "20", "11", "72", "84", "86", "50", "78", "39", "40", "27", "94", "81", "67", "61", "26", "12", "96", "19", "71", "92", "47", "75", "6", "42", "55", "54", "17", "21", "66", "8", "59", "63", "45", "88", "44", "49", "41", "4", "83", "22", "31", "82", "99", "5", "48", "79", "1", "73", "77", "65", "38", "90", "30", "91", "32", "43", "25", "33", "35", "85", "60", "87", "51", "36", "70", "7", "29", "56", "93", "24", "15", "89", "52", "13", "95", "10", "34", "64", "3", "23", "97")
	SortArrayInPlace(arr)
	AssertArrayEq(arr,_
		Array("0", "1", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "2", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "3", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "4", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "5", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "6", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "7", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "8", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "9", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99"))
		
	' Create N arrays
	Dim arrs(1) As Variant
	Dim arr2 As Variant
	Dim i, j, size As Long

	For i=LBound(arrs) To UBound(arrs)
		size = Int(Rnd() * 5000) + 5000
		Redim arr(size-1)
		For j=LBound(arr) To UBound(arr)
			arr(j) = Int(Rnd() * 10000) - 5000
		Next j
		arrs(i) = arr
	Next i

	Dim t0, t1 As Double
	t0 = Now()
	
	For Each arr In arrs
		arr2 = SortedArray(arr)
		For j=LBound(arr) To UBound(arr)-1
			AssertEq(arr2(j) <= arr2(j+1), True)
		Next j
	Next arr

	t1 = Now()
	MsgBox(t1 - t0)

	t0 = Now()

	For Each arr In arrs
		SortArrayInPlace(arr)
		For j=LBound(arr) To UBound(arr)-1
			AssertEq(arr(j) <= arr(j+1), True)
		Next j
	Next arr		
		
	t1 = Now()
	MsgBox(t1 - t0)

	MsgBox "Shuffle :" & ArrayToString(ShuffledArray(SortedArray(Array(5, 7, 9, 5, 4, 2, 1, 10, 18, 16))))
End Sub



Sub ListTests
	Dim list As Variant
	list = NewList(Array(4, 5, 6))
	AppendListElement(list, "a")
	AppendListElement(list, "b")
	AppendListElement(list, "c")
	
	AssertEq(GetListSize(list), 6)
	AssertEq(ListIsEmpty(list), False)
	AssertEq(ListIndexOf(list, "b"), 4)
	AssertEq(ListLastIndexOf(list, "b"), 4)
	AssertEq(ListIndexOf(list, "d"), -1)
	
	AppendListElement(list, "d")
	AppendListElement(list, 2)
	AppendListElement(list, "f")
	AppendListElement(list, "g")
	AppendListElement(list, "h")
	SetListElement(list, 9, "z")
	AppendListElements(list, Array("Z", "W", "T"))
	InsertListElement(list, 5, "XYZ")
	RemoveListElement(list, 10)
	AssertArrayEq(ListToArray(list), Array(4, 5, 6, "a", "b", "XYZ", "c", "d", 2, "f", "h", "Z", "W", "T"))
	AssertEq(ListToString(list), "[4, 5, 6, ""a"", ""b"", ""XYZ"", ""c"", ""d"", 2, ""f"", ""h"", ""Z"", ""W"", ""T""]")
End Sub


Sub SetTests
	Dim s As Variant
	Dim e As Variant
	s = NewEmptySet("string")
	AssertEq(SetIsEmpty(s), True)
	
	
	s = NewSetFromArray("long", Array(1, 3, 5))
	AddSetElement(s, 15)
	AddSetElement(s, 8)
	AddSetElement(s, 3)

	AssertArrayEq(SetToArray(s), Array(1, 3, 5, 8, 15))
	AssertEq(SetToString(s), "{1, 3, 5, 8, 15}")
	AssertEq(GetSetSize(s), 5)
	
	Do While Not SetIsEmpty(s)
		e = TakeSetElement(s) 
		AssertEq(IsEmpty(e), False)
	Loop
	
	AssertEq(SetIsEmpty(s), True)
	AssertArrayEq(SetToArray(s), Array())

	e = TakeSetElement(s) 
	AssertEq(IsEmpty(e), True)
End Sub


Sub MapTests
	Dim m As Variant
	
	m = NewEmptyMap("string", "long")
	
	AssertEq(MapIsEmpty(m), True)
	
	PutMapElement(m, "a", 1)

	AssertEq(MapIsEmpty(m), False)
	AssertEq(GetMapSize(m), 1)

	PutMapElement(m, "b", 2)
	PutMapElement(m, "c", 3)

	AssertArrayEq(MapKeysToArray(m), SortedArray(Array("a", "b", "c")))
	AssertArrayEq(MapValuesToArray(m), SortedArray(Array(1, 2, 3)))

	AssertEq(GetMapSize(m), 3)
	
	PutMapElement(m, "d", 4)
	AssertEq(GetMapElementOrDefault(m, "e", 20), 20)

	AssertEq(MapContains(m, "b"), True)
	AssertEq(GetMapSize(m), 4)
	AssertEq(MapToString(m), "{""a"": 1, ""b"": 2, ""c"": 3, ""d"": 4}")
	AssertEq(GetMapElement(m, "b"), 2)
	AssertEq(RemoveMapElement(m, "b"), 2)
	AssertEq(MapContains(m, "b"), False)
	AssertEq(GetMapSize(m), 3)
	
	AssertRaises("GetMapElement", Array(m, "b"))
	
	Dim m2 As Variant
	m2 = NewEmptyMap("string", "long")
	PutMapElement(m2, "foo", 100)
	
	m = MergeMaps(m, m2)
	AssertEq(MapToString(m), "{""a"": 1, ""c"": 3, ""d"": 4, ""foo"": 100}")
End Sub

Sub AllTests()
	AssertTests()
	ArrayTests()
	ListTests()
	SetTests()
	MapTests()
End Sub

