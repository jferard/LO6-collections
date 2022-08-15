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
	Dim i As Integer

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

Sub AssertExamples()
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
	Dim i As Integer

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
Function GetEnumSize(e As com.sun.star.container.XEnumeration) As Integer
	Dim i As Integer

	i = 0
	Do While e.hasMoreElements()
		e.nextElement()
		i = i + 1
	Loop
	
	GetEnumSize = i
End Function


REM 
REM Arrays
REM 

''
'' Copy an array
''
Function CopyArray(arr() As Variant) As Variant
	Dim newArr() As Variant
	newArr = arr
	ReDim Preserve newArr(LBound(arr) To Ubound(arr))
	CopyArray = newArr
End Function

''
'' Return the reversed array
''
Function ReversedArray(arr() As Variant) As Variant
	Dim i, j As Variant
	
	If UBound(arr) < LBound(arr) Then
		ReversedArray = Array()
		Exit Function
	End If 
	
	Dim reversed(LBound(arr) To UBound(arr)) As Variant
	

	i = LBound(arr)
	j = UBound(arr)
	Do While i <= j
		reversed(i) = arr(j)
		reversed(j) = arr(i)
		i = i + 1
		j = j - 1
	Loop
	
	ReversedArray = reversed
End Function

''
'' Return the array sorted (merge sort algorithm)
''
Function SortedArray(arr() As Variant) As Variant
	Dim size As Integer
	Dim i, j, n As Integer
	Dim cur, other, temp As Variant
	
	size = UBound(arr) - LBound(arr) + 1
	
	If size <= 0 Then
		ReversedArray = Array()
		Exit Function
	End If 

	Dim arr1 As Variant
	Dim arr2(LBound(arr) To UBound(arr)) As Variant

	arr1 = _CopySwap(arr)

	n = 2
	cur = arr1
	other = arr2
	Do While n <= size
		For i = LBound(arr) To UBound(arr) Step 2 * n
			_Merge(cur, i, n, other)
		Next i
		n = n * 2
		temp = cur
		cur = other
		other = temp
	Loop

	SortedArray = cur
End Function

' prepare the first copy : swap pairs if necessary
Function _CopySwap(arr As Variant) As Variant
	Dim arr1(LBound(arr) To UBound(arr)) As Variant
	Dim i As Integer
	
	arr1(UBound(arr)) = arr(UBound(arr))
	For i=LBound(arr) To UBound(arr) - 1 Step 2
		If arr(i) < arr(i+1) Then
			arr1(i)     = arr(i)
			arr1(i + 1) = arr(i + 1)
		Else
			arr1(i)     = arr(i + 1)
			arr1(i + 1) = arr(i)
		End If
	Next i
	_CopySwap = arr1
End Function

' merge two sorted sequences
Sub _Merge(cur As Variant, i As Integer, n As Integer, other As Variant)
	Dim a, b, c, s As Integer

	s = i + 2 * n
	If s > UBound(other) + 1 Then s = UBound(other) + 1
	
	c = i
	a = i
	b = i + n
	If b >= s Then Exit Sub 
	
	Do While True
		If cur(a) < cur(b) Then
			other(c) = cur(a)
			c = c + 1
			If c = s Then Exit Sub
			a = a + 1
			If a = i + n Then 
				' flush b
				Do While c < s
					other(c) = cur(b)
					c = c + 1
					b = b + 1
				Loop
				Exit Sub
			End If
		Else
			other(c) = cur(b)
			c = c + 1
			If c = s Then Exit Sub
			b = b + 1
			If b = s Then 
				Do While c < s
					other(c) = cur(a)
					c = c + 1
					a = a + 1
				Loop
				Exit Sub
			End If
		End If
	Loop
End Sub

''
'' Return the array shuffled (Fisher Yates shuffle, a.k.a Knuth)
''
Function ShuffledArray(arr() As Variant) As Variant
	If UBound(arr) < LBound(arr) Then
		ShuffledArray = Array()
		Exit Function
	End If

	Dim i, j As Integer
	Dim temp, newArr() As Variant
	newArr = copyArray(arr)

    For i = LBound(newArr) To UBound(newArr) - 1
	    j = i + Int(Rnd() * (UBound(newArr) - i + 1)) ' i <= j < i + UBound(arr) - i + 1 = UBound(arr) + 1
	    temp = newArr(i)
	    newArr(i) = newArr(j)
	    newArr(j) = temp
    Next i
    
    ShuffledArray = newArr
End Function

''
'' Return a representation of this array 
''
Function ArrayToString(arr() As Variant) As String
	Dim i, vt As Integer

	Dim newArr(LBound(arr) To UBound(arr)) As String
	For i=LBound(arr) To UBound(arr)
		Select Case VarType(arr(i))
		Case V_STRING
			newArr(i) = """" & arr(i) & """"
		Case 11
			If arr(i) Then 
				newArr(i) = "True"
			Else
				newArr(i) = "False"
			End If
		Case 9
			newArr(i) = "<obj>"
		Case Else
			newArr(i) = CStr(arr(i))
		End Select
	Next i
	ArrayToString = Join(newArr, ", ")
End Function


Sub ArrayExamples
	AssertEq(ArrayToString(Array(4, False, "x", ThisComponent)), "4, False, ""x"", <obj>")
	AssertArrayEq(ReversedArray(Array()), Array())
	
	AssertArrayEq(ReversedArray(Array(4, 5, 6)), Array(7, 5, 4))
	AssertArrayEq(ReversedArray(Array(3, 4, 5, 6)), Array(6, 5, 4, 3))

	AssertArrayEq(SortedArray(Array(5, 7, 9, 5, 4, 2, 1)), Array(1, 2, 4, 5, 5, 7, 9))
	AssertArrayEq(SortedArray(Array(5, 7, 9, 5, 4, 2, 1, 10, 18, 16)), Array(1, 2, 4, 5, 5, 7, 9, 10, 16, 18))
	AssertArrayEq(SortedArray(Array(80, 16, 9, 14, 68, 0, 46, 98, 74, 37, 18, 58, 69, 28, 62, 53, 76, 2, 57, 20, 11, 72, 84, 86, 50, 78, 39, 40, 27, 94, 81, 67, 61, 26, 12, 96, 19, 71, 92, 47, 75, 6, 42, 55, 54, 17, 21, 66, 8, 59, 63, 45, 88, 44, 49, 41, 4, 83, 22, 31, 82, 99, 5, 48, 79, 1, 73, 77, 65, 38, 90, 30, 91, 32, 43, 25, 33, 35, 85, 60, 87, 51, 36, 70, 7, 29, 56, 93, 24, 15, 89, 52, 13, 95, 10, 34, 64, 3, 23, 97)),_
		Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 23, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99))
	AssertArrayEq(SortedArray(Array("80", "16", "9", "14", "68", "0", "46", "98", "74", "37", "18", "58", "69", "28", "62", "53", "76", "2", "57", "20", "11", "72", "84", "86", "50", "78", "39", "40", "27", "94", "81", "67", "61", "26", "12", "96", "19", "71", "92", "47", "75", "6", "42", "55", "54", "17", "21", "66", "8", "59", "63", "45", "88", "44", "49", "41", "4", "83", "22", "31", "82", "99", "5", "48", "79", "1", "73", "77", "65", "38", "90", "30", "91", "32", "43", "25", "33", "35", "85", "60", "87", "51", "36", "70", "7", "29", "56", "93", "24", "15", "89", "52", "13", "95", "10", "34", "64", "3", "23", "97")),_
		Array("0", "1", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "2", "20", "21", "22", "24", "25", "26", "27", "28", "29", "3", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "4", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "5", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "6", "60", "61", "62", "63", "64", "23", "65", "66", "67", "68", "69", "7", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "8", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "9", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99"))
	MsgBox ArrayToString(ShuffledArray(SortedArray(Array(5, 7, 9, 5, 4, 2, 1, 10, 18, 16))))
End Sub

REM
REM Lists
REM

''
'' The ArrayList Type
'' The type of elements may be mixed.
''
Type ArrayList
	arr() As Variant
	size As Integer
End Type

''
'' Create a new ArrayList, given an initial capacity or array of elements.
'' 
Function NewList(Optional initialCapacityOrArr As Variant) As ArrayList
	If IsMissing(initialCapacityOrArr) Then 
		NewList = NewListWithCapacity(8)
	ElseIf IsArray(initialCapacityOrArr) Then
		NewList = NewListFromArray(initialCapacityOrArr)
	Else
		NewList = NewListWithCapacity(initialCapacityOrArr)
	End If
End Function

''
'' Create a new ArrayList, given an initial capacity.
'' 
Function NewListWithCapacity(capacity As Integer) As ArrayList
	Dim list As ArrayList
	Dim arr(capacity - 1) As Variant
	
	list.arr = arr
	list.size = 0
	NewListWithCapacity = list
End Function

''
'' Create a new ArrayList, given an initial array of elements
'' 
Function NewListFromArray(arr As Variant) As ArrayList
	If LBound(arr) <> 0 Then arr(0) ' trigger an error

	Dim list as ArrayList
	Dim newArr() As Variant
	Dim capacity As Integer

	If UBound(arr) < 7 Then
		capacity = 8
	Else
		capacity = Round(UBound(arr) * 1,5)
	End If
	
	newArr = arr
	ReDim Preserve newArr(capacity - 1) As Variant
	
	list.arr = newArr
	list.size = UBound(arr) + 1 
	NewListFromArray = list
End Function

''
'' Update the list capacity
''
Sub SetListCapacity(list As ArrayList, newCapacity As Integer)
	Dim arr As Variant

	If newCapacity < list.size Then Exit Sub
	
	arr = list.arr
	Redim Preserve arr(newCapacity) As Variant
	list.arr = arr
End Sub

''
'' Append an element to a list
''
Sub AppendListElement(list As ArrayList, element As Variant)
	Dim arr() As Variant
	
	_EnsureListCapacity(list, list.size + 1)
	
	list.arr(list.size) = element
	list.size = list.size + 1
End Sub

''
'' Ensure list capacity
''
Sub _EnsureListCapacity(list As ArrayList, capacity As Integer)
	Dim arr As Variant

	If capacity <= UBound(list.arr) + 1 Then Exit Sub
	
	arr = list.arr
	Redim Preserve arr(capacity * 1.2) As Variant
	list.arr = arr
End Sub


''
'' Append an array of elements to a list
''
Sub AppendListElements(list As ArrayList, elements() As Variant)
	Dim i, newCapacity, elementsSize As Integer
	Dim arr() As Variant
	
	If LBound(elements) <> 0 Then elements(0) ' trigger an error

	elementsSize = UBound(elements) + 1
	_EnsureListCapacity(list, list.size + elementsSize)

	For i=0 To elementsSize - 1
		list.arr(list.size) = elements(i)
		list.size = list.size + 1	
	Next i
End Sub

''
'' Pop an element from a list
'' 
Function PopListElement(list As ArrayList) As Variant
	PopListElement = list.arr(list.size - 1)
	list.size = list.size - 1
End Function

''
'' Get the element of a list at a given index. 
''
Function GetListElement(list As ArrayList, index As Integer) As Variant
	If index >= list.size Then list.arr(-1)

	GetListElement = list.arr(index)
End Function

''
'' Set the element of a list at a given index to a value
''
Sub SetListElement(list As ArrayList, index As Integer, element As Variant)
	If index >= list.size Then list.arr(-1)

	list.arr(index) = element
End Sub

''
'' Insert the element at a given index
''
Sub InsertListElement(list As ArrayList, index As Integer, element As Variant)
	Dim arr() As Variant
	Dim i As Integer
	
	_EnsureListCapacity(list, list.size + 1)

	For i = list.size To index + 1 Step -1
		list.arr(i) = list.arr(i - 1)
	Next i
	
	list.arr(index) = element
	list.size = list.size + 1
End Sub

''
'' Remove the element at a given index
''
Sub RemoveListElement(list As ArrayList, index As Integer)
	Dim arr() As Variant
	Dim i As Integer
	
	For i = index To list.size - 2
		list.arr(i) = list.arr(i + 1)
	Next i
	
	list.size = list.size - 1
End Sub

''
'' Reverse a List
''
Sub ReverseList(list As ArrayList)
	Dim arr() As Variant
	
	arr = list.arr
	arr = ReversedArray(arr)
	list.arr = arr
End Sub

''
'' Sort a List
''
Sub SortList(list As ArrayList)
	Dim arr() As Variant
	
	arr = list.arr
	arr = SortedArray(arr)
	list.arr = arr
End Sub

''
'' Shuffle a List
''
Sub ShuffleList(list As ArrayList)
	Dim arr() As Variant
	
	arr = list.arr
	arr = ShuffledArray(arr)
	list.arr = arr
End Sub

''
'' Copy the list into an array
''
Function ListToArray(list As ArrayList) As Variant
	If list.size = 0 Then
		ListToArray = Array()
	Else
		Dim arr() As Variant
		arr = list.arr
		Redim Preserve arr(list.size - 1)
		ListToArray = arr
	End If
End Function

''
'' Return the size of a list
''
Function GetListSize(list As ArrayList) As Integer
	GetListSize = list.size
End Function

''
'' Return True if the list is empty
''
Function ListIsEmpty(list As ArrayList) As Boolean
	ListIsEmpty = (list.size = 0)
End Function

''
'' Return the index of element in the list, or -1 if element is not in the list
''
Function ListIndexOf(list As ArrayList, element As Variant) As Integer
	Dim i As Integer

	For i=0 To list.size - 1
		If list.arr(i) = element Then
			ListIndexOf = i
			Exit Function
		End If
	Next i
	ListIndexOf = -1
End Function


Sub ListExamples
	Dim list As ArrayList
	list = NewList(Array(4, 5, 6))
	AppendListElement(list, "a")
	AppendListElement(list, "b")
	AppendListElement(list, "c")
	
	AssertEq(GetListSize(list), 6)
	AssertEq(ListIsEmpty(list), False)
	AssertEq(ListIndexOf(list, "b"), 4)
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
End Sub

REM 
REM Sets
REM

''
'' The HashSet type
''
Type HashSet
	typeName As String
	map As com.sun.star.container.EnumerableMap
End Type

''
'' Create a new Set of type typeName
''
Function NewSet(typeName As String, Optional arr As Variant) As HashSet
	If IsMissing(arr) Then 
		NewSet = NewEmptySet(typeName)
	Else
		NewSet = NewSetFromArray(typeName, arr)
	End If
End Function

''
'' Create a new empty Set of type typeName
''
Function NewEmptySet(typeName As String) As HashSet
	Dim s As HashSet

	s.typeName = typeName
	s.map = com.sun.star.container.EnumerableMap.create(typeName, "byte")

	NewEmptySet = s
End Function

''
'' Create a new Set of type typeName having the elements arr
''
Function NewSetFromArray(typeName As String, arr() As Variant) As HashSet
	Dim s As HashSet

	s = NewEmptySet(typeName)
	AddSetElements(s, arr)

	NewSetFromArray = s
End Function

''
'' Add an element to a Set
''
Function AddSetElement(s As HashSet, element As Variant)
	s.map.put(CreateUnoValue(s.typeName, element), 1)
End Function

''
'' Add some elements to a Set
''
Sub AddSetElements(s As HashSet, arr() As Variant)
	Dim element As Variant
	
	For Each element In arr	
		AddSetElement(s, element)
	Next element
End Sub

''
'' Remove an element from a Set
''
Function RemoveSetElement(s As HashSet, element As Variant)
	s.map.remove(CreateUnoValue(s.typeName, element))
End Function

''
'' Return True if the Set contains this element
''
Function SetContains(s As HashSet, element As Variant) As Boolean
	Contains = s.map.containsKey(element)
End Function

''
'' Remove a random element from the Set and return it
''
Function TakeSetElement(s As HashSet) As Variant
	Dim e As Object
	Dim element As Variant
	e = s.map.createKeyEnumeration(True)

	If e.hasMoreElements() Then
		element = e.nextElement()
		s.map.remove(element)
		TakeSetElement = element
	Else
		TakeSetElement = Empty
	End If
End Function

''
'' Copy the elements of the Set to an Array
''
Function SetToArray(s As HashSet) As Variant
	Dim e As Object
	
	e = s.map.createKeyEnumeration(True)
	SetToArray = EnumToArray(e)
End Function

''
'' Return the Set size
''
Function GetSetSize(s As HashSet) As Variant
	Dim e As Object
	
	e = s.map.createKeyEnumeration(True)
	GetSetSize = _GetEnumSize(e)
End Function

''
'' Return True if the Set is empty
''
Function SetIsEmpty(s As HashSet) As Variant
	Dim e As Object
	
	e = s.map.createKeyEnumeration(True)
	SetIsEmpty = Not e.hasMoreElements()
End Function

Sub SetExamples
	Dim s As HashSet
	s = NewSetFromArray("long", Array(1, 3, 5))
	AddSetElement(s, 15)
	AddSetElement(s, 8)
	AddSetElement(s, 3)
	AssertArrayEq(SetToArray(s), Array(1, 3, 5, 8, 15))

	AssertEq(GetSetSize(s), 5)
	
	Dim e As Variant
	Do While True
    		e = TakeSetElement(s) 
		If IsEmpty(e) Then Exit Do
	Loop
	AssertEq(SetIsEmpty(s), True)
	AssertArrayEq(SetToArray(s), Array())
End Sub

REM 
REM Maps
REM

''
'' The HashMap type
''
Type HashMap
	keyTypeName As String
	valueTypeName As String
	map As com.sun.star.container.EnumerableMap
End Type

''
'' Create a new Map.
''
Function NewEmptyMap(keyTypeName As String, valueTypeName As String) As HashSet
	Dim m As HashMap

	m.keyTypeName = keyTypeName
	m.valueTypeName = valueTypeName
	m.map = com.sun.star.container.EnumerableMap.create(keyTypeName, valueTypeName)

	NewEmptyMap = m
End Function

''
'' Remove a key-value pair in the Map
''
Function PutMapElement(m As HashMap, key As Variant, value As Variant)
	m.map.put(CreateUnoValue(m.keyTypeName, key), CreateUnoValue(m.valueTypeName, value))
End Function

''
'' Remove a Map element by key
''
Function RemoveMapElement(m As HashMap, key As Variant)
	m.map.remove(CreateUnoValue(m.keyTypeName, key))
End Function

''
'' Return True if the Map contains this Key
''
Function MapContains(m As HashMap, key As Variant) As Boolean
	MapContains = m.map.containsKey(CreateUnoValue(m.keyTypeName, key))
End Function

''
'' Return the value mapped to this key, or raise an exception
''
Function GetMapElement(m As HashMap, key As Variant) As Variant
	GetMapElement = m.map.get(CreateUnoValue(m.keyTypeName, key))
End Function

''
'' Return the value mapped to this key, or a default value
''
Function GetMapElementOrDefault(m As HashMap, key As Variant, default As Variant) As Variant
	If m.map.containsKey(CreateUnoValue(m.keyTypeName, key)) Then
		GetMapElementOrDefault = m.map.get(CreateUnoValue(m.keyTypeName, key))
	Else
		GetMapElementOrDefault = default
	End If
End Function

''
'' Return an Array of the keys.
''
Function MapKeysToArray(m As HashMap) As Variant
	Dim e As Object
	
	e = m.map.createKeyEnumeration(True)
	MapKeysToArray = EnumToArray(e)
End Function

''
'' Return an Array of the values.
''
Function MapValuesToArray(m As HashMap) As Variant
	Dim e As Object
	
	e = m.map.createValueEnumeration(True)
	MapValuesToArray = EnumToArray(e)
End Function


''
'' Return an Set of the keys.
''
Function MapKeysToSet(m As HashMap) As Variant
	Dim s As HashSet

	s = NewEmptySet(m.keyTypeName)
	
	e = m.map.createKeyEnumeration(True)
	Do While e.hasNext()
		AddSetElement(e.nextElement())
	Loop	
	
	MapKeysToSet = s
End Function

''
'' Return the Map size
''
Function GetMapSize(m As HashMap) As Variant
	Dim e As Object
	
	e = m.map.createKeyEnumeration(True)
	GetMapSize = GetEnumSize(e)
End Function

''
'' Return True if the Map is empty
''
Function MapIsEmpty(m As HashMap) As Variant
	Dim e As Object
	
	e = m.map.createKeyEnumeration(True)
	MapIsEmpty = Not e.hasMoreElements()
End Function

Sub MapExamples
	Dim m As HashMap
	
	m = NewEmptyMap("string", "long")
	
	AssertEq(MapIsEmpty(m), True)
	
	PutMapElement(m, "a", 1)

	AssertEq(MapIsEmpty(m), False)
	AssertEq(GetMapSize(m), 1)

	PutMapElement(m, "b", 2)
	PutMapElement(m, "c", 3)

	AssertArrayEq(MapKeysToArray(m), SortedArray(Array("a", "b", "c")))
	AssertArrayEq(MapValuesToArray(m), SortedArray(Array(1, 2, 3)))
	
	PutMapElement(m, "d", 4)
	AssertEq(MapContains(m, "b"), True)
	AssertEq(GetMapElement(m, "b"), 2)
	RemoveMapElement(m, "b")
	AssertEq(GetMapElementOrDefault(m, "e", 20), 20)
	AssertEq(MapContains(m, "b"), False)
	AssertRaises("GetMapElement", Array(m, "b"))
End Sub