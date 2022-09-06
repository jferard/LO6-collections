REM Collections For LibreOffice Basic
REM Copyright (C) 2022 J. Férard <https://github.com/jferard>
REM
REM Trying to create a sane API for collections in Basic. This is a work in progress.
REM
Option Explicit

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
	Dim size As Long
	Dim i, j, n As Long
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
	Dim i As Long
	
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
Sub _Merge(cur As Variant, i As Long, n As Long, other As Variant)
	Dim a, b, c, s As Long

	s = i + 2 * n
	If s > UBound(other) + 1 Then s = UBound(other) + 1
	
	c = i
	a = i
	b = i + n
	If b >= s Then 
		Do While c < s
			other(c) = cur(a)
			c = c + 1
			a = a + 1
		Loop
		Exit Sub 
	End If
	
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
				' flush a
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
'' Sort the array in place.
''
Sub SortArrayInPlace(ByRef arr() As Variant)
	_QuickSort(arr, LBound(arr), UBound(arr))
End Sub

Sub _QuickSort(ByRef arr() As Variant, p As Long, r As Long)
	If p >= r Then Exit Sub
	Dim q As Long
	q = _Partition(arr, p, r)
	_QuickSort(arr, p, q - 1)
	_QuickSort(arr, q + 1, r)
End Sub

Function _Partition(ByRef arr() As Variant, p As Long, r As Long) As Long
	Dim x, temp As Variant
	Dim i, j As Long

	i = Int(Rnd() * (r - p + 1)) + p  ' 0 <= Rnd() < 1 => p <= i < (r - p + 1) + p = r + 1
	x = arr(i)
	arr(i) = arr(r)
	i = p - 1
	For j = p To r - 1
		If arr(j) <= x Then
			i = i + 1 ' p <= i <= j, swap j
			temp = arr(j)
			arr(j) = arr(i)
			arr(i) = temp
		End If
	Next j
	arr(r) = arr(i + 1)
	arr(i + 1) = x
	_Partition = i + 1
End Function

''
'' Return the array shuffled (Fisher Yates shuffle, a.k.a Knuth)
''
Function ShuffledArray(arr() As Variant) As Variant
	If UBound(arr) < LBound(arr) Then
		ShuffledArray = Array()
		Exit Function
	End If

	Dim i, j As Long
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
	Dim i As Long

	Dim newArr(LBound(arr) To UBound(arr)) As String
	For i=LBound(arr) To UBound(arr)
		newArr(i) = UnoValueToString(arr(i))
	Next i
	ArrayToString = Join(newArr, ", ")
End Function


''
'' Return a String of a UNOValue
''
Function UnoValueToString(value As Variant) As String
	Select Case VarType(value)
	Case V_STRING
		UnoValueToString = """" & value & """"
	Case 11
		If value Then 
			UnoValueToString = "True"
		Else
			UnoValueToString = "False"
		End If
	Case 9
		UnoValueToString = "<obj>"
	Case Else
		UnoValueToString = CStr(value)
	End Select
End Function

REM
REM Lists
REM

''
'' The ArrayList Type
'' The type of elements may be mixed.
''
Type ArrayList
	arr() As Variant
	size As Long
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
Function NewListWithCapacity(capacity As Long) As ArrayList
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
	Dim capacity As Long

	If UBound(arr) < 7 Then
		capacity = 8
	Else
		capacity = Round(UBound(arr) * 1.5)
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
Sub SetListCapacity(list As ArrayList, newCapacity As Long)
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
Sub _EnsureListCapacity(list As ArrayList, capacity As Long)
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
	Dim i, newCapacity, elementsSize As Long
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
Function GetListElement(list As ArrayList, index As Long) As Variant
	If index >= list.size Then list.arr(-1)

	GetListElement = list.arr(index)
End Function

''
'' Set the element of a list at a given index to a value
''
Sub SetListElement(list As ArrayList, index As Long, element As Variant)
	If index >= list.size Then list.arr(-1)

	list.arr(index) = element
End Sub

''
'' Insert the element at a given index
''
Sub InsertListElement(list As ArrayList, index As Long, element As Variant)
	Dim arr() As Variant
	Dim i As Long
	
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
Sub RemoveListElement(list As ArrayList, index As Long)
	Dim arr() As Variant
	Dim i As Long
	
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
Function GetListSize(list As ArrayList) As Long
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
Function ListIndexOf(list As ArrayList, element As Variant) As Long
	Dim i As Long

	For i=0 To list.size - 1
		If list.arr(i) = element Then
			ListIndexOf = i
			Exit Function
		End If
	Next i
	ListIndexOf = -1
End Function


Function ListLastIndexOf(list As ArrayList, element As Variant) As Long
	Dim i As Long

	For i=list.size-1 To 0 Step -1
		If list.arr(i) = element Then
			ListLastIndexOf = i
			Exit Function
		End If
	Next i
	ListLastIndexOf = -1
End Function

Function ListToString(list As ArrayList) As String
	Dim i As Long

	Dim newArr(0 To list.size - 1) As String
	For i=0 To list.size - 1
		newArr(i) = UnoValueToString(list.arr(i))
	Next i
	ListToString = "[" & Join(newArr, ", ") & "]"
End Function

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
	GetSetSize = GetEnumSize(e)
End Function

''
'' Return True if the Set is empty
''
Function SetIsEmpty(s As HashSet) As Variant
	Dim e As Object
	
	e = s.map.createKeyEnumeration(True)
	SetIsEmpty = Not e.hasMoreElements()
End Function

Function SetToString(s As HashSet) As String
	SetToString = "{" & ArrayToString(SetToArray(s)) & "}"
End Function


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
Function RemoveMapElement(m As HashMap, key As Variant) As Variant
	RemoveMapElement = m.map.remove(CreateUnoValue(m.keyTypeName, key))
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

''
'' Return a copy of the map
''
Function CopyMap(m As HashMap) As HashMap
	Dim newM As HashMap
	Dim e, p As Variant
	
	newM = NewEmptyMap(m.keyTypeName, m.valueTypeName)
	e = m.map.createElementEnumeration(True)
	Do While e.hasMoreElements()
		p = e.nextElement()
		PutMapElement(newM, p.First, p.Second)
	Loop
	
	CopyMap = newM
End Function


''
'' Return a merged map
''
Function MergeMaps(m1 As HashMap, m2 As HashMap) As HashMap
	Dim newM As HashMap
	Dim e, p As Variant
	
	newM = NewEmptyMap(m1.keyTypeName, m1.valueTypeName)
	e = m1.map.createElementEnumeration(True)
	Do While e.hasMoreElements()
		p = e.nextElement()
		PutMapElement(newM, p.First, p.Second)
	Loop
	e = m2.map.createElementEnumeration(True)
	Do While e.hasMoreElements()
		p = e.nextElement()
		PutMapElement(newM, p.First, p.Second)
	Loop
	
	MergeMaps = newM
End Function

Function MapToString(m As HashMap) As String
	Dim i As Long
	Dim e, p, arr As Variant
	
	e = m.map.createElementEnumeration(True)
	arr = EnumToArray(e)

	Dim newArr(LBound(arr) To UBound(arr)) As String
	For i=LBound(arr) To UBound(arr)
		p = arr(i)
		newArr(i) = UnoValueToString(p.First) & ": " & UnoValueToString(p.Second)
	Next i
	MapToString = "{" & Join(newArr, ", ") & "}"
End Function

