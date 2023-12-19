Attribute VB_Name = "ListTests"
' Copyright 2023 Sam Vanderslink
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy 
' of this software and associated documentation files (the "Software"), to deal 
' in the Software without restriction, including without limitation the rights 
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
' copies of the Software, and to permit persons to whom the Software is 
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in 
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
' IN THE SOFTWARE.

Option Explicit

Private passTests As New Collection
Private failTests As New Collection

Public Sub RunTests()
Attribute RunTests.VB_Description = "Runs all tests."
'   Runs all tests.
'
    Set passTests = New Collection
    Set failTests = New Collection

    Dim testName As Variant
    For Each testName In GetTestNames()
        RunTest CStr(testName)
    Next testName

    Dim p As Long, f As Long
    p = passTests.Count
    f = failTests.Count

    Debug.Print "-------------------------------------------"
    Debug.Print "   Passed: " & p & " (" & Format(p / (p + f), "0.00%)")
    Debug.Print "   Failed: " & f & " (" & Format(f / (p + f), "0.00%)")
    Debug.Print "-------------------------------------------"
    
End Sub

Sub RunSingle()
    Dim tr As TestResult
    Set tr = TestList_IndexOfReturnsValueIndex()
    Debug.Print tr.ToString
End Sub

Private Sub RunTest(testName As String)
Attribute RunTest.VB_Description = "Runs the named test and stores the result."
'   Runs the named test and stores the result.
'
'   Args:
'       testName: The name of the function returning a TestResult.
'
    Dim tr As TestResult
    Set tr = Application.Run(testName)
    tr.Name = testName
    Debug.Print tr.ToString

    If tr.Failed Then failTests.Add tr Else passTests.Add tr
End Sub

Private Function GetTestNames() As Collection
Attribute GetTestNames.VB_Description = "Gets the test names from this module."
'   Gets the test names from this module.
'   A valid test starts with Private Function TestList_ and takes no args.
'
'   Returns:
'       A collection of strings representing names of tests.
'
    Const MODULENAME As String = "ListTests"
    Const FUNCTIONID As String = "Private Function "
    Const TESTSTARTW As String = "Private Function TestList_"

    Dim tswLen As Long
    tswLen = Len(TESTSTARTW)

    Dim codeMod As Object
    Set codeMod = ThisWorkbook.VBProject.VBComponents(MODULENAME).CodeModule

    Dim i As Long
    Dim results As New Collection
    For i = 1 To codeMod.CountOfLines
        Dim lineContent As String
        lineContent = codeMod.Lines(i, 1)

        If Left(lineContent, tswLen) = TESTSTARTW Then
            Dim funcName As String
            funcName = Split(Split(lineContent, FUNCTIONID)(1), "(")(0)
            results.Add funcName
        End If
    Next i
    
    Set GetTestNames = results
End Function

Private Function TestList_ItemPositiveIndexReturnsItem() As TestResult
Attribute TestList_ItemPositiveIndexReturnsItem.VB_Description = "The correct item at a postive index is returned."
'   The correct item at a postive index is returned.
    Dim tr As New TestResult

'   Arrange
    Dim items As Variant
    items = Array("a", "b", "c")

    Dim foo As New List
    Dim i As Long
    For i = 0 To UBound(items)
        foo.Push items(i)
    Next i

'   Act Assert
    For i = 0 To UBound(items)
        If tr.AssertAreEqual(items(i), foo(i), CStr(i)) Then GoTo Finally
    Next i

Finally:
    Set TestList_ItemPositiveIndexReturnsItem = tr
End Function

Private Function TestList_ItemNegativeIndexReturnsItem() As TestResult
Attribute TestList_ItemNegativeIndexReturnsItem.VB_Description = "The correct item at a negative index is returned."
'   The correct item at a negative index is returned.
    Dim tr As New TestResult

'   Arrange
    Dim items As Variant
    items = Array("a", "b", "c")

    Dim myList As New List
    Dim i As Long
    For i = 0 To UBound(items)
        myList.Push items(i)
    Next i

'   Act Assert
    Dim neg As Long
    Dim exp As String
    Dim act As String
    Dim msg As String

    For i = 0 To UBound(items)
        neg = i - UBound(items) - 1
        exp = items(i)
        act = myList(neg)
        msg = i & " = " & neg & " for count " & UBound(items) + 1 & "."
        If tr.AssertAreEqual(exp, act, msg) Then GoTo Finally
    Next i

Finally:
    Set TestList_ItemNegativeIndexReturnsItem = tr
End Function

Private Function TestList_ItemsAsObjectsReturned() As TestResult
Attribute TestList_ItemsAsObjectsReturned.VB_Description = "An item (requires Set keyword) as an object is returned."
'   An item (requires Set keyword) as an object is returned.
    Dim tr As New TestResult

'   Arrange
    Dim foo As TestResult
    Set foo = New TestResult
    Dim bar As TestResult

    Dim myList As New List
    myList.Push foo

'   Act
    Set bar = myList.Pop()

'   Assert
    If tr.AssertIs(foo, bar, "foo and bar") Then GoTo Finally

Finally:
    Set TestList_ItemsAsObjectsReturned = tr
End Function

Private Function TestList_InsertInsertsAtZero() As TestResult
Attribute TestList_InsertInsertsAtZero.VB_Description = "Insert can insert at the first element."
'   Insert can insert at the first element.
    Dim tr As New TestResult

'   Arrange
    Dim items As Variant
    items = Array("a", "b", "c")

    Dim myList As New List
    Dim i As Long
    For i = 0 To UBound(items)
        myList.Push items(i)
    Next i

'   Act
    myList.Insert "x", 0

'   Assert
    If tr.AssertAreEqual("x", myList(0), "0") Then Goto Finally
    For i = 1 To UBound(items) + 1
        If tr.AssertAreEqual(items(i - 1), myList(i), CStr(i)) Then GoTo Finally
    Next i


Finally:
    Set TestList_InsertInsertsAtZero = tr
End Function

Private Function TestList_InsertInsertsMid() As TestResult
Attribute TestList_InsertInsertsMid.VB_Description = "Insert can insert in the middle."
'   Insert can insert in the middle.
    Dim tr As New TestResult

'   Arrange
    Dim items As Variant
    items = Array("a", "b", "c")

    Dim myList As New List
    Dim i As Long
    For i = 0 To UBound(items)
        myList.Push items(i)
    Next i

    Const INSERTEDVAL As String = "xxx"
    Const INSERTEDLOC As Long = 1

'   Act
    myList.Insert INSERTEDVAL, INSERTEDLOC

'   Assert
    For i = 0 To UBound(items)
        Dim j As Long
        j = Iif(i < INSERTEDLOC, i, i + 1)
        If tr.AssertAreEqual(items(i), myList(j)) Then GoTo Finally
    Next i

    tr.AssertAreEqual INSERTEDVAL, myList(INSERTEDLOC)

Finally:
    Set TestList_InsertInsertsMid = tr
End Function

Private Function TestList_RemoveRemovesItem() As TestResult
Attribute TestList_RemoveRemovesItem.VB_Description = "Remove can remove an item."
'   Remove can remove an item.
    Dim tr As New TestResult

'   Arrange
    Dim items As Variant
    items = Array("a", "b", "c")

    Dim myList As New List
    Dim i As Long
    For i = 0 To UBound(items)
        myList.Push items(i)
    Next i

    Const REMOVEDLOC As Long = 1

'   Act
    myList.Remove REMOVEDLOC

'   Assert
    For i = 0 To UBound(items)
        If i = REMOVEDLOC Then i = i + 1

        Dim j As Long
        j = Iif(i < REMOVEDLOC, i, i - 1)
        If tr.AssertAreEqual(items(i), myList(j)) Then GoTo Finally
    Next i

Finally:
    Set TestList_RemoveRemovesItem = tr
End Function

Private Function TestList_IndexOfReturnsValueIndex() As TestResult
Attribute TestList_IndexOfReturnsValueIndex.VB_Description = "IndexOf finds the index when passed a value."
'   IndexOf finds the index when passed a value.
    Dim tr As New TestResult

'   Arrange
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    Dim myList As New List

'   Act
'   Add random unique characters to the list. Track items added
'   with a dictionary so we know their position and that they are unique.
    Dim i As Long, j As Long
    For i = 0 To 100
        Dim randChar As String * 1
        randChar = Chr(tr.GetRandomBetween(65, 122))
        If Not items.Exists(randChar) Then
            items.Add randChar, j
            myList.Push randChar
            j = j + 1
        End If
    Next i

'   Assert
    Dim key As Variant
    For Each key In items
        If tr.AssertAreEqual(items(key), myList.IndexOf(key)) Then GoTo Finally
    Next key

Finally:
    Set TestList_IndexOfReturnsValueIndex = tr
End Function

Private Function TestList_IndexOfDoesntFindValueIndex() As TestResult
Attribute TestList_IndexOfDoesntFindValueIndex.VB_Description = "IndexOf returns -1 if the value doesn't exist."
'   IndexOf returns -1 if the value doesn't exist.
    Dim tr As New TestResult

'   Arrange
    Dim items As Variant
    items = Array("a", "b", "c")

    Dim myList As New List
    Dim i As Long
    For i = 0 To UBound(items)
        myList.Push items(i)
    Next i

'   Act
    Dim result As Long
    result = myList.IndexOf("x")

'   Assert
    tr.AssertAreEqual result, -1

Finally:
    Set TestList_IndexOfDoesntFindValueIndex = tr
End Function

Private Function TestList_IndexOfDoesntFindValueNoItems() As TestResult
Attribute TestList_IndexOfDoesntFindValueNoItems.VB_Description = "IndexOf returns -1 when there are no items."
'   IndexOf returns -1 when there are no items.
    Dim tr As New TestResult

'   Arrange
    Dim myList As New List

'   Act
    Dim result As Long
    result = myList.IndexOf("x")

'   Assert
    tr.AssertAreEqual result, -1

Finally:
    Set TestList_IndexOfDoesntFindValueNoItems = tr
End Function

Private Function TestList_IndexOfReturnsObjectIndex() As TestResult
Attribute TestList_IndexOfReturnsObjectIndex.VB_Description = "IndexOf finds the index when passed an object."
'   IndexOf finds the index when passed an object.
    Dim tr As New TestResult

'   Arrange
    Dim items As Object: Set items = CreateObject("Scripting.Dictionary")
    Dim indices As Object: Set indices = CreateObject("Scripting.Dictionary")
    
    Dim myList As New List

'   Act
'   Add random unique characters to the list. Track items added and indices
'   with a dictionary so we know their position and that they are unique.
    Dim i As Long, j As Long
    For i = 0 To 100
        Dim randChar As String * 1
        Dim obj As Collection
        randChar = Chr(tr.GetRandomBetween(65, 122))
        If Not items.Exists(randChar) Then
            Set obj = New Collection
            items.Add randChar, obj
            indices.Add randChar, j
            myList.Push obj
            j = j + 1
        End If
    Next i

'   Assert
    Dim key As Variant
    For Each key In items
        Dim result As Long
        result = myList.IndexOf(items(key))
        If tr.AssertAreEqual(indices(key), result) Then GoTo Finally
    Next key

Finally:
    Set TestList_IndexOfReturnsObjectIndex = tr
End Function

Private Function TestList_PushAddsItem() As TestResult
Attribute TestList_PushAddsItem.VB_Description = "Push adds an item to the list."
'   Push adds an item to the list.
    Dim tr As New TestResult

'   Arrange
    Dim items() As Variant
    items = Array(1, 2, 3)

    Dim myList As New List

'   Act
    Dim val As Variant
    For Each val In items
        myList.Push val
    Next val

'   Assert
    For Each val In items
        If tr.AssertAreNotEqual(myList.IndexOf(val), -1) Then GoTo Finally
    Next val


Finally:
    Set TestList_PushAddsItem = tr
End Function

Private Function TestList_PopGetsAndRemovesFromQueue() As TestResult
Attribute TestList_PopGetsAndRemovesFromQueue.VB_Description = "Pop gets the first-in item from the queue."
'   Pop gets the first-in item from the queue.
    Dim tr As New TestResult

'   Arrange
    Dim items() As Variant
    items = Array(1, 2, 3)

    Dim myList As New List
    myList.Mode = Queue
    
    Dim val As Variant
    For Each val In items
        myList.Push val
    Next val

'   Act / Assert
    Dim i As Long
    For i = LBound(items) To UBound(items)
        Dim res As Variant
        res = myList.Pop()
        If tr.AssertAreEqual(res, items(i), "value") Then GoTo Finally
        If tr.AssertAreEqual(myList.Count, UBound(items) - i, "count") Then GoTo Finally
    Next i

Finally:
    Set TestList_PopGetsAndRemovesFromQueue = tr
End Function

Private Function TestList_PopGetsAndRemovesFromStack() As TestResult
Attribute TestList_PopGetsAndRemovesFromStack.VB_Description = "Pop gets the last-in item from the stack."
'   Pop gets the last-in item from the stack.
    Dim tr As New TestResult

'   Arrange
    Dim items() As Variant
    items = Array(1, 2, 3)

    Dim myList As New List
    myList.Mode = Stack

    Dim val As Variant
    For Each val In items
        myList.Push val
    Next val

'   Act / Assert
    Dim i As Long
    For i = UBound(items) To LBound(items) Step -1
        Dim res As Variant
        res = myList.Pop()
        If tr.AssertAreEqual(res, items(i), "value") Then GoTo Finally
        If tr.AssertAreEqual(myList.Count, i, "count") Then GoTo Finally
    Next i

Finally:
    Set TestList_PopGetsAndRemovesFromStack = tr
End Function
