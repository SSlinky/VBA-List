VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' TODO: add list slicers to return a list.

'-------------------------------------------------------------------------------
'   Class: List
'   A wrapper for a Collection that gives it features of a list or stack.
'-------------------------------------------------------------------------------

' Enums
'-------------------------------------------------------------------------------
Public Enum ListMode
'   Last in, first out.
    Stack

'   First in, first out.
    Queue
End Enum


' Events
'-------------------------------------------------------------------------------


' Private Backing Store
'-------------------------------------------------------------------------------
Private mBaseCollection As Collection
Private mMode As ListMode


' Properties
'-------------------------------------------------------------------------------
Public Property Let Mode(val As ListMode)
Attribute Mode.VB_Description = "The mode affects the way the Pop function works."
'   The mode affects the way the Pop function works.
    mMode = val
End Property

Public Property Get Mode() As ListMode
    Mode = mMode
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the count for the collection."
'   Gets the count for the collection.
    Count = mBaseCollection.Count
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mBaseCollection = New Collection
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Function Item(itemIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns an item or slice at the given index."
'   Returns an item or slice at the given index.
'
'   Args:
'       itemIndex: A zero based item index or slice.
'                  Negative numbers return from end.
'                  Slices return an array, e.g., 2:7
'
'   Returns:
'       The item at the position
'
    If IsNumeric(itemIndex) Then
        Dim i As Long
        i = GetRealIndex(CLng(itemIndex))
        
        If IsObject(mBaseCollection(i)) Then
            Set Item = mBaseCollection(i)
        Else
            Item = mBaseCollection(i)
        End If
    Else
        Set Item = GetItemsBySlice(CStr(itemIndex))
    End If
End Function

Public Sub Push(val As Variant)
Attribute Push.VB_Description = "Adds an item to the list."
'   Adds an item to the list.
'
'   Args:
'       val: An expression of any type that specifies the itemm to be added.
'
    mBaseCollection.Add val
End Sub

Public Function Pop() As Variant
Attribute Pop.VB_Description = "Pops the item from the list."
'   Pops the item from the list.
'
'   Returns:
'       An item from the list (affected by Mode).
'
'   Raises:
'       An error when there are no more items to pop.
'
    Dim popIndex As Long
    popIndex = Iif(mMode = Queue, 1, mBaseCollection.Count)

    If IsObject(mBaseCollection(popIndex)) Then
        Set Pop = mBaseCollection(popIndex)
    Else
        Pop = mBaseCollection(popIndex)
    End If
    mBaseCollection.Remove(popIndex)
End Function

Public Sub Insert(val As Variant, itemIndex As Long)
Attribute Insert.VB_Description = "Inserts an item at the specified index location."
'   Inserts an item at the specified index location.
'
'   Args:
'       val: An expression of any type that specifies the item to be added.
'       itemIndex: A zero based item index. Negative index returns from end.
'
    Dim i As Long
    i = GetRealIndex(itemIndex)

    If mBaseCollection.Count = 0 Or i > mBaseCollection.Count Then
        Push val
    Else
        mBaseCollection.Add val, Before:=i
    End If
End Sub

Public Sub Remove(itemIndex As Long)
Attribute Remove.VB_Description = "Removes the item at the passed index location."
'   Removes the item at the passed index location.
'
'   Args:
'       itemIndex: A zero based item index. Negative index removes from end.
'
    Dim i As Long
    i = GetRealIndex(itemIndex)

    mBaseCollection.Remove i
End Sub

Public Function IndexOf(val As Variant) As Long
Attribute IndexOf.VB_Description = "Returns the index of the passed in item if it exists."
'   Returns the index of the passed in item if it exists.
'
'   Args:
'       val: The item to find the index of.
'
'   Returns:
'       A positive number if the item is found or negative 1 if not.
'
'   Set the default state and return if we have no items.
    IndexOf = -1
    If mBaseCollection.Count = 0 Then Exit Function

    Dim i As Long
    If IsObject(val) Then
'       Search for the object comparing memory address
        For i = 1 To mBaseCollection.Count
            If val Is mBaseCollection(i) Then
                IndexOf = i - 1
                Exit Function
            End If
        Next i
    Else
'       Search for the object comparing the value
        For i = 1 To mBaseCollection.Count
            If val = mBaseCollection(i) Then
                IndexOf = i - 1
                Exit Function
            End If
        Next i
    End If
End Function

Public Function Filter(filterCriteria As String) As List
Attribute Filter.VB_Description = "Returns a new filtered list."
'   Returns a new filtered list.
'
'   Filters the list in the same way that an ADODB.Recordset does.
'   This filter does not support complex objects due to limited
'   reflection in VBA. Objects should be implemented separately
'   so that their properties can be added to the Recordset.
'
'   Args:
'       filterCriteria: The ADODB filter string.
'       Filter on field Value, e.g., "Value LIKE 'FOO*'"
'
'   Returns:
'       A List with the objects filtered.
'
    Dim rs As Object
    Set rs = ToRecordSet(mBaseCollection)
    rs.Filter = filterCriteria

    Dim result As New List
    result.Mode = Me.Mode

    Do While Not rs.EOF
        result.Push rs.Fields!Value.Value
        rs.MoveNext
    Loop

    Set Filter = result
End Function

Public Sub Sort(sortCriteria As String)
Attribute Sort.VB_Description = "Sorts the list in place."
'   Sorts the list in place.
'
'   Sorts the list in the same way that an ADODB.Recordset does.
'   This sort does not support complex objects due to limited
'   reflection in VBA. Objects should be implemented separately
'   so that their properties can be added to the Recordset.
'
'   Args:
'       sortCriteria: The ADODB filter string.
'       Sort on field Value, e.g., "Value ASC"
'
    Dim rs As Object
    Set rs = ToRecordSet(mBaseCollection)
    rs.Sort = sortCriteria

    Set mBaseCollection = New Collection

    Do While Not rs.EOF
        Me.Push rs.Fields!Value.Value
        rs.MoveNext
    Loop
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Function GetRealIndex(itemIndex) As Long
Attribute GetRealIndex.VB_Description = "Returns the converted item index appropriate for the base collection."
'   Returns the converted item index appropriate for the base collection.
'
'   Args:
'       itemIndex: A zero based item index. Negative index returns from end.
'
'   Returns:
'       The item index, offset for the base collection.
'
    If itemIndex < 0 Then
        GetRealIndex = mBaseCollection.Count + itemIndex + 1
    Else
        GetRealIndex = itemIndex + 1
    End If
End Function

Private Function GetItemsBySlice(itemIndex As String) As List
Attribute GetItemsBySlice.VB_Description = "Returns the items at the slice."
'   Returns the items at the slice.
'
'   Args:
'       itemIndex: A zero based slice index. Negative index returns from end.
'
'   Returns:
'       The items at the slice range.
'
'   Get the beginning and end index for the slice range.
    Dim indices() As String
    indices = Split(itemIndex, ":")

    Dim begIndex As Long: begIndex = GetRealIndex(indices(0))
    Dim endIndex As Long: endIndex = GetRealIndex(indices(1))

    If begIndex > endIndex Then
        Dim tmpIndex As Long
        tmpIndex = begIndex
        begIndex = endIndex
        endIndex = tmpIndex
    End If

    Dim result As New List
    result.Mode = Me.Mode

    Dim i As Long
    For i = begIndex To endIndex
        result.Push mBaseCollection(i)
    Next i

    Set GetItemsBySlice = result
End Function

Private Function ToRecordSet(coll As Collection) As Object
Attribute ToRecordSet.VB_Description = "Converts the base Collection to a Recordset."
'   Converts the base Collection to a Recordset.
'
'   Args:
'       coll: The Collection to convert.
'
'   Returns:
'       A Recordset.
'
'   Set up the ADODB in-memory Recordset.
    Const adInteger As Long = 3
    Const adVarChar As Long = 200
    Const adLockPessimistic As Long = 2

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    rs.Fields.Append "ID", adInteger
    rs.Fields.Append "Value", adVarChar, 255
    rs.LockType = adLockPessimistic
    rs.Open

'   Add the List items to the recordset.
    Dim val As Variant
    Dim i As Long
    For Each val In mBaseCollection
        i = i + 1
        rs.AddNew
        rs!ID.Value = i
        rs!Value.Value = val
        rs.Update
    Next val

    Set ToRecordSet = rs
End Function


' Event Handlers
'-------------------------------------------------------------------------------


' Interface Implementation
'-------------------------------------------------------------------------------

