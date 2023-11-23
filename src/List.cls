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
    Pipe
End Enum


' Events
'-------------------------------------------------------------------------------


' Private Backing Store
'-------------------------------------------------------------------------------
Private mBaseCollection As Collection
Private mMode As ListMode


' Properties
'-------------------------------------------------------------------------------
Public Property Let Mode(var As ListMode)
Attribute Mode.VB_Description = "The mode affects the way the Pop function works."
'   The mode affects the way the Pop function works.
    mMode = var
End Property

Public Property Get Mode() As ListMode
    Mode = mMode
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mBaseCollection = New Collection
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Function Item(Optional itemIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns the item at the position."
'   Returns the item at the position.
'
'   Args:
'       itemIndex: A zero based item index. Negative index returns from end.
'
'   Returns:
'       The item at the position
'
    Dim i As Long
    i = GetRealIndex(itemIndex)
    
    If IsObject(mBaseCollection(i)) Then
        Set Item = mBaseCollection(i)
    Else
        Item = mBaseCollection(i)
    End If
End Function

Public Sub Push(Item As Variant)
Attribute Push.VB_Description = "Adds an item to the list."
'   Adds an item to the list.
'
'   Args:
'       Item: An expression of any type that specifies the itemm to be added.
'
    mBaseCollection.Add Item
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
    popIndex = Iif(mMode = Pipe, 0, mBaseCollection.Count)

    If IsObject(mBaseCollection(popIndex)) Then
        Set Pop = mBaseCollection(popIndex)
    Else
        Pop = mBaseCollection(popIndex)
    End If
    mBaseCollection.Remove(popIndex)
End Function

Public Sub Insert(Item As Variant, itemIndex As Long)
Attribute Insert.VB_Description = "Inserts an item at the specified index location."
'   Inserts an item at the specified index location.
'
'   Args:
'       Item: An expression of any type that specifies the itemm to be added.
'       itemIndex: A zero based item index. Negative index returns from end.
'
    Dim i As Long
    i = GetRealIndex(itemIndex)

    If mBaseCollection.Count = 0 Or i > mBaseCollection.Count Then
        Push Item
    Else
        mBaseCollection.Add Item, Before:=i
    End If
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


' Event Handlers
'-------------------------------------------------------------------------------


' Interface Implementation
'-------------------------------------------------------------------------------

