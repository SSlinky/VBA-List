# Insert method

Inserts an item at the specified index location.

## Syntax

_object_.**Insert** _val_, _itemIndex_

The **Insert** method has the following parts:

Part               | Description
:---               | :---
_object_           | Required. An expression representing a **List** object.
_val_              | Required. An expression of any type that specifies the item to be added.
_itemIndex_        | Required. A zero based item index. Negative index inserts from end.

Supplying a negative item index will insert at the nth position from the end.

## Example Usage

Insert two items.

```vba
Dim myList As New List

With myList
    .Push "a"
    .Push "b"
    .Push "c"
End With

' List items: ("a", "b", "c")

myList.Insert "x", 0  ' ("x", "a", "b", "c")
myList.Insert "y", -1 ' ("x", "a", "b", "c", "y")
myList.Insert "z", -2 ' ("x", "a", "b", "c", "z", "y")
```
