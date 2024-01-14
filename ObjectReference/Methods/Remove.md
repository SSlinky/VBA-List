# Remove method

Removes the item at the passed index location.

## Syntax

_object_.**Remove** _itemIndex_

The **Remove** method has the following parts:

Part                | Description
:---                | :---
_object_            | Required. An expression representing a **List** object.
_itemIndex_         | Required. A zero based item index. Negative index removes from end.

If the value provided as index doesn't match any existing member of the list, an error occurs.

## Example Usage

Remove the first from start (zero based) and second from end (1 based).

```vba
Dim myList As New List

With myList
    .Push "a"
    .Push "b"
    .Push "c"
    .Push "d"
    .Push "e"
End With


' List items: ("a", "b", "c", "d", "e")

myList.Remove  1  ' ("a", "c", "d", "e")
myList.Remove -2  ' ("a", "c", "e")

```
