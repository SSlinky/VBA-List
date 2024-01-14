# Item method

Returns an item or slice at the given index.

## Syntax

_object_.**Item** _itemIndex_

The **Item** method has the following parts:

Part                | Description
:---                | :---
_object_            | Required. An expression representing a **List** object.
_itemIndex_         | Required. A zero based item index or slice. Negative numbers return from end. Slices return an array, e.g., 2:7

If the value provided as index doesn't match any existing member of the list, an error occurs. The Item method is the default method for a list. Therefore, the following lines of code are equivalent:

```vba
    var = myList(1)
    var = myList.Item(1)
```

## Example Usage

Get some items from a list.

```vba
Dim myList As New List

With myList
    .Push "a"
    .Push "b"
    .Push "c"
    .Push "d"
    .Push "e"
End With

Debug.Print "Item  0:", myList(0)          ' a
Debug.Print "Item -1:", myList(-1)         ' c
Debug.Print "Item  2 to 4:", myList("2:4") ' (c, d, e)
Debug.Print "Item  3:", myList(5)          ' Run-time error '9': Subscript out of range.

```
