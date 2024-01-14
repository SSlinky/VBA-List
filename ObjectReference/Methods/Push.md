# Push method

Adds a item to the List.

## Syntax

_object_.**Push** _Val_

The **Push** method has the following parts:

Part               | Description
:---               | :---
_object_           | Required. An expression representing a **List** object.
_Val_              | Required. An expression of any type that specifies the item to be added.

## Example Usage

Push three items to the list.

```vba
Dim myList As New List
With myList
    .Push "a"
    .Push "b"
    .Push "c"
End With
```
