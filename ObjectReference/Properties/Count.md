# Count property

Gets the count for the collection.

## Syntax

_object_.**Count** _req arg1_, _[opt arg1]_

The **Count** method has the following parts:

Part                | Description
:---                | :---
_object_            | Required. An expression representing a **List** object.

## Example Usage

Get the count of items added.

```vba
Dim myList As New List

With myList
    .Push 1
    .Push 2
    .Push 3
    .Push 4
    .Push 5
End With

Debug.Print "Item count: " & myList.Count ' 5
```
