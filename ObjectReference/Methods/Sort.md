# Sort method

Sorts the list in place.

## Syntax

_object_.**Sort** _sortCriteria_, _[opt arg1]_

The **Sort** method has the following parts:

Part                | Description
:---                | :---
_object_            | Required. An expression representing a **List** object.
_sortCriteria_      | Required. The ADODB filter string.

Sorts the list in the same way that an ADODB.Recordset does. This sort does not support complex objects due to limited reflection in VBA. Specific object sorts should be implemented separately so that their properties can be added to the Recordset.

The Recordset field to sort on is "Value".

## Example Usage

Sort the items by ascending and descending.

```vba
Dim myList As New List

With myList
    .Push 5
    .Push 3
    .Push 2
    .Push 1
    .Push 6
End With

                         ' List items: (5, 3, 2, 1, 6)
myList.Sort "Value ASC"  ' List items: (1, 2, 3, 5, 6)
myList.Sort "Value DESC" ' List items: (6, 5, 3, 2, 1)
```
