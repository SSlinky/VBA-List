# IndexOf method

Returns the index of the passed in item if it exists. The returned number will be positive number if the item is found or -1 if not.

## Syntax

_object_.**IndexOf** _val_

The **IndexOf** method has the following parts:

Part               | Description
:---               | :---
_object_           | Required. An expression representing a **List** object.
_val_              | Required. The item we're attempting to get the index of.

## Example Usage

Find the index for some values.

```vba
Dim myList As New List

With myList
    .Push "a"
    .Push "b"
    .Push "c"
End With

Debug.Print "IndexOf(""a"") = " & myList.IndexOf("a") '  0
Debug.Print "IndexOf(""b"") = " & myList.IndexOf("b") '  1
Debug.Print "IndexOf(""c"") = " & myList.IndexOf("c") '  2
Debug.Print "IndexOf(""z"") = " & myList.IndexOf("z") ' -1

```
