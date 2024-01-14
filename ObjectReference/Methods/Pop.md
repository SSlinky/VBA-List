# Pop method

Removes and returns an item from the List.

## Syntax

_object_.**Pop** _Val_

The **Pop** method has the following parts:

Part               | Description
:---               | :---
_object_           | Required. An expression representing a **List** object.

Whether the item is popped from the start or end of the list depends on the [Mode](ObjectReference/Properties/Mode.md "VBA-List - Properties - Mode").

## Example Usage

Pop all items from the list.

```vba
With myList
    While .Count > 0
        myVar = myList.Pop()
    Loop
End With
```

_Note:_ To pop an object, the `Set` keyword is required.
