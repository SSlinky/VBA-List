# Mode Property

Returns or sets the List mode. The mode affects the way the Pop function works.

## Syntax

_object_.**DataCols** (_val_)

Part                | Description
:---                | :---
_object_            | Required. An expression representing a **List** object.
_val_               | Required. The `ListMode` to operate in.

## Remarks

This defines the Mode as either `ListMode.Stack` or `ListMode.Queue`. By default, a List is in Stack mode.

Stack and Queue pop items from the top and bottom repectively, i.e.;

- Stack operates as last in, first out.
- Queue operates as first in, first out.
