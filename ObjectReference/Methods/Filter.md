# Filter method

Returns a new filtered list.

## Syntax

_object_.**Filter** _filterCriteria_

The **Filter** method has the following parts:

Part               | Description
:---               | :---
_object_           | Required. An expression representing a **List** object.
_filterCriteria_   | Required. An ADODB filter string.

Filters the list in the same way that an ADODB.Recordset does. This filter does not support complex objects due to limited reflection in VBA. Objects should be implemented separately so that their properties can be added to the Recordset.

For more information, see Microsoft's documentation for [ADO Filter](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/filter-property-ado).

## Example Usage

Return a new filtered list where values begin with Foo.

```vba
Dim filteredList As List
Set filteredList = originalList.Filter("Value LIKE 'Foo*'")
```
