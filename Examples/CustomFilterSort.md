# Extending the Filter and Sort

Sorting and filtering make use of an ADODB Recordset to work. The basic List class only provides support for simple filtering and sorting, i.e., converts all items to a Recordset with a Value field. This is due to VBA's lack of Reflection.

To sort or filter on the properties of a class, you'll need to implement your own `ToRecordSet` helper method.

## Example Scenario

We have a list of `Customer` objects. Corporate have decided, in their wisdom, to hand out gold star stickers as a way of showing thier thanks. They only have 500 stickers so they have instructed you to do this in order of region and total sales. Just go with it.

## ToCustomerRecordSet Helper

Copy the existing method and modify it so that it records the `Region` and `TotalSales` properties.

```vba
Private Function ToCustomerRecordSet( _
    coll As Collection, Optional asNumeric As Boolean) As Object
'   Converts the base Customer Collection to a Recordset.
'
'   Args:
'       coll: The Collection to convert.
'       asNumeric: Uses numbers for the values.
'
'   Returns:
'       A Recordset.
'
'   Set up the ADODB in-memory Recordset.
    Const adInteger As Long = 3
    Const adVarChar As Long = 200
    Const adVarNumeric As Long = 139
    Const adLockPessimistic As Long = 2

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    rs.Fields.Append _
        "ID", _
        adInteger
' ########## Modify this part ##########
    rs.Fields.Append "Region", adVarChar, 255
    rs.Fields.Append "TotalSales", adVarNumeric, 255
' #####################################
    rs.LockType = adLockPessimistic
    rs.Open

'   Add the List items to the recordset.
    Dim val As Variant
    Dim cust As Customer
    Dim i As Long
    For Each val In mBaseCollection
        Set cust = val
        i = i + 1
        rs.AddNew
        rs!ID.Value = i
' ########## Modify this part ##########
        rs!Region.Value = cust.Region
        rs!TotalSales.Value = cust.TotalSales
' #####################################
        rs.Update
    Next val

    Set ToCustomerRecordSet = rs
End Function
```

## Sort Method

You could implement this any fancy way you wished, but for simplicity, we're just going to create another sort method. All we're changing here is the helper that generates the Recordset.

```vba

Public Sub SortCustomers(sortCriteria As String)
'   Sorts the customer list in place.
'
'   Sorts the list in the same way that an ADODB.Recordset does.
'   This sort does not support complex objects due to limited
'   reflection in VBA. Objects should be implemented separately
'   so that their properties can be added to the Recordset.
'
'   Args:
'       sortCriteria: The ADODB filter string.
'       Sort on field Value, e.g., "Region ASC"
'
    Dim rs As Object
' ########## Modify this part ##########
    Set rs = ToCustomerRecordSet( _
        mBaseCollection, _
        AllValuesNumeric(mBaseCollection))
' #####################################
    rs.Sort = sortCriteria

    Set mBaseCollection = New Collection

    Do While Not rs.EOF
        Me.Push rs.Fields!Value.Value
        rs.MoveNext
    Loop
End Sub
```

## Putting It Into Practice

```vba
Function GetEligibleCustomers(customers As List, stickers As Long) As Variant()
'   Returns the customers eligible for a gold star sticker.
'
'   Args:
'       customers: The list of customers.
'       stickers: The maximum customers.
'
    customers.SortCustomers "Region ASC, ToatlSales DESC"

    With customers
        If .Count > stickers Then
            GetEligibleCustomers = .Item("0:" & stickers - 1)
        Else
            GetEligibleCustomers = .Item("0:" & .Count - 1)
        End If
    End With
End
```
