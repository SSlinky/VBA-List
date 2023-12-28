# VBA-List

List object that extends a Collection

[![MIT license](https://img.shields.io/badge/License-MIT-blue.svg)](https://github.com/SSlinky/VBA-List/blob/master/README.md#license)
[![VBA](https://img.shields.io/badge/vba-VB--6-success)](https://docs.microsoft.com/en-us/office/vba/api/overview/)
[![Buy me a Beer!](https://img.shields.io/badge/Buy%20me%20a-Beer-yellow)](https://www.buymeacoffee.com/sslinky)

List exposes the standard functionality of a [Collection object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object) as well as providing additional useful functionality that avoids boilerplate.

* Push and Pop in Stack or Queue mode.
* IndexOf method to search list.
* Reference objects by negative index gets from end.
* Slice list by range of indices.
* Filter list by predicate.
* Sort list.

## Installation

Download the List.cls file and add it to your project.

## Licence

Released under [MIT](/LICENCE) by [Sam Vanderslink](https://github.com/SSlinky).
Free to modify and reuse.

## Documentation

[Read the docs](https://sslinky.github.io/VBA-List/#/) for usage and examples.

## Test Results

```txt
   Pass: TestList_ItemPositiveIndexReturnsItem
   Pass: TestList_ItemNegativeIndexReturnsItem
   Pass: TestList_ItemsAsObjectsReturned
   Pass: TestList_InsertInsertsAtZero
   Pass: TestList_InsertInsertsMid
   Pass: TestList_RemoveRemovesItem
   Pass: TestList_IndexOfReturnsValueIndex
   Pass: TestList_IndexOfDoesntFindValueIndex
   Pass: TestList_IndexOfDoesntFindValueNoItems
   Pass: TestList_IndexOfReturnsObjectIndex
   Pass: TestList_PushAddsItem
   Pass: TestList_PopGetsAndRemovesFromQueue
   Pass: TestList_PopGetsAndRemovesFromStack
   Pass: TestList_ItemGetsItemsSlice
   Pass: TestList_FilterFiltersStrings
   Pass: TestList_FilterFiltersNumbers
   Pass: TestList_SortListSortsStrings
   Pass: TestList_SortListSortsNumbers
-------------------------------------------
   Passed: 18 (100.00%)
   Failed: 0 (0.00%)
-------------------------------------------
```
