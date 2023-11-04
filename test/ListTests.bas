Attribute VB_Name = "ListTests"
' Copyright 2023 Sam Vanderslink
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy 
' of this software and associated documentation files (the "Software"), to deal 
' in the Software without restriction, including without limitation the rights 
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
' copies of the Software, and to permit persons to whom the Software is 
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in 
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
' IN THE SOFTWARE.

Option Explicit

Private passTests As New Collection
Private failTests As New Collection

Public Sub RunTests()
Attribute RunTests.VB_Description = "Runs all tests."
'   Runs all tests.
'
    Set passTests = New Collection
    Set failTests = New Collection

    Dim testName As Variant
    For Each testName In GetTestNames()
        RunTest CStr(testName)
    Next testName

    Dim p As Long, f As Long
    p = passTests.Count
    f = failTests.Count

    Debug.Print "-------------------------------------------"
    Debug.Print "   Passed: " & p & " (" & Format(p / (p + f), "0.00%)")
    Debug.Print "   Failed: " & f & " (" & Format(f / (p + f), "0.00%)")
    Debug.Print "-------------------------------------------"
    
End Sub

Sub RunSingle()
    Dim tr As TestResult
    Set tr = TestDictionary_ItemReturnsItem()
    Debug.Print tr.ToString
End Sub

Private Sub RunTest(testName As String)
Attribute RunTest.VB_Description = "Runs the named test and stores the result."
'   Runs the named test and stores the result.
'
'   Args:
'       testName: The name of the function returning a TestResult.
'
    Dim tr As TestResult
    Set tr = Application.Run(testName)
    tr.Name = testName
    Debug.Print tr.ToString

    If tr.Failed Then failTests.Add tr Else passTests.Add tr
End Sub

Private Function GetTestNames() As Collection
Attribute GetTestNames.VB_Description = "Gets the test names from this module."
'   Gets the test names from this module.
'   A valid test starts with Private Function TestList_ and takes no args.
'
'   Returns:
'       A collection of strings representing names of tests.
'
    Const MODULENAME As String = "ListTests"
    Const FUNCTIONID As String = "Private Function "
    Const TESTSTARTW As String = "Private Function TestList_"

    Dim tswLen As Long
    tswLen = Len(TESTSTARTW)

    Dim codeMod As Object
    Set codeMod = ThisWorkbook.VBProject.VBComponents(MODULENAME).CodeModule

    Dim i As Long
    Dim results As New Collection
    For i = 1 To codeMod.CountOfLines
        Dim lineContent As String
        lineContent = codeMod.Lines(i, 1)

        If Left(lineContent, tswLen) = TESTSTARTW Then
            Dim funcName As String
            funcName = Split(Split(lineContent, FUNCTIONID)(1), "(")(0)
            results.Add funcName
        End If
    Next i
    
    Set GetTestNames = results
End Function
