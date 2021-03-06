# VBA-Style
Visual Basic for Applications - Style Guide

##Settings

* Always use ```Option Explicit```

> Why? Using this avoids unnecessary reference errors and makes your code cleaner by convention 

```vba
' NOT Option Explicit

Dim MyNumber as Long, MyResult as Long

MyNumber = 1

MyResult = MyNumber + 1 ' MyResult = 2
MyResult = MyMunber + 1 ' MyResult = 1 -> MyNumber spelt incorrectly

' with Option Explicit enabled this would result in a Compile Error as MyMunber is not referenced
```

###Variables

* Declare all module variables at the top of the procedure

> Why? It makes code refactoring much easier

```vba

'good
Dim a As Long, b As Long

a = 1
b = 2

Debug.Print a + b

'bad
Dim a As Long

a = 1

Dim b As Long

Debug.Print a + b
```

* Use an appropriate data type instead of allowing VBA to automatically declare as variant for you

> Why? Variants consume more storage space and - while their flexibility can be an advantage - it may cause errors in your code if a specific type is not declared

```vba

'good
Dim i as Integer
Dim myStr as String
Dim ws as Worksheet

'bad
Dim i, myStr, ws

```

* Declare variables of the same type in blocks where possible

> Why? Readability is greatly improved

```vba

'good
Dim aLg As Long, bLg As Long, cLg As Long
Dim aWs As Worksheet, bWs As Worksheet, cWs As Worksheet
Dim aSt As String, bSt As String, cSt As String

'bad
Dim aLg As Long, cSt As String, aWs As Worksheet
Dim aSt As String, cWs As Worksheet, cLg As Long
Dim bWs As Worksheet, bLg As Long, bSt As String

'awful
Dim aLg As Long
Dim bLg As Long
Dim cLg As Long
Dim aWs As Worksheet
Dim bWs As Worksheet
Dim cWs As Worksheet
Dim aSt As String
Dim bSt As String
Dim cSt As String
``` 

* Don't reuse variables

> Why? Reusing variables can be confusing. To avoid errors, just create another variable where needed

* Use ```Const``` for all static references

> Why? This ensures references are not errantly reassigned. Referencing is also much easier; if the variable needs to be changed throughout the project, it only needs to be changed in a single place

```vba
' good
Public Const myFile = "C:\Users\Me\Folder\This.xlsb"
```

* Avoid using ```Public``` variables where they are not needed

> Why? Though global variables are useful, module scoped variables are easy to manage and the chance of errantly overwriting variables decreases

##Style

* Avoid overlong, clunky code and aim compartmentalise procedures into smaller Functions and Sub Procedures. If there is an action which you will repeat over time, create a function for reuse.

> Why? Long blocks of code can be hard to understand when read back. Breaking code into smaller chunks allows you to repeat procedures and minimise the length of code blocks

```vba

'good
Sub myModule()
Dim ws As Worksheet
Dim boo As Boolean

boo = worksheetExists("Data")

End Sub

Function worksheetExists(ByRef wsName As String) As Boolean
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
  If ws.Name = wsName Then
    worksheetExists = True
      Exit Function
  End If
Next

worksheetExists = False

End Function

'bad
Sub myModule()
Dim ws As Worksheet
Dim boo As Boolean

For Each ws In ThisWorkbook.Worksheets
  If ws.Name = "Data" Then
    boo = True
  End If
Next

End Sub
```

* Where appropriate, seperate statements with line breaks and attempt to group them logically

> Why? Purely for readability

```vba

'good -> group together similar statements
ws.range("A1").value = 1
ws.range("Z10").value = "Hello"
' Line Break Here!
ThisWorkbook.Close savechanges:=True

'bad -> all grouped together
ws.range("A1").value = 1
ws.range("Z10").value = "Hello"
ThisWorkbook.Close savechanges:=True
```

* Don't use a trailing ```Else``` if no clause is specified

> Why? They are superflous. An ```End If``` block exit is more efficient and easier to read

```vba

'good
If ref = "Err" then
  Exit Sub
End If

'bad
If ref = "Err" Then
  Exit Sub
Else
End If
```

* Always indent where appropriate

> Why? Indentation makes your code much more readable

```vba 

'good
For Each wb In Application.Workbooks
    If wb.FullName = "C:\Users\Me\Folder\MyWorkbook.xlsb" Then
        MsgBox "Found"
    End If
Next

'bad
For Each wb In Application.Workbooks
If wb.FullName = "C:\Users\Me\Folder\MyWorkbook.xlsb" Then
MsgBox "Found"
End If
Next

'awful
For Each wb In Application.Workbooks
For Each ws In wb.Worksheets
  Debug.Print ws.Name
Next
Next
```

* When testing for truthy or falsey values don't use ```True``` or ```False```

> Why? They are superfluous. ```If``` and ```If Not```, are shorter and read better 

```vba

'good
If foo Then
  Debug.Print "Yes"
End if

If Not foo Then
  Debug.Print "No"
End If

'bad 
If foo = True Then
  Debug.Print "Yes"
ElseIf foo = False Then
  Debug.print "No"
End If

```

* Avoid using ```GoTo```

> Why? It's bad practice and lazy. More often than not there will be a better way to code around it. If you must use it alongside ```On Error```, immediately use ```On Error GoTo 0``` to de-register any created Error Blocks 

```vba

'good
Dim a() As Variant, v As Variant

a = Array(1, 2, 3, 4, "Five")

For Each v In a
  If IsNumeric(v) Then
    Debug.Print v + 10
  End If
Next

'bad
Dim a() As Variant, v As Variant

a = Array(1, 2, 3, 4, "Five")

For Each v In a
  On Error Resume Next
  Debug.Print v + 10
Next

```

* Use Arrays

> Why? Arrays in VBA are extremely useful for creating efficent, ordered code

```vba

' useful Array functions

' inArray

Function inArray(ByRef arr() As Variant, ByVal item As Variant) As Boolean

' -> searches through the items in an array and returns True if the argument item is found

Dim i As Long

inArray = False

For i = LBound(arr()) To UBound(arr())
  If UCase(arr(i)) = UCase(item) Then
    inArray = True
  End If
Next

End Function

' usage

Sub myModule()
Dim a() As Variant

a = Array("Robin", "Monica", "Ralph", "Eva", "Omar")

Debug.Print inArray(a, "Robin") ' -> True
Debug.Print inArray(a, "Bruce") ' -> False

End Sub

' aPush

Function aPush(ByRef arr() As Variant, ByVal Value as Variant)

' -> ReDims a dynamic array's boundaries and adds the selected item to the end of it

Dim i As Long

If Not IsArray(Value) Then
  ReDim Preserve arr(LBound(arr()) To UBound(arr()) + 1)
    arr(UBound(arr)) = Value
Else
  For i = LBound(Value) To UBound(Value)
    ReDim Preserve arr(LBound(arr()) To UBound(arr()) + 1)
    arr(UBound(arr)) = Value(i)
  Next
End If

End Function

Sub myModule()
Dim ws As Worksheet
Dim a() As Variant
Dim cell As Range
Dim v As Variant

Set ws = ThisWorkbook.Sheets("Data")

a = Array()

For Each cell In ws.Range("A1:A5")
  aPush a, cell.Value
Next

' the Array a() now contains the values from Range("A1:A5")

End Sub

```

* Use Named Parameters where appropriate

> Why> Named parameters make code much easier to read

```vba

'good
ThisWorkbook.Names.Add Name:="MyRange", RefersTo:=ThisWorkbook.Worksheets("Data").Range("A1:A5"), Visible:=True

'bad
ThisWorkbook.Names.Add "MyRange", ThisWorkbook.Worksheets("Data").Range("A1:A5"), True

```

##Sub Procedures

* Don't overcomplicate the names of sub-procedures. Keep them short, descriptive and use camelCase if appropriate

> Why? Readability is key to refactoring and improvement

```vba 

'good
Sub countTwo()
  '''
End Sub

'bad
Sub ThisMacroOpensTwoWorkbooksThenRunsAnotherMacro()
  '''
End Sub
```

##Functions

* 

##Workbook

* Always use a variable to refer to ```Workbook```s you are referencing

> Why? This allows you to have clear, complete control over the objects you are using
 
```vba
' good
Dim wb as Workbook

Set wb = Workbooks.open(fileName, ReadOnly:=True)

' bad
Workbooks.open fileName

ActiveWorkbook.Activate

```

* Use ```ThisWorkbook``` rather than ```ActiveWorkbook```

> Why? This avoids confusion when creating procedures which run across multiple workbooks
