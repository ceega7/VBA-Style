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

* When using blocks, indent where appropriate

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
