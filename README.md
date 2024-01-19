# VBA Style Guide

This is a guide on how to format VBA code so that it is easy to read. It will also make it easy to understand for future programmers. 
If you are working on a Macro that does not follow these conventions, only follow these conventions for your code. The codebase can then be updated at a later date.
# Module Layout
A module should be created for all functions and subprocesses that are similar. I.e. All pricing macros would be on one module etc. However, if the VBA code for a function is over thirty (30) lines. It can have its own module.
## Module Naming Convention
Module names need to be clear and easy to understand. They should use the snake case convention:
```vb
this_is_a_module
```
The name should match the text on the button on the Macro combo.
For modules that are no longer used, change the name to include OLD so we know it is no longer in use:
```vb
OLD_this_is_a_module
```
## General Layout
Limit line length to 72 characters unless splitting a line will make it more difficult to read.

Use 4-character tabs to indent, this is the standard in VBA and requires you to press the Tab key.
```vb
'Dont do this:
If thisIsAwesome Then
  MsgBox "Yes" '<- 2 character indents
Else
  MsgBox "No"
End if

'Good
If thisIsAwesome Then
    MsgBox "Yes" '<- 4 character indents'
```
Do not indent the contents of a function/subroutine unless it's part of a loop or control statement.
```vb
'Bad'
Sub anAwesomeRoutine()
    MsgBox "Nope"
End Sub

'Bad'
Function anotherFunction()
    If checkThis() Then
    Do This
    End If
End Function

'Good'
Sub anAwesomeRoutine()
MsgBox "Nope"
End Sub

'Good'
Function anotherFunction()
If checkThis() Then
    Do This
End If
End Function
```
Keep it as simple as possible. If you can split a calculation onto multiple lines, please do:
```vb
'Bad'
bigNumber = ((2 * 12)/5) * 100

'Good'
bigNumber = 2 * 12
bigNumber = bigNumber / 5
bigNumber = bigNumber * 100
```
# Variables and Naming
At the start of a module define **Option Explicit**, so that all variables need to be declared:
```vb
'Bad'
Sub anotherSub()
  'Do Things please'
End Sub

'Good'
Option Explicit
Sub anotherSub()
  'Do Things please'
End Sub
```
> It’s time to say “Goodbye,” to Hungarian Notation
>*Some programmer on the web.*

It has previously standard process to use Hungarian Notation, however even Microsoft have since moved on from this. Instead use camelcase and a name that explains the contents:
```vb
'Bad'
blnThisIsTrue =  true
intThisIsANumber = 22

'Good'
isEmpty = true
rowCounter = 22
```
# Function and Subroutines
Functions should only be used when returning a variable. Subroutines (sub) should be used in all other cases.
The naming convention for both functions and subroutines should be camel-case.
Functions should start with 'get', 'create', and 'calculate'.
Subroutines should start with 'do'.
```vb
'Bad'
Sub thisIsAwesome()
  'Do something'
End Sub

'Bad'
Function thisIsAwesome()
  'Do something'
End Sub

'Good'
Sub doChangeColour()
  'Do something'
End Sub

'Good'
Function getLastRow()
  'Do something'
End Sub

```
Arguments for subroutines and functions should be defined using the snakecase format, so that they vary from in-function (or sub) variables.
```vb
'Bad'
Sub doSomething(i, x)
  'Do Something'
End Sub

'Bad'
Function getSomething(number, isRight)
  'Do Something'
End Function

'Good'
Sub doSomething(countof_rows, count_of_columns)
  'Do Something'
End Sub

'Good'
Function getSomething(count_of_cells, count_of_rows)
  'Do Something'
End Function
  
```
