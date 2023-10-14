﻿// See https://aka.ms/new-console-template for more information
using VBA2CS;

Console.WriteLine("Hello, World!");

string code = @"
Option Explicit

' Declare a public variable
Public myGlobalVar As Integer

' Declare a constant
Public Const MY_CONST As Integer = 10

' Declare an enumeration
Enum Days
    Sunday = 1
    Monday = 2
    Tuesday = 3
End Enum

' Declare a user-defined type
Type Person
    Name As String
    Age As Integer
End Type

' A simple function to add two numbers
Function AddNumbers(a As Integer, b As Integer) As Integer
    Dim result As Integer
    result = a + b
    AddNumbers = result
End Function

' A subroutine to set a global variable
Sub SetGlobalVar()
    myGlobalVar = 5
End Sub

' Using If-Then-Else and Select Case statements
Sub ControlStatementsDemo(x As Integer)
    Dim day As Days
    day = x
    
    If x > 5 Then
        MsgBox ""x is greater than 5""
    ElseIf x < 5 Then
        MsgBox ""x is less than 5""
    Else
        MsgBox ""x is equal to 5""
    End If
    
    Select Case day
        Case Days.Sunday
            MsgBox ""It's a relaxing day!""
        Case Days.Monday
            MsgBox ""The start of the workweek.""
        Case Else
            MsgBox ""It's a regular day.""
    End Select
End Sub

' Using loops
Sub LoopDemo()
    Dim i As Integer
    Dim person As Person
    person.Name = ""John""
    person.Age = 30
    
    For i = 1 To 5
        Debug.Print i
    Next i
    
    For Each element In Array(1, 2, 3)
        Debug.Print element
    Next element
    
    Do While person.Age < 40
        person.Age = person.Age + 1
    Loop
End Sub

' Entry point
Sub Main()
    Dim result As Integer
    SetGlobalVar
    result = AddNumbers(myGlobalVar, MY_CONST)
    ControlStatementsDemo result
    LoopDemo
End Sub
";

string code2 = @"
If a = ""Hello"" And b = 123 Then
    ' This is a comment
    Dim c As String
    c = a & "" World""
    End If
";

var tokens = Tokenizer.Tokenize(code);

foreach (var token in tokens)
{
    Console.WriteLine(token.ToString());
}

/*
Parser parser = new Parser();
Node root = parser.Parse(tokens);

string csharpCode = parser.ConvertToCSharp(root);
*/
//Console.WriteLine($"CSCode = \n{csharpCode}");