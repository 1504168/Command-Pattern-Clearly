Attribute VB_Name = "Tester"
Option Explicit

Sub Test()
    
    Dim FormUsingFactory As EmployeeForm
    Set FormUsingFactory = EmployeeForm.Create("OOP In VBA FB Group")
    FormUsingFactory.Show vbModeless
    FormUsingFactory.Caption = "This is from Create function"
    FormUsingFactory.Left = 200
    
    EmployeeForm.Show vbModeless
    EmployeeForm.Caption = "This is Global one"
    EmployeeForm.Left = 600
    
End Sub


Sub AnotherTest()
    Dim FirstPerson As Person
    Set FirstPerson = Person.Create("Md Ismail Hosen", 26)
    Dim SecondPerson As Person
    Set SecondPerson = Person.Create("Khadiza", 1)
    Debug.Print "First Person Name : " & FirstPerson.Name
    Debug.Print "Second Person Name : " & SecondPerson.Name
End Sub
