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



Sub AnotherTest2()
    
'    Person.Name = "Test Name"
    Debug.Print "Global one name : " & Person.Name
    
    Dim FirstPerson As Person
    Set FirstPerson = Person.Create("Md Ismail Hosen", 26)
    
    Debug.Print "Global one name : " & Person.Name
    
    Dim SecondPerson As Person
    Set SecondPerson = Person.Create("Khadiza", 1)
    
    Debug.Print "First Person Name : " & FirstPerson.Name
    Debug.Print "Second Person Name : " & SecondPerson.Name
    
    ''Let's be little crazy and then get the global one
    Dim ThirdPerson As Person
    Set ThirdPerson = Person.GetMe
    
    Debug.Print "Third Person Name : " & ThirdPerson.Name
    Set ThirdPerson = Nothing
    
End Sub

