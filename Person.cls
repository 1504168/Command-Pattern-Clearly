VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit


Private Type TPerson
    Name As String
    Age As Integer
End Type

Private This As TPerson

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal RHS As String)
    This.Name = RHS
End Property

Public Property Get Age() As Integer
    Age = This.Age
End Property

Public Property Let Age(ByVal RHS As Integer)
    This.Age = RHS
End Property

Public Property Get GetMe() As Person
    Set GetMe = Me
    Me.Name = "Test Name"
End Property

Public Function Create(GivenName As String, GivenAge As Integer) As Person
    With Me
        .Name = GivenName
        .Age = GivenAge
        Set Create = Me
    End With
End Function

