VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmployeeForm 
   Caption         =   "Employee User Form"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "EmployeeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmployeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@PredeclaredId
Option Explicit
Private Type TEmployeeForm
    CompanyName As String
End Type

Private This As TEmployeeForm

Public Property Get CompanyName() As String
    CompanyName = This.CompanyName
End Property

Public Property Let CompanyName(ByVal RHS As String)
    This.CompanyName = RHS
End Property

Public Function Create(GivenCompanyName As String) As EmployeeForm
    
    Dim CurrentEmployeeForm As EmployeeForm
    Set CurrentEmployeeForm = New EmployeeForm
    
    With CurrentEmployeeForm
        .CompanyName = GivenCompanyName
        'Using Ref
        .CompanyNameLabel.Caption = GivenCompanyName
        'Using Ref >> Only Change in this line
        SetEmployeeNameUsingRef CurrentEmployeeForm
    End With
    Set Create = CurrentEmployeeForm
    
End Function


Private Sub SetEmployeeNameUsingMe()
    Me.EmployeeComboBox.List = Array("Ismail", "Kamal", "Petr")
End Sub

Private Sub SetEmployeeNameUsingRef(SetToUF As EmployeeForm)
    SetToUF.EmployeeComboBox.List = Array("Ismail", "Kamal", "Petr")
End Sub

