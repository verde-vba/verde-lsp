Attribute VB_Name = "AccessModifiers"

' Private declarations
Private Number As Integer
Private NameArray(1 To 5) As String
Private MyVar, YourVar, ThisVar As Integer

' Public declarations
Public PubNumber As Integer
Public PubNameArray(1 To 5) As String
Public PubMyVar, PubYourVar, PubThisVar As Integer

' Static variables
Sub UseStaticVar()
    Static counter As Integer
    counter = counter + 1
    Debug.Print counter
End Sub

' Static in module-level context
Static loanAmt
Static loanInt
Static loanTerm
