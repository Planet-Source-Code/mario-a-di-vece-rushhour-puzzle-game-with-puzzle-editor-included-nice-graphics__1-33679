Attribute VB_Name = "Module1"
Type Car
    CarID As Integer
    Movement As String
    NextRightBlock As Integer
    NextLeftBlock As Integer
    NextTopBlock As Integer
    NextDownBlock As Integer
    Position As String
End Type

Global CarSet(29) As Car
