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
Global LevelToLoad As Integer
Global CancelLoad As Boolean
Global TotalLevels As Integer
Global LevelPack As String

Global TargetMoves As Long
Global TargetTime As Long

Global X As Long
Global F As String
Global Score As Long

Global YourMoves As Long
Global YourTime As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Enum T_WindowStyle
    Maximized = 3
    Normal = 1
    ShowOnly = 5
End Enum


Public Sub OpenInternet(Parent As Form, URL As String, WindowStyle As T_WindowStyle)
    ShellExecute Parent.hwnd, "Open", URL, "", "", WindowStyle
End Sub

