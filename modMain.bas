Attribute VB_Name = "modMain"
Public Declare Function DrawText Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Type VB_Objects
    FormObj As Collection
    ModuleObj As Collection
    ClsModuleObj As Collection
    UserCtlObj As Collection
    DesignerObj As Collection
    PropertyPageObj As Collection
End Type

Public Enum TypeOfObject
    objForm
    objModule
    objClassModule
    objUserControl
    objDesigner
    objPropertyPage
    objTotal
End Enum

Public ProjectPath As String
Public vbObj As VB_Objects

