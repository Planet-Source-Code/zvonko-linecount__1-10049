Attribute VB_Name = "modBrowseFolder"
Option Explicit

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Const WM_USER = &H400
Public Const LPTR = (&H0 Or &H40)
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Public Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If uMsg = 1 Then
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If
End Function

Public Function FunctionPointer(FunctionAddress As Long) As Long
    FunctionPointer = FunctionAddress
End Function

Public Function BrowseForFolder(hWnd As Long, Optional Title As String = "Mape", Optional Path As String = "C:\") As String
    Dim Bff As BROWSEINFOTYPE
    Dim itemID As Long
    Dim PathPointer As Long
    Dim tmpPath As String * 256
    With Bff
        .hOwner = hWnd
        .lpszTitle = Title
        .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr)
        PathPointer = LocalAlloc(LPTR, Len(Path) + 1)
        CopyMemory ByVal PathPointer, ByVal Path, Len(Path) + 1
        .lParam = PathPointer
    End With
    itemID = SHBrowseForFolder(Bff)
    If itemID Then
        If SHGetPathFromIDList(itemID, tmpPath) Then
            BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(itemID)
    End If
    Call LocalFree(PathPointer)
End Function

