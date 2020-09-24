VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line counter"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   4125
      Top             =   855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Project"
      Filter          =   "Visual Basic Project (*.vbp)|*.vbp"
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4680
      TabIndex        =   40
      Top             =   3450
      Width           =   4680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   285
      Left            =   2590
      TabIndex        =   5
      Top             =   945
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   285
      Left            =   1090
      TabIndex        =   4
      Top             =   945
      Width           =   1000
   End
   Begin VB.CheckBox chkProperties 
      Caption         =   "Count lines written by VB (properties)"
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Value           =   1  'Checked
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   285
      Left            =   3765
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label grdPpags 
      Caption         =   "Property pages"
      Height          =   195
      Left            =   165
      TabIndex        =   45
      Top             =   2865
      Width           =   1200
   End
   Begin VB.Label grdPpagsC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   1470
      TabIndex        =   44
      Tag             =   "ok"
      Top             =   2865
      Width           =   705
   End
   Begin VB.Label grdPpagsD 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   2220
      TabIndex        =   43
      Tag             =   "ok"
      Top             =   2865
      Width           =   705
   End
   Begin VB.Label grdPpagsP 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   195
      Left            =   3000
      TabIndex        =   42
      Tag             =   "ok"
      Top             =   2865
      Width           =   405
   End
   Begin VB.Label grdPpagsLC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   3480
      TabIndex        =   41
      Tag             =   "ok"
      Top             =   2865
      Width           =   1050
   End
   Begin VB.Line Line5 
      X1              =   60
      X2              =   4600
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label grdTotalLC 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3480
      TabIndex        =   39
      Top             =   3135
      Width           =   1050
   End
   Begin VB.Label grdTotalP 
      Alignment       =   2  'Center
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2940
      TabIndex        =   38
      Top             =   3135
      Width           =   540
   End
   Begin VB.Label grdTotalD 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2220
      TabIndex        =   37
      Top             =   3135
      Width           =   705
   End
   Begin VB.Label grdTotalC 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1470
      TabIndex        =   36
      Top             =   3135
      Width           =   705
   End
   Begin VB.Label grdTotal 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   35
      Top             =   3135
      Width           =   1110
   End
   Begin VB.Line Line4 
      X1              =   60
      X2              =   4615
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Label grdDsrsLC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   3480
      TabIndex        =   34
      Tag             =   "ok"
      Top             =   2625
      Width           =   1050
   End
   Begin VB.Label grdDsrsP 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   195
      Left            =   3000
      TabIndex        =   33
      Tag             =   "ok"
      Top             =   2625
      Width           =   405
   End
   Begin VB.Label grdDsrsD 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   2220
      TabIndex        =   32
      Tag             =   "ok"
      Top             =   2625
      Width           =   705
   End
   Begin VB.Label grdDsrsC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   1470
      TabIndex        =   31
      Tag             =   "ok"
      Top             =   2625
      Width           =   705
   End
   Begin VB.Label grdCtlsLC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   3480
      TabIndex        =   30
      Tag             =   "ok"
      Top             =   2385
      Width           =   1050
   End
   Begin VB.Label grdCtlsP 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   195
      Left            =   3000
      TabIndex        =   29
      Tag             =   "ok"
      Top             =   2385
      Width           =   405
   End
   Begin VB.Label grdCtlsD 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   2220
      TabIndex        =   28
      Tag             =   "ok"
      Top             =   2385
      Width           =   705
   End
   Begin VB.Label grdCtlsC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   1470
      TabIndex        =   27
      Tag             =   "ok"
      Top             =   2385
      Width           =   705
   End
   Begin VB.Label grdClassesLC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   3480
      TabIndex        =   26
      Tag             =   "ok"
      Top             =   2130
      Width           =   1050
   End
   Begin VB.Label grdClassesP 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   195
      Left            =   3000
      TabIndex        =   25
      Tag             =   "ok"
      Top             =   2130
      Width           =   405
   End
   Begin VB.Label grdClassesD 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   2220
      TabIndex        =   24
      Tag             =   "ok"
      Top             =   2130
      Width           =   705
   End
   Begin VB.Label grdClassesC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   1470
      TabIndex        =   23
      Tag             =   "ok"
      Top             =   2130
      Width           =   705
   End
   Begin VB.Label grdModulesLC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   3480
      TabIndex        =   22
      Tag             =   "ok"
      Top             =   1905
      Width           =   1050
   End
   Begin VB.Label grdModulesP 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   195
      Left            =   3000
      TabIndex        =   21
      Tag             =   "ok"
      Top             =   1905
      Width           =   405
   End
   Begin VB.Label grdModulesD 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   2220
      TabIndex        =   20
      Tag             =   "ok"
      Top             =   1905
      Width           =   705
   End
   Begin VB.Label grdModulesC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   1470
      TabIndex        =   19
      Tag             =   "ok"
      Top             =   1905
      Width           =   705
   End
   Begin VB.Label grdFormsLC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   3480
      TabIndex        =   18
      Tag             =   "ok"
      Top             =   1665
      Width           =   1050
   End
   Begin VB.Label grdFormsP 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   195
      Left            =   3000
      TabIndex        =   17
      Tag             =   "ok"
      Top             =   1665
      Width           =   405
   End
   Begin VB.Label grdFormsD 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   2220
      TabIndex        =   16
      Tag             =   "ok"
      Top             =   1665
      Width           =   705
   End
   Begin VB.Label grdFormsC 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   195
      Left            =   1470
      TabIndex        =   15
      Tag             =   "ok"
      Top             =   1665
      Width           =   705
   End
   Begin VB.Label grdLinesCount 
      Alignment       =   2  'Center
      Caption         =   "Lines count"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   14
      Top             =   1395
      Width           =   1050
   End
   Begin VB.Label grdPercent 
      Alignment       =   2  'Center
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      Top             =   1395
      Width           =   405
   End
   Begin VB.Label grdDone 
      Alignment       =   2  'Center
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2220
      TabIndex        =   12
      Top             =   1395
      Width           =   705
   End
   Begin VB.Label hdrCount 
      Alignment       =   2  'Center
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1470
      TabIndex        =   11
      Top             =   1395
      Width           =   705
   End
   Begin VB.Label grdDsrs 
      Caption         =   "Designers"
      Height          =   195
      Left            =   165
      TabIndex        =   10
      Top             =   2625
      Width           =   1200
   End
   Begin VB.Label grdCtls 
      Caption         =   "User Controls"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   2385
      Width           =   1200
   End
   Begin VB.Label grdClasses 
      Caption         =   "Class Modules"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   2130
      Width           =   1200
   End
   Begin VB.Label grdModules 
      Caption         =   "Modules"
      Height          =   195
      Left            =   165
      TabIndex        =   7
      Top             =   1905
      Width           =   1200
   End
   Begin VB.Label grdForms 
      Caption         =   "Forms"
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   1665
      Width           =   1200
   End
   Begin VB.Line Line3 
      X1              =   4605
      X2              =   4605
      Y1              =   1395
      Y2              =   3380
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   60
      Y1              =   1395
      Y2              =   3380
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   4615
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Select project:"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   30
      Width           =   2475
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const EXAMINE_PROJECT = "Examining project..."
Private Const SCAN_FORM = "Scanning form ("
Private Const SCAN_MODULE = "Scanning module ("
Private Const SCAN_CTL = "Scanning user control ("
Private Const SCAN_CLASS = "Scanning class module ("
Private Const SCAN_DSR = "Scanning designer ("
Private Const SCAN_PROPERTYPAGE = "Scanning property page ("

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Private Function ThinBorder(ByVal bState As Boolean)
    Dim lS As Long

    For Each ctl In Me.Controls
        If TypeOf ctl Is PictureBox Or TypeOf ctl Is TextBox Then 'can be all controls which have hWnd property
            lS = GetWindowLong(ctl.hWnd, GWL_EXSTYLE)
            If Not (bState) Then
                lS = lS Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
            Else
                lS = lS Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
            End If
            SetWindowLong ctl.hWnd, GWL_EXSTYLE, lS
            SetWindowPos ctl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
        End If
    Next
End Function

Private Sub Command1_Click()
    With cDlg
        .Flags = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNHideReadOnly
        .ShowOpen
        Text1.Text = .FileName
    End With
End Sub

Private Sub Command2_Click()
    Dim F, G
    Dim I As Long
    Dim S As String
    Dim j As String
    Dim k As Variant
    Dim OnlyWritten As Boolean, OnlyWritten2 As Boolean
        
    If Len(Trim(Text1.Text)) = 0 Then
        MsgBox "Filename is not valid!", vbCritical, App.Title
        Exit Sub
    End If
    If Dir(Trim(Text1.Text)) = "" Then
        MsgBox "File does not exist!", vbCritical, App.Title
        Exit Sub
    End If
        
'    On Error GoTo Nnapaka
        
    DrawText picStatus.hDC, 3, 1, EXAMINE_PROJECT, Len(EXAMINE_PROJECT)
    
    ClearObjects
    ClearLabels
    
    S = Text1.Text
    For I = 1 To Len(S)
        S = Left(S, Len(S) - 1)
        If Right(S, 1) = "\" Then Exit For
    Next
        
    F = FreeFile
    Open Text1.Text For Input As F
        Do While Not EOF(F)
            Line Input #F, bstr
            If Left(bstr, 4) = "Form" Then
                If Mid(Right(bstr, Len(bstr) - 5), 2, 1) = ":" Then
                    vbObj.FormObj.Add Right(bstr, Len(bstr) - 5), Right(bstr, Len(bstr) - 5)
                Else
                    vbObj.FormObj.Add S & Right(bstr, Len(bstr) - 5), Right(bstr, Len(bstr) - 5)
                End If
            End If
            If Left(bstr, 11) = "UserControl" Then
                If Mid(Right(bstr, Len(bstr) - 12), 2, 1) = ":" Then
                    vbObj.UserCtlObj.Add Right(bstr, Len(bstr) - 12), Right(bstr, Len(bstr) - 12)
                Else
                    vbObj.UserCtlObj.Add S & Right(bstr, Len(bstr) - 12), Right(bstr, Len(bstr) - 12)
                End If
            End If
            If Left(bstr, 8) = "Designer" Then
                If Mid(Right(bstr, Len(bstr) - 9), 2, 1) = ":" Then
                    vbObj.DesignerObj.Add Right(bstr, Len(bstr) - 9), Right(bstr, Len(bstr) - 9)
                Else
                    vbObj.DesignerObj.Add S & Right(bstr, Len(bstr) - 9), Right(bstr, Len(bstr) - 9)
                End If
            End If
            If Left(bstr, 6) = "Module" Then
                k = Split(Right(bstr, Len(bstr) - 7), ";", , vbTextCompare)
                If Mid(Trim(k(1)), 2, 1) = ":" Then
                    vbObj.ModuleObj.Add Trim(k(1)), Trim(k(1))
                Else
                    vbObj.ModuleObj.Add S & Trim(k(1)), Trim(k(1))
                End If
            End If
            If Left(bstr, 5) = "Class" Then
                k = Split(Right(bstr, Len(bstr) - 6), ";", , vbTextCompare)
                If Mid(Trim(k(1)), 2, 1) = ":" Then
                    vbObj.ClsModuleObj.Add Trim(k(1)), Trim(k(1))
                Else
                    vbObj.ClsModuleObj.Add S & Trim(k(1)), Trim(k(1))
                End If
            End If
            If Left(bstr, 12) = "PropertyPage" Then
                If Mid(Right(bstr, Len(bstr) - 13), 2, 1) = ":" Then
                    vbObj.PropertyPageObj.Add Right(bstr, Len(bstr) - 13), Right(bstr, Len(bstr) - 13)
                Else
                    vbObj.PropertyPageObj.Add S & Right(bstr, Len(bstr) - 13), Right(bstr, Len(bstr) - 13)
                End If
            End If
            grdFormsC.Caption = vbObj.FormObj.Count
            grdModulesC.Caption = vbObj.ModuleObj.Count
            grdClassesC.Caption = vbObj.ClsModuleObj.Count
            grdCtlsC.Caption = vbObj.UserCtlObj.Count
            grdDsrsC.Caption = vbObj.DesignerObj.Count
            grdPpagsC.Caption = vbObj.PropertyPageObj.Count
            CalculateTotalCount
        Loop
    Close #F
    
    If vbObj.FormObj.Count > 0 Then
        For I = 1 To vbObj.FormObj.Count
            G = FreeFile
            cnt = 0
            OnlyWritten = False
            OnlyWritten2 = True
            picStatus.Cls
            DrawText picStatus.hDC, 3, 1, SCAN_FORM & vbObj.FormObj.Item(I) & ")...", Len(SCAN_FORM & vbObj.FormObj.Item(I) & ")...")
            Debug.Print vbObj.FormObj.Item(I)
            Open vbObj.FormObj.Item(I) For Input As #G
                Do While Not EOF(G)
                    Line Input #G, asdf
                    If chkProperties.Value = 0 Then
                        If Left(asdf, 10) = "Attribute " Then
                            OnlyWritten = True
                            OnlyWritten2 = True
                        Else
                            OnlyWritten2 = False
                        End If
                        If OnlyWritten = True And OnlyWritten2 = False Then
                            cnt = cnt + 1
                        End If
                    Else
                        cnt = cnt + 1
                    End If
                Loop
            Close #G
            grdFormsD.Caption = I
            grdFormsLC.Caption = grdFormsLC.Caption + cnt
            CalculateTotalDone
            CalculatePercents objForm
            CalculateLinesCount
        Next
    End If
    
    If vbObj.ModuleObj.Count > 0 Then
        For I = 1 To vbObj.ModuleObj.Count
            G = FreeFile
            cnt = 0
            picStatus.Cls
            DrawText picStatus.hDC, 3, 1, SCAN_MODULE & vbObj.ModuleObj.Item(I) & ")...", Len(SCAN_MODULE & vbObj.ModuleObj.Item(I) & ")...")
            Open vbObj.ModuleObj.Item(I) For Input As #G
                Do While Not EOF(G)
                    Line Input #G, asdf
                    cnt = cnt + 1
                Loop
            Close #G
            grdModulesD.Caption = I
            grdModulesLC.Caption = grdModulesLC.Caption + cnt
            CalculateTotalDone
            CalculatePercents objModule
            CalculateLinesCount
        Next
    End If
    
    If vbObj.ClsModuleObj.Count > 0 Then
        For I = 1 To vbObj.ClsModuleObj.Count
            G = FreeFile
            cnt = 0
            picStatus.Cls
            DrawText picStatus.hDC, 3, 1, SCAN_CLASS & vbObj.ClsModuleObj.Item(I) & ")...", Len(SCAN_CLASS & vbObj.ClsModuleObj.Item(I) & ")...")
            Open vbObj.ClsModuleObj.Item(I) For Input As #G
                Do While Not EOF(G)
                    Line Input #G, asdf
                    cnt = cnt + 1
                Loop
            Close #G
            grdClassesD.Caption = I
            grdClassesLC.Caption = grdClassesLC.Caption + cnt
            CalculateTotalDone
            CalculatePercents objClassModule
            CalculateLinesCount
        Next
    End If
    
    If vbObj.UserCtlObj.Count > 0 Then
        For I = 1 To vbObj.UserCtlObj.Count
            G = FreeFile
            cnt = 0
            OnlyWritten = False
            OnlyWritten2 = True
            picStatus.Cls
            DrawText picStatus.hDC, 3, 1, SCAN_CTL & vbObj.UserCtlObj.Item(I) & ")...", Len(SCAN_CTL & vbObj.UserCtlObj.Item(I) & ")...")
            Open vbObj.UserCtlObj.Item(I) For Input As #G
                Do While Not EOF(G)
                    Line Input #G, asdf
                    If chkProperties.Value = 0 Then
                        If Left(asdf, 10) = "Attribute " Then
                            OnlyWritten = True
                            OnlyWritten2 = True
                        Else
                            OnlyWritten2 = False
                        End If
                        If OnlyWritten = True And OnlyWritten2 = False Then
                            cnt = cnt + 1
                        End If
                    Else
                        cnt = cnt + 1
                    End If
                Loop
            Close #G
            grdCtlsD.Caption = I
            grdCtlsLC.Caption = grdCtlsLC.Caption + cnt
            CalculateTotalDone
            CalculatePercents objUserControl
            CalculateLinesCount
        Next
    End If
    
    If vbObj.DesignerObj.Count > 0 Then
        For I = 1 To vbObj.DesignerObj.Count
            G = FreeFile
            cnt = 0
            OnlyWritten = False
            OnlyWritten2 = True
            picStatus.Cls
            DrawText picStatus.hDC, 3, 1, SCAN_DSR & vbObj.DesignerObj.Item(I) & ")...", Len(SCAN_DSR & vbObj.DesignerObj.Item(I) & ")...")
            Open vbObj.DesignerObj.Item(I) For Input As #G
                Do While Not EOF(G)
                    Line Input #G, asdf
                    If chkProperties.Value = 0 Then
                        If Left(asdf, 10) = "Attribute " Then
                            OnlyWritten = True
                            OnlyWritten2 = True
                        Else
                            OnlyWritten2 = False
                        End If
                        If OnlyWritten = True And OnlyWritten2 = False Then
                            cnt = cnt + 1
                        End If
                    Else
                        cnt = cnt + 1
                    End If
                Loop
            Close #G
            grdDsrsD.Caption = I
            grdDsrsLC.Caption = grdDsrsLC.Caption + cnt
            CalculateTotalDone
            CalculatePercents objDesigner
            CalculateLinesCount
        Next
    End If
    
    If vbObj.PropertyPageObj.Count > 0 Then
        For I = 1 To vbObj.PropertyPageObj.Count
            G = FreeFile
            cnt = 0
            OnlyWritten = False
            OnlyWritten2 = True
            picStatus.Cls
            DrawText picStatus.hDC, 3, 1, SCAN_PROPERTYPAGE & vbObj.PropertyPageObj.Item(I) & ")...", Len(SCAN_PROPERTYPAGE & vbObj.PropertyPageObj.Item(I) & ")...")
            Debug.Print vbObj.PropertyPageObj.Item(I)
            Open vbObj.PropertyPageObj.Item(I) For Input As #G
                Do While Not EOF(G)
                    Line Input #G, asdf
                    If chkProperties.Value = 0 Then
                        If Left(asdf, 10) = "Attribute " Then
                            OnlyWritten = True
                            OnlyWritten2 = True
                        Else
                            OnlyWritten2 = False
                        End If
                        If OnlyWritten = True And OnlyWritten2 = False Then
                            cnt = cnt + 1
                        End If
                    Else
                        cnt = cnt + 1
                    End If
                Loop
            Close #G
            grdPpagsD.Caption = I
            grdPpagsLC.Caption = grdPpagsLC.Caption + cnt
            CalculateTotalDone
            CalculatePercents objPropertyPage
            CalculateLinesCount
        Next
    End If
    picStatus.Cls
    Exit Sub


Nnapaka:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    End If
    ClearLabels
    picStatus.Cls
    Exit Sub
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ClearObjects
    ThinBorder True
End Sub

Function GetFileName(txt As String) As String
    Dim S As String
    Dim I As Long
    
    S = txt
    
    For I = 1 To Len(txt)
        S = Left(S, Len(S) - 1)
        If Right(S, 1) = "\" Then
            GetFileName = Right(txt, Len(txt) - (Len(S) + 1))
            Exit For
        End If
    Next
End Function

Sub ClearObjects()
    Set vbObj.ClsModuleObj = Nothing
    Set vbObj.DesignerObj = Nothing
    Set vbObj.FormObj = Nothing
    Set vbObj.ModuleObj = Nothing
    Set vbObj.UserCtlObj = Nothing
    Set vbObj.PropertyPageObj = Nothing
    
    Set vbObj.ClsModuleObj = New Collection
    Set vbObj.DesignerObj = New Collection
    Set vbObj.FormObj = New Collection
    Set vbObj.ModuleObj = New Collection
    Set vbObj.UserCtlObj = New Collection
    Set vbObj.PropertyPageObj = New Collection
End Sub

Sub CalculateTotalCount()
    grdTotalC.Caption = CLng(grdFormsC.Caption) + CLng(grdModulesC.Caption) + CLng(grdClassesC.Caption) + CLng(grdCtlsC.Caption) + CLng(grdDsrsC.Caption) + CLng(grdPpagsC.Caption)
    RefreshLabels
End Sub

Sub CalculateTotalDone()
    grdTotalD.Caption = CLng(grdFormsD.Caption) + CLng(grdModulesD.Caption) + CLng(grdClassesD.Caption) + CLng(grdCtlsD.Caption) + CLng(grdDsrsD.Caption) + CLng(grdPpagsD.Caption)
    RefreshLabels
End Sub

Sub CalculatePercents(what As TypeOfObject)
    Select Case what
        Case objForm
            grdFormsP.Caption = Round((grdFormsD.Caption / grdFormsC.Caption) * 100, 0) & "%"
        Case objModule
            grdModulesP.Caption = Round((grdModulesD.Caption / grdModulesC.Caption) * 100, 0) & "%"
        Case objClassModule
            grdClassesP.Caption = Round((grdClassesD.Caption / grdClassesC.Caption) * 100, 0) & "%"
        Case objUserControl
            grdCtlsP.Caption = Round((grdCtlsD.Caption / grdCtlsC.Caption) * 100, 0) & "%"
        Case objDesigner
            grdDsrsP.Caption = Round((grdDsrsD.Caption / grdDsrsC.Caption) * 100, 0) & "%"
        Case objPropertyPage
            grdPpagsP.Caption = Round((grdPpagsD.Caption / grdPpagsC.Caption) * 100, 0) & "%"
    End Select
    grdTotalP.Caption = Round((grdTotalD.Caption / grdTotalC.Caption) * 100, 0) & "%"
    RefreshLabels
End Sub

Sub CalculateLinesCount()
    grdTotalLC.Caption = CLng(grdFormsLC.Caption) + CLng(grdModulesLC.Caption) + CLng(grdClassesLC.Caption) + CLng(grdCtlsLC.Caption) + CLng(grdDsrsLC.Caption) + CLng(grdPpagsLC.Caption)
    RefreshLabels
End Sub

Sub RefreshLabels()
    grdFormsC.Refresh
    grdModulesC.Refresh
    grdClassesC.Refresh
    grdCtlsC.Refresh
    grdDsrsC.Refresh
    grdPpagsC.Refresh
    grdFormsD.Refresh
    grdModulesD.Refresh
    grdClassesD.Refresh
    grdCtlsD.Refresh
    grdDsrsD.Refresh
    grdPpagsD.Refresh
    grdFormsP.Refresh
    grdModulesP.Refresh
    grdClassesP.Refresh
    grdCtlsP.Refresh
    grdDsrsP.Refresh
    grdPpagsP.Refresh
    grdFormsLC.Refresh
    grdModulesLC.Refresh
    grdClassesLC.Refresh
    grdCtlsLC.Refresh
    grdDsrsLC.Refresh
    grdDsrsLC.Refresh
End Sub

Sub ClearLabels()
    grdFormsC.Caption = "0"
    grdModulesC.Caption = "0"
    grdClassesC.Caption = "0"
    grdCtlsC.Caption = "0"
    grdDsrsC.Caption = "0"
    grdPpagsC.Caption = "0"
    grdTotalC.Caption = "0"
    grdFormsD.Caption = "0"
    grdModulesD.Caption = "0"
    grdClassesD.Caption = "0"
    grdCtlsD.Caption = "0"
    grdDsrsD.Caption = "0"
    grdPpagsD.Caption = "0"
    grdTotalD.Caption = "0"
    grdFormsP.Caption = "0%"
    grdModulesP.Caption = "0%"
    grdClassesP.Caption = "0%"
    grdCtlsP.Caption = "0%"
    grdDsrsP.Caption = "0%"
    grdPpagsP.Caption = "0%"
    grdTotalP.Caption = "0%"
    grdFormsLC.Caption = "0"
    grdModulesLC.Caption = "0"
    grdClassesLC.Caption = "0"
    grdCtlsLC.Caption = "0"
    grdDsrsLC.Caption = "0"
    grdPpagsLC.Caption = "0"
    grdTotalLC.Caption = "0"
End Sub
