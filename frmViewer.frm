VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViewer 
   Caption         =   "Form Viewer"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstControls 
      Height          =   4545
      ItemData        =   "frmViewer.frx":0000
      Left            =   4800
      List            =   "frmViewer.frx":0002
      TabIndex        =   4
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox lstForms 
      Height          =   3180
      ItemData        =   "frmViewer.frx":0004
      Left            =   120
      List            =   "frmViewer.frx":0006
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   1560
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDefault 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuForms 
      Caption         =   "Forms"
      Visible         =   0   'False
      Begin VB.Menu mnuPrev 
         Caption         =   "Preview"
      End
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type cmdLine
    LValue As String
    RValue As String
End Type
Private Type VBForm
    FName As String
    Name As String
    FHeight As Long
    FWidth As Long
End Type
Private Type VBControl
    CType As String
    CHeight As Integer
    CWidth As Integer
    CLeft As Integer
    CTop As Integer
    CName As String
    CCaption As String
    CTabIndex As Integer
    CIndex As Integer
End Type
Private Forms(20) As VBForm
Private MyControl(100) As VBControl
Private frmCount, ctrlCount As Integer
Private Function Split(Strn As String) As cmdLine
    Split.LValue = Mid(Strn, 1, InStr(Strn, " ") - 1)
    Split.RValue = Mid(Strn, InStr(Strn, " ") + 1, Len(Strn) - InStr(Strn, " "))
End Function
Private Function Slim(Strn As String) As String
    Dim i As Integer
    Dim tmp As String
    For i = 1 To Len(Strn)
        If Mid(Strn, i, 1) <> " " Then tmp = tmp & Mid(Strn, i, 1)
    Next i
    Slim = tmp
End Function
Private Sub cmdOpen_Click()
    On Error GoTo canceled
    Dim ln As String
    Dim ln2 As String
    Dim go As Integer
    comDlg.CancelError = True
    comDlg.DefaultExt = ".vbp"
    comDlg.Filter = "Visual Basic 6 Project (*.vbp)|*.vbp"
    comDlg.ShowOpen
    lstForms.Clear
    Open comDlg.FileName For Input As #1
    While Not EOF(1)
        Line Input #1, ln
        ln = Slim(ln)
        If Mid(ln, 1, 8) = "Startup=" Then lblDefault.Caption = "StartUp Form = " & Mid(ln, 9, Len(ln))
        If Mid(ln, 1, 5) = "Name=" Then lblTitle.Caption = "Project Name = " & Mid(ln, 6, Len(ln))
        If Mid(ln, 1, 5) = "Form=" Then
            Open Mid(ln, 6, Len(ln)) For Input As #2
            Forms(frmCount).FName = Mid(ln, 6, Len(ln))
            go = 1
            While Not EOF(2) And go < 4
                Line Input #2, ln2
                ln2 = Slim(ln2)
                If Mid(ln2, 1, 8) = "Caption=" Then
                    Forms(frmCount).Name = Mid(ln2, 9, Len(ln2) - 8)
                    lstForms.AddItem Forms(frmCount).FName & " (" & Forms(frmCount).Name & ")"
                    go = go + 1
                End If
                If Mid(ln2, 1, 13) = "ClientHeight=" Then
                    Forms(frmCount).FHeight = Val(Mid(ln2, 14, Len(ln2) - 13))
                    go = go + 1
                End If
                If Mid(ln2, 1, 12) = "ClientWidth=" Then
                    Forms(frmCount).FWidth = Val(Mid(ln2, 13, Len(ln2) - 12))
                    go = go + 1
                End If
            Wend
            Close #2
            frmCount = frmCount + 1
        End If
    Wend
    Close #1
    Exit Sub
canceled:
End Sub

Private Sub lstForms_Click()
    Dim FileName As String
    Dim ln As String
    Dim Controlln As cmdLine
    lstControls.Clear
    ctrlCount = 0
    FileName = Mid(lstForms.List(lstForms.ListIndex), 1, InStr(lstForms.List(lstForms.ListIndex), ".frm") + 3)
    Open FileName For Input As #1
    While Not EOF(1)
        Line Input #1, ln
        ln = Trim(ln)
        If InStr(ln, "Begin VB.") Then
            MyControl(ctrlCount).CCaption = ""
            MyControl(ctrlCount).CIndex = 0
            ln = Mid(ln, 10, Len(ln) - 9)
            Controlln = Split(ln)
            lstControls.AddItem Controlln.LValue & "->" & Controlln.RValue
            MyControl(ctrlCount).CName = Controlln.RValue
            MyControl(ctrlCount).CType = "VB." & Controlln.LValue
            ctrlCount = ctrlCount + 1
        End If
        ln = Slim(ln)
        If Mid(ln, 1, 8) = "Caption=" Then MyControl(ctrlCount - 1).CCaption = Mid(ln, 10, Len(ln) - 10)
        If Mid(ln, 1, 4) = "Top=" Then MyControl(ctrlCount - 1).CTop = Val(Mid(ln, 5, Len(ln) - 4))
        If Mid(ln, 1, 5) = "Left=" Then MyControl(ctrlCount - 1).CLeft = Val(Mid(ln, 6, Len(ln) - 5))
        If Mid(ln, 1, 7) = "Height=" Then MyControl(ctrlCount - 1).CHeight = Val(Mid(ln, 8, Len(ln) - 7))
        If Mid(ln, 1, 6) = "Width=" Then MyControl(ctrlCount - 1).CWidth = Val(Mid(ln, 7, Len(ln) - 6))
        If Mid(ln, 1, 9) = "TabIndex=" Then MyControl(ctrlCount - 1).CTabIndex = Val(Mid(ln, 10, Len(ln) - 9))
        If Mid(ln, 1, 6) = "Index=" Then MyControl(ctrlCount - 1).CIndex = Val(Mid(ln, 7, Len(ln) - 6))
    Wend
    Close #1
End Sub

Private Sub lstForms_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuForms
End Sub

Private Sub mnuExit_Click()
    End
End Sub
Private Sub mnuPrev_Click()
    On Error Resume Next
    Dim i As Integer
    Dim tmpCtrl As String
    frmPreview.Show
    frmPreview.Width = Forms(lstForms.ListIndex).FWidth
    frmPreview.Height = Forms(lstForms.ListIndex).FHeight
    For i = 0 To ctrlCount - 1
'        If MyControl(i).CType <> "VB.Form" And MyControl(i).CType <> "VB.Line" Then
            tmpCtrl = MyControl(i).CName & MyControl(i).CIndex
            frmPreview.Controls.Add MyControl(i).CType, tmpCtrl
            frmPreview.Controls(tmpCtrl).Left = MyControl(i).CLeft
            frmPreview.Controls(tmpCtrl).Top = MyControl(i).CTop
            frmPreview.Controls(tmpCtrl).Height = MyControl(i).CHeight
            If MyControl(i).CCaption <> "" Then frmPreview.Controls(tmpCtrl).Caption = MyControl(i).CCaption
            frmPreview.Controls(tmpCtrl).Visible = True
'        End If
    Next i
Fin:
End Sub

