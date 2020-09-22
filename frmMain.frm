VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Finds all File Extensions in a Directory"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "frmMain.frx":0000
      Left            =   5400
      List            =   "frmMain.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.FileListBox File 
      Height          =   3015
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.DirListBox Dir 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Filter By:"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Dave Lein (davelein@hotmail.com)
'This is a very simple program to understand that will perform a filter
'I wrote this for someone who emailed me needing this exact process
'The SortList Function is really crap because its a bubble sort
'It was the first sort I thought of and one of the easiest to do

Private Sub Dir_Change()
    File.Path = Dir.Path 'Update Path
    List1.Clear          'Clear ListBox
    
    UpdateList           'Call Function
    DoSort               'Call Function
End Sub

Private Sub Drive_Change()
On Error Resume Next 'This is here because if no disk in A: it will fail
    Dir.Path = Drive.Drive
End Sub

Private Sub Form_Load()
    UpdateList 'Call Function
    DoSort     'Call Function
End Sub

Private Sub List1_Click()
    File.Pattern = "*" & List1.List(List1.ListIndex) 'Just change the Pattern
End Sub

Function DoSort()
'This has to be run 1 less than the total list of items
'because calling SortList will only place 1 item in the right place for sure
'that being the last item
    For x = 0 To List1.ListCount - 1
        SortList
    Next x
End Function

Function SortList()
'This is a simple bubble sort
'This is the least efficient sort of all sorts
'It has to Run the sort 1 minus the total List of items
    Dim Temp As String
    
    For x = 0 To List1.ListCount - 1
        If List1.List(x) < List1.List(x + 1) Then
            Temp = List1.List(x) 'Have to store otherwise will be overwritten
            List1.List(x) = List1.List(x + 1) 'Overwrites x with value of next
            List1.List(x + 1) = Temp          'Overwrites x+1 with the Stored Temp
        End If
    Next x
End Function

Function UpdateList()

    Dim GetExt As String
    Dim cnt As Integer, loc As Integer
    
    cnt = 0                                 'Start Counter at 0
    
    For x = 0 To File.ListCount - 1
        loc = InStr(1, File.List(cnt), ".") 'Find location of . in the File Name
        If loc = 0 Then GoTo No             'No . in File Name
            
        GetExt = Mid(File.List(cnt), loc, Len(File.List(cnt)) - Len(loc)) 'Start at loc and take out whatevers after it
                                            'Reads entire contents after . not just 3 or 2 characters
        GoTo Yes  'Sloppy I know
No:
        GetExt = "" 'Set the Extension to Nothing because theres no .
Yes:
        For y = 0 To List1.ListCount - 1
            If LCase(GetExt) = LCase(List1.List(y)) Then 'Check for redundancy
                GetExt = "" 'Set Extension to Nothing
            End If
        Next y
        
        If GetExt <> vbNullString Then ' (VbNullString = "")
            List1.AddItem GetExt, 0 'Add only if its a valid extension
        End If
        
        cnt = cnt + 1 'Remember to add 1 to count
    Next x
   
End Function
