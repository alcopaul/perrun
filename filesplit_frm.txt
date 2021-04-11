VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "File Splitter And Merger by alcopaul"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   5280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   4800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   14
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Merge"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   9615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   9
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Split"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   9735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "merge stats :"
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   17
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "split stats :"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Output file (with path) :"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   13
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "File to merge :"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Name (no path and extension, pls) :"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "byte files"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Segment size"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File to split :"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================================================
'splitter and merger by alcopaul
'email:(alcopaulvx@yahoo.com)
'homepage:(do a yahoo search)
'
' "programming is like an art.. whatever methods used, whether they're sloppy or not, the
' important thingie is you get what you want..." - anonymous
'
' this file can split files into specified chunks and can bring them back to a single
' file... enjoy.. don't forget to vote..
'
' complete application is short.. no external apis used....
'
'=====================================================================================
Option Explicit
Private Sub Command1_Click()
On Error GoTo handler
Dim sFile As String
ProgressBar1.Value = 0
With CommonDialog1
.DialogTitle = "Open"
.CancelError = False
.Filter = "All Files (*.*)|*.*"
.ShowOpen
If Len(.FileName) = 0 Then
Exit Sub
End If
sFile = .FileName
End With
Text1.Text = sFile
GoTo kko
handler:
MsgBox "invalid input", , "<Error>"
kko:
End Sub

Private Sub Command2_Click()
'split file
If Text3.Text = "" Then GoTo ggg
split Text1.Text, Text2.Text
MsgBox "File Splitted"
GoTo hhh
ggg:
MsgBox "you forgot to put the name of the segments", , "oi!"
hhh:
End Sub

Function split(file As String, filesize As String)
Dim filesize1 As Long, n As Long, retrieve As String, lastchunk As String, i As Long, num As Long, fileextension As String, lastnum As Long
'open file and get chunk size
filesize1 = CLng(filesize)
Open file For Binary Access Read As #11
ProgressBar1.Max = Int(LOF(11) / filesize1)
For i = 1 To Int(LOF(11) / filesize1)
ProgressBar1.Value = i
DoEvents
retrieve = Space(filesize1)
'get chunks
Get #11, , retrieve

'===============================================================
'maintain a three character extension name
'===============================================================
num = num + 1
fileextension = CStr(num)
If Len(fileextension) = 1 Then
fileextension = "00" & fileextension
ElseIf Len(fileextension) = 2 Then
fileextension = "0" & fileextension
End If
'===============================================================

'save chunks to phile
writetofile Text3.Text, fileextension, retrieve
'next chunk
Next i

'if remainder = 0 then close file
If LOF(11) - ((Int((LOF(11) / filesize1))) * (filesize1)) = 0 Then GoTo clse
'else
ProgressBar1.Max = ProgressBar1.Max + 1
'get remaining chunks
lastchunk = Space(LOF(11) - ((Int(LOF(11) / filesize1)) * (filesize1)))
ProgressBar1.Value = ProgressBar1.Value + 1
Get #11, , lastchunk
'===============================================================
'maintain a three character extension name
'===============================================================
lastnum = num + 1
fileextension = CStr(lastnum)
If Len(fileextension) = 1 Then
fileextension = "00" & fileextension
ElseIf Len(fileextension) = 2 Then
fileextension = "0" & fileextension
End If

'save remaining chunk to file
writetofile Text3.Text, fileextension, lastchunk
clse:

'close file
Close #11
End Function

'save chunks to file
Function writetofile(file As String, extension As String, data As String)
Dim datatofile As String
datatofile = data
Open resolvepath & file & "." & extension For Binary Access Write As #33
Put #33, , datatofile
Close #33
End Function

'save files will be in the directory where this proggie is located
Function resolvepath() As String
Dim currentpath As String
currentpath = App.Path
If Right(currentpath, 1) <> "\" Then currentpath = currentpath & "\"
resolvepath = currentpath
End Function

'open splitted file to merge
Private Sub Command3_Click()
On Error GoTo handler
Dim sFile As String
ProgressBar2.Value = 0
With CommonDialog2
.DialogTitle = "Open"
.CancelError = False
.Filter = "001 Files (*.001)|*.001"
.ShowOpen
If Len(.FileName) = 0 Then
Exit Sub
End If
sFile = .FileName
End With
Text4.Text = sFile
GoTo kko
handler:
MsgBox "invalid input", , "<Error>"
kko:
End Sub
Private Function namefile(strPath As String) As String
  namefile = Mid(strPath, InStrRev(strPath, "\") + 1)
End Function

'merge routine
Function merge(file As String, file1 As String)
Dim ext As Long, data As String, segfile As String, i As Long, num1 As Long, num As Long, fileextension As String

'open the output file
Open file1 For Binary Access Write As #22
' determine the number of segments
Do
ext = ext + 1
num1 = num1 + 1
Select Case num1
Case Is < 10
segfile = remext(file) & ".00" & CStr(ext)
Case 10 To 99
segfile = remext(file) & ".0" & CStr(ext)
Case 100 To 999
segfile = remext(file) & "." & CStr(ext)
End Select
If Dir(segfile) = namefile(segfile) Then
Else
Exit Do
End If
Loop
'
ProgressBar2.Max = ext - 1
'variable ext contains the number of segments
For i = 1 To (ext - 1)
DoEvents
ProgressBar2.Value = i
'===============================================================
'maintain a three character extension name
'===============================================================
num = num + 1
fileextension = CStr(num)
If Len(fileextension) = 1 Then
fileextension = "00" & fileextension
ElseIf Len(fileextension) = 2 Then
fileextension = "0" & fileextension
End If
'================================================================

'open segment for reading
Open remext(file) & "." & fileextension For Binary Access Read As #1
data = Space(LOF(1))
'get segment bytes
Get #1, , data
'record segment bytes to output file
Put #22, , data
Close #1
Next i ' until all files are read

'all files are read, close output file
Close #22
End Function
'remove extension
Private Function remext(strPath As String) As String
  remext = Mid(strPath, 1, Len(strPath) - 4)
End Function

Private Sub Command4_Click()
If Text5.Text = "" Then GoTo msg
If InStr(1, Text5.Text, "\") Then
merge Text4.Text, Text5.Text
MsgBox "Segments merged into file " & Text5.Text
Else
MsgBox "You didn't put the path of the output file", , "hey!"
End If
GoTo rrr
msg:
MsgBox "You forgot to input somethin'", , "tsk,tsk,tsk"
rrr:
End Sub

