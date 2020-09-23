VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ListClass Test"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog COM 
      Left            =   4920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstList 
      Height          =   7815
      IntegralHeight  =   0   'False
      Left            =   7200
      TabIndex        =   6
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtList 
      Height          =   7815
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   360
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test the class"
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdDecode 
         Caption         =   "Decode"
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdEncode 
         Caption         =   "Encode"
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkMatchCase 
         Caption         =   "Case match"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CheckBox chkMatchExact 
         Caption         =   "Exact match"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   6840
         Width           =   735
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "Test"
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton cmdCredits 
         Caption         =   "About the sorting algorithm (please click here!!)"
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   7320
         Width           =   1575
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "Get"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox txtGet 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Text            =   "0"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkCaps 
         Caption         =   "Caps first"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CheckBox chkRev 
         Caption         =   "Reverse alphabet"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtDel 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "0"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtAddIndex 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "-1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtAddText 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Line Line8 
         X1              =   1920
         X2              =   3480
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line7 
         X1              =   1800
         X2              =   1800
         Y1              =   120
         Y2              =   8160
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Expression:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   5640
         Width           =   810
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   1680
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   1680
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Index:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   4560
         Width           =   435
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   1680
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   1680
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   1680
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Index:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   435
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1680
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Index:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Label lblMsg 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Events are shown here :D"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   8160
      Width           =   10815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "In a listbox:"
      Height          =   195
      Left            =   7200
      TabIndex        =   7
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "In a textbox:"
      Height          =   195
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private APIs
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Private event references
Private WithEvents vList As cListClass
Attribute vList.VB_VarHelpID = -1

'\/List events
Private Sub vList_ItemAdd(Index As Long)
'An item has been added
'Index represents the index of the new item
lblMsg.Caption = "New item index: " & Index
End Sub

Private Sub vList_ItemRemove(Index As Long)
'An item has removed
'Index represents the index of the removed item
lblMsg.Caption = "Removed item index: " & Index
End Sub

Private Sub vList_Clear()
'The list has been cleared
lblMsg.Caption = "The list has been cleared"
End Sub

Private Sub vList_Sort(SortTime As Double, Success As Boolean)
'The list has been sorted
'SortTime represents the time it took to sort the list (in seconds)
'Success represents the success of the sorting. False if error
lblMsg.Caption = "List sorted in " & SortTime & " seconds. Success = " & Success
End Sub

Private Sub vList_ListCreated(Ctl As Control)
'The list has been transported to a control
'Ctl represents the control the list has been transported to
'lblMsg.Caption = "List created in " & Ctl.Name
End Sub

Private Sub vList_Error(Number As Byte, Description As String)
'An error occurred in the class
'Number represents the identification number of the error
'Description represents the default description of the error
lblMsg.Caption = "Error " & Number & ": " & Description
End Sub

Private Sub vList_SearchFinish(Result As Long)
'The search finished
'Result represents the result (index) of the search
End Sub

Private Sub vList_ExportDone(Filename As String, BytesWritten As Long)
'The list has been exported to a file
'Filename represents the name of the destination file
'BytesWritten is the size of the file in bytes
End Sub

Private Sub vList_ImportDone(Filename As String, BytesRead As Long)
'The list has been imported from a file
'Filename represents the name of the source file
'BytesRead is the size of the file in bytes
End Sub

Private Sub vList_EncodeDone()
'The list has been encoded
End Sub

Private Sub vList_DecodeDone()
'The list has been decoded
End Sub
'/\List events

'Update all controls
Private Sub Update()
vList.PutInControl txtList, True
vList.PutInControl lstList, True
Me.Caption = vList.ListCount
End Sub

'\/Button events
Private Sub cmdAdd_Click()
vList.AddItem txtAddText.Text, txtAddIndex.Text
Update
End Sub

Private Sub cmdClear_Click()
vList.Clear
Update
End Sub

Private Sub cmdDel_Click()
vList.RemoveItem txtDel.Text
Update
End Sub

Private Sub cmdSort_Click()
vList.Sort chkCaps.Value, chkRev.Value
Update
End Sub

Private Sub cmdGet_Click()
MsgBox vList.Text(txtGet.Text)
End Sub

Private Sub cmdCredits_Click()
MsgBox "Please take a moment to vote for Rde at" & vbCrLf & "http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=64095&lngWId=1" & vbCrLf & vbCrLf & "He wrote the sorting code :D (thank you Rde, if you read this)" & vbCrLf & "The link is also in the class."
End Sub

Private Sub cmdSearch_Click()
lstList.ListIndex = vList.SearchItem(txtSearch.Text, chkMatchExact.Value, chkMatchCase.Value)
End Sub

Private Sub cmdExport_Click()
On Error GoTo FndErr

With COM
    .CancelError = True
    .DialogTitle = "Export list"
    .Filename = ""
    .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    .Flags = &H2 'Warn before overwrite
    .ShowSave
End With
If Trim(COM.Filename) = "" Then Exit Sub 'No filename returned

vList.Export COM.Filename

FndErr: 'An error occurred
End Sub

Private Sub cmdImport_Click()
On Error GoTo FndErr

With COM
    .CancelError = True
    .DialogTitle = "Import list"
    .Filename = ""
    .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    .Flags = &H1000 'File MUST exist
    .ShowOpen
End With
If Trim(COM.Filename) = "" Then Exit Sub 'No filename returned

vList.Import COM.Filename
Update

FndErr: 'An error occurred
End Sub

Private Sub cmdEncode_Click()
vList.Encode
Update
End Sub

Private Sub cmdDecode_Click()
vList.Decode
Update
End Sub
'/\Button events

'\/Form events
Private Sub Form_Load()
Set vList = New cListClass
End Sub

Private Sub Form_Unload(Cancel As Integer)
lblMsg.Caption = "If you found any errors, please report them :D"
Me.Caption = "Please read the message label, this form will close automaticly in 5 seconds"
DoEvents
Sleep 5000
End Sub
'/\Form events
