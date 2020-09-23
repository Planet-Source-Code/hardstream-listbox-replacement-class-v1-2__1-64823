VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' __________________________________________________________________________
'/ListClass version 1.2                                                     \
'\==========================================================================/
'|Author: HardStream Group                                                  |
'|Date  : 04-01-2006                                                        |
'|Name  : cListClass                                                        |
'|Verson: 1.2                                                               |
'/==========================================================================\
'|Version 1.0:                                                              |
'|First version of this list class. This class is meant to be a replacement |
'|for the listbox. If you're writing an application with forms, this won't  |
'|be very helpful, but if you're writing a DLL, a console application or    |
'|any different type of program without visible objects, this can be a very |
'|helpful class.                                                            |
'\==========================================================================/
'|Version 1.1                                                               |
'|The class has been updated with a sorting function, so items can be sorted|
'|alphabetically.                                                           |
'|For the sorting, I've used Rde's QuickSort 2.1 module. Link:              |
'|http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=64095&lngWId=1    |
'|                                                                          |
'|I've also added the ability to add items somewhere in the middle of the   |
'|list using an index, just like in the VB listbox.                         |
'|You can now automaticly put the list in a textbox and in a listbox.       |
'|If you want to put it in other controls, you can use the same technique,  |
'|but apply it to the control you want :D                                   |
'|The class now also has events, so you can call the class as a WithEvents. |
'/==========================================================================\
'|WithEvents explanation                                                    |
'|                                                                          |
'|In in the Declarations part of the form/module etc. code, you put:        |
'|Private WithEvents vList As cListClass                                    |
'|                                                                          |
'|In the Form_Load/UserControl_Initialize etc. part you put:                |
'|Set vList = New cListClass                                                |
'|                                                                          |
'|Then you can call events like with any other (visible) control.           |
'|If you want to call the ItemAdd event, just copypaste the following code: |
'|Private Sub vList_ItemAdd(Index As Long)                                  |
'|                                                                          |
'|Just make sure you're using the same datatypes as used in the Events in   |
'|the class. Using different datatypes may result in error.                 |
'\==========================================================================/
'|Version 1.2                                                               |
'|New added features:                                                       |
'|1) Search: search through the listbox to get the index of an item.        |
'|2) Import and export functions, so you can easily import from and export  |
'|   to a file.                                                             |
'|3) Encryption/decryption so you can store private data without worrying   |
'|   about other people reading it.                                         |
'|   The encrypting/decrypting routine still needs work, check the example  |
'|   program to see for yourself what goes wrong. I'm not an expert on      |
'|   encrypting/decrypting, but because I wanted my own code as much as     |
'|   possible (but there's no way I'm going to write the sorting stuff :P), |
'|   I just wanted to have my own encrypting/decrypting code, but it's not  |
'|   working very well, so if anyone else could do this for me... :)        |
'\__________________________________________________________________________/
' |ps.                                                                     |
' |1) Sorry for my bad English, it's not my native language :P             |
' |2) If you find any error(s), please report them, I'll try to fix them   |
' \________________________________________________________________________/

'Events
Event ItemAdd(Index As Long)
Event ItemRemove(Index As Long)
Event Clear()
Event Sort(SortTime As Double, Success As Boolean)
Event ListCreated(Ctl As Control)
Event Error(Number As Byte, Description As String)
Event SearchFinish(Result As Long)
Event ExportDone(Filename As String, BytesWritten As Long)
Event ImportDone(Filename As String, BytesRead As Long)
Event EncodeDone()
Event DecodeDone()

'Private APIs
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'Private variables
Private sList As String
Private cFrequency As Currency
Private cStart As Currency
Private cNow As Currency
Private bAvailable As Boolean

'\/Stopwatch related
'If you want to remove the stopwatch from the class, you have to remove all
'API calls and the following 3 functions.
'You can use the Find function ([Ctrl]+[F]) in VB to find all calls to
'stopwatch related functions in other functions.
'You can also remove all private variable, except sList.
Private Sub Class_Initialize()
bAvailable = (QueryPerformanceFrequency(cFrequency) <> 0)
If (bAvailable) Then Debug.Print "Ticks/sec: "; cFrequency * 10000 Else Debug.Print "Performance Counter not available"
End Sub

'Reset the stopwatch
Private Sub sw_Reset()
QueryPerformanceCounter cStart
End Sub

'Get the elapsed time
Private Function sw_Elapsed() As Double
Call QueryPerformanceCounter(cNow)
If (bAvailable) Then sw_Elapsed = 1000 * (cNow - cStart) / cFrequency
End Function
'/\Stopwatch related

Private Function PrettySort(StringArray() As String, Optional CapsFirst As Boolean = True, Optional ReverseOrder As Boolean = False) As String
On Error GoTo FndErr
Dim LB As Long, UB As Long
Dim p_CapsFirst As Boolean
Dim p_eDirection As Long

p_CapsFirst = (CapsFirst = 1)
p_eDirection = (ReverseOrder * -2) + 1 '0 >> 1 : 1 >> -1
LB = LBound(StringArray)
UB = UBound(StringArray)
strPrettySort StringArray, LB, UB, p_CapsFirst, p_eDirection
PrettySort = Join(StringArray, vbCrLf)
Exit Function

FndErr:
    PrettySort = ""
End Function

'Count the number of occurrences of a (set of) character(s) in an expression
Private Function Occur(Expression As String, Optional Delimiter = vbCrLf) As Long
Occur = UBound(Split(Expression, Delimiter))
End Function

'Read one line from a multiline string
Private Function ReadLine(Expression As String, Line As Long) As String
If Line > Occur(Expression) Then 'Line number is too high
    ReadLine = ""
    Exit Function
        ElseIf Line < 0 Then 'Line number is too low
    ReadLine = ""
    Exit Function
End If

ReadLine = Split(Expression, vbCrLf)(Line) 'Read the line
End Function

'Add a new line to an expression
Private Sub AddToString(Storage As String, Expression As String)
If Trim(Storage) = "" Then Storage = Expression Else Storage = Storage & vbCrLf & Expression
End Sub

'Add an item to the list
Sub AddItem(Item As String, Optional Index = -1)
Dim Part(0 To 1) As String
Dim i As Long
Dim ReturnIndex As Long

If IsNumeric(Index) = False Then 'The index isn't numeric
    Err.Raise 102, "ListClass", "Non-numeric index"
    RaiseEvent Error(102, "Non-numeric index")
    Exit Sub
End If

If Trim(sList) = "" Then 'The list is empty
    sList = Item
    ReturnIndex = 0
    GoTo AddDone
End If

If Index = -1 Then 'Add the item as the last item
    sList = sList & vbCrLf & Item
    ReturnIndex = Occur(sList)
        ElseIf Index = 0 Then 'Add the item as the first item
    sList = Item & vbCrLf & sList
    ReturnIndex = 0
        ElseIf Index > 0 Then 'Add somewhere in the middle
    If Index = Occur(sList) + 1 Then 'Add the item as the last item
        sList = sList & vbCrLf & Item
        ReturnIndex = Index 'Return the correct index
            ElseIf Occur(sList) + 1 < Index Then 'The list is smaller than index+1, so just add as last item and return corrected index
        sList = sList & vbCrLf & Item
        ReturnIndex = Occur(sList) - 1
            Else 'Really add the item somewhere in the middle of the list
        For i = 0 To Index - 1 'Loop through the first part of the list
            AddToString Part(0), ReadLine(sList, i)
        Next i
        For i = Index To Occur(sList) 'Loop through the second part of the list
            AddToString Part(1), ReadLine(sList, i)
        Next i
        
        sList = Part(0) & vbCrLf & Item & vbCrLf & Part(1) 'Add the item
        ReturnIndex = Index 'Return the correct index
    End If
End If

AddDone:
    RaiseEvent ItemAdd(ReturnIndex)
End Sub

'Remove an item from the list
Sub RemoveItem(Index As Long)
Dim Tmp As String
Dim i As Long

If IsNumeric(Index) = False Then 'Non-numeric index
    Err.Raise 102, "ListClass", "Non-numeric index"
    RaiseEvent Error(102, "Non-numeric index")
    Exit Sub
End If

If Index = 0 Then 'Remove the first item
    For i = 1 To Occur(sList)
        AddToString Tmp, ReadLine(sList, i)
    Next i
        ElseIf Index < 0 Then 'Invalid index
        ElseIf Index > 0 Then 'Remove a later item
    For i = 0 To Index - 1
        AddToString Tmp, ReadLine(sList, i)
    Next i
    
    If Index < Occur(sList) Then 'The removed item isn't the last item in the list
        For i = Index + 1 To Occur(sList)
            AddToString Tmp, ReadLine(sList, i)
        Next i
    End If
End If

sList = Tmp 'Save
Tmp = "" 'Clear up

RaiseEvent ItemRemove(Index)
End Sub

'Clear the list
Sub Clear()
sList = ""
RaiseEvent Clear
End Sub

'Get the text of a list item
'Set Index=-1 to get the complete list
Function Text(Optional Index As Long = -1) As String
If Index = vbNullString Then
    Text = ""
    Exit Function
End If

If IsNumeric(Index) = False Then 'Non-numeric index
    Err.Raise 102, "ListClass", "Non-numeric index"
    RaiseEvent Error(102, "Non-numeric index")
    Text = ""
    Exit Function
End If

If Index >= 0 Then Text = ReadLine(sList, Index)
If Index = -1 Then Text = sList
End Function

'Get the number of items
Function ListCount()
ListCount = Occur(sList) + 1
End Function

'Sort the list
Function Sort(Optional CapsFirst As Boolean = True, Optional ReverseOrder As Boolean = False) As Double
Dim tSort() As String
Dim Result As String, OldList As String
Dim SortTime As Double

sw_Reset 'Reset the stopwatch

If Trim(sList) = "" Then GoTo Pointless 'Check if the list string is not empty
If Occur(sList) = 0 Then GoTo Pointless 'Check if there really is a multi-line string, not just a single-line string
Result = "" 'Make sure the result buffer is empty
tSort = Split(sList, vbCrLf) 'Create the array
Result = PrettySort(tSort, CapsFirst, ReverseOrder) 'Sort and store in a temporary buffer
OldList = sList 'Store the old list (unsorted). In case of an error, this is restored, so the list won't be lost

If Trim(Result) <> "" Then
    sList = Result 'Apply
    Result = "" 'Clean the temporary buffer
    OldList = "" 'Clean the old list buffer
    SortTime = sw_Elapsed 'Return elapsed time since sort start
    Sort = SortTime
    RaiseEvent Sort(SortTime, True)
    Exit Function
        Else
    SortTime = sw_Elapsed
    RaiseEvent Sort(SortTime, False)
End If

Pointless: 'Unnecessary to sort, but not error worthy
    Sort = 0
    Exit Function

FndErr: 'An error occurred
    RaiseEvent Sort(0, False)
    Err.Raise 101, "ListClass", "Sorting failed"
    RaiseEvent Error(101, "Sorting failed")
    sList = OldList 'Restore the old list (in case the corrupted sorted list has been applied)
    Sort = 0
End Function

'Put all items in a control.
'Currently there is only support for Textbox, Listbox and Combobox controls, but you can easily add support for more controls
Sub PutInControl(Ctl As Control, Optional AutoClean As Boolean = True)
Dim i As Long

'Auto clean (if true)
If AutoClean = True Then
    If TypeOf Ctl Is TextBox Then Ctl.Text = "" 'Textbox
    If TypeOf Ctl Is ListBox Or TypeOf Ctl Is ComboBox Then Ctl.Clear 'Listbox/Combobox
End If

For i = 0 To Occur(sList) 'Loop through all items
    If TypeOf Ctl Is TextBox Then 'Using a textbox
        If Trim(Ctl.Text) = "" Then Ctl.Text = ReadLine(sList, i) Else Ctl.Text = Ctl.Text & vbCrLf & ReadLine(sList, i)
            ElseIf TypeOf Ctl Is ListBox Or TypeOf Ctl Is ComboBox Then 'Using a listbox or combobox (doesn't matter in this case, because they both have the necessary functions)
        Ctl.AddItem ReadLine(sList, i)
    End If
Next i

RaiseEvent ListCreated(Ctl)
End Sub

'Search for an item in the list
Function SearchItem(Expression As String, Optional ExactMatch As Boolean = False, Optional CaseMatch As Boolean = False) As Long
Dim i As Long
Dim Tmp(0 To 1) As String

If CaseMatch = True Then Tmp(1) = Expression Else Tmp(1) = LCase(Expression)

For i = 0 To Occur(sList) 'Loop through the list
    Tmp(0) = ReadLine(sList, i) 'Read the current line
    
    If ExactMatch = True Then 'Exact match required
        If Tmp(0) = Expression Then 'Match
            SearchItem = i
            RaiseEvent SearchFinish(i)
            Exit Function
                Else 'No match
            GoTo TryNext
        End If 'Match
            Else 'Exact match not required
        If InStr(1, Tmp(0), Expression) > 0 Then 'Match
            SearchItem = i
            RaiseEvent SearchFinish(i)
            Exit Function
                Else 'No match
            GoTo TryNext
        End If
    End If 'ExactMatch
TryNext:
Next i

'The code hereunder will only be executed if there are no results
SearchItem = -1
RaiseEvent SearchFinish(-1)
End Function

'Export the list to a file
Sub Export(Filename As String)
Dim FF As Long

FF = FreeFile
Open Filename For Binary Access Write As #FF 'Open the file for writing
    Put #FF, , sList 'Write to the file
Close #FF 'Close the file
RaiseEvent ExportDone(Filename, Len(sList))
End Sub

'Import list from a file
Sub Import(Filename As String)
Dim FF As Long

FF = FreeFile
Open Filename For Binary Access Read As #FF 'Open the file for reading
    sList = Input(LOF(FF), FF) 'Read from the file
Close #FF 'Close the file
RaiseEvent ImportDone(Filename, Len(sList))
End Sub

'Encode the list
Sub Encode()
Dim i As Long, Tmp As String

For i = 0 To Occur(sList)
    Tmp = Tmp & vbCrLf & Encrypt(ReadLine(sList, i))
Next i
sList = Tmp
Tmp = ""

RaiseEvent EncodeDone
End Sub

'Decode the list
Sub Decode()
Dim i As Long, Tmp As String

For i = 0 To Occur(sList)
    Tmp = Tmp & vbCrLf & Decrypt(ReadLine(sList, i))
Next i
sList = Tmp
Tmp = ""

RaiseEvent DecodeDone
End Sub
