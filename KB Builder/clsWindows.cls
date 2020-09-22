VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarCollection As Collection
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Function ByHandle(Handle As String) As clsWindow
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim myW As clsWindow
    rsTot = mvarCollection.Count
    For rsCnt = 1 To rsTot
        Set myW = mvarCollection(rsCnt)
        If Handle = myW.Handle Then
            Set ByHandle = myW
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function ByPosition(Position As Long) As clsWindow
    On Error Resume Next
    Set ByPosition = New clsWindow
    Set ByPosition = mvarCollection(Position)
    Err.Clear
End Function
Public Sub Add(ByVal strHandle As String, ByVal strClass As String, ByVal strTitle As String)
    On Error Resume Next
    Dim clsW As clsWindow
    Set clsW = New clsWindow
    strTitle = Trim$(strTitle)
    strClass = Trim$(strClass)
    If Len(strTitle) > 0 Then
        strHandle = Format$(strHandle, "0000000000")
        clsW.Class = strClass
        clsW.Handle = strHandle
        clsW.Parent = GetParent(Val(strHandle))
        clsW.Title = strTitle
        mvarCollection.Add clsW, strHandle
    End If
    Err.Clear
End Sub
Public Property Get Count() As Long
    On Error Resume Next
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Count
    Count = mvarCollection.Count
    Err.Clear
End Property
Private Sub Class_Initialize()
    On Error Resume Next
    Set mvarCollection = New Collection
    Err.Clear
End Sub
