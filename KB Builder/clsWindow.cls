VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarHandle As Long 'local copy
Private mvarClass As String 'local copy
Private mvarTitle As String 'local copy
'local variable(s) to hold property value(s)
Private mvarParent As Long 'local copy
'local variable(s) to hold property value(s)
Public Property Let Parent(ByVal vData As Long)
    On Error Resume Next
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Parent = 5
    mvarParent = vData
    Err.Clear
End Property
'    On Error Resume Next
'    'used when retrieving value of a property, on the right side of an assignment.
'    'Syntax: Debug.Print X.Parent
'    Parent = mvarParent
'    Err.Clear
'End Property
Public Property Let Title(ByVal vData As String)
    On Error Resume Next
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Title = 5
    mvarTitle = vData
    Err.Clear
End Property
Public Property Get Title() As String
    On Error Resume Next
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Title
    Title = mvarTitle
    Err.Clear
End Property
Public Property Let Class(ByVal vData As String)
    On Error Resume Next
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Class = 5
    mvarClass = vData
    Err.Clear
End Property
'    On Error Resume Next
'    'used when retrieving value of a property, on the right side of an assignment.
'    'Syntax: Debug.Print X.Class
'    Class = mvarClass
'    Err.Clear
'End Property
Public Property Let Handle(ByVal vData As Long)
    On Error Resume Next
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Handle = 5
    mvarHandle = vData
    Err.Clear
End Property
Public Property Get Handle() As Long
    On Error Resume Next
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Handle
    Handle = mvarHandle
    Err.Clear
End Property
