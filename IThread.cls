VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'This is a simple implementation of how you can do "threads" in VB
'It can off course be expanded and specialized to whatever you need...
'variants are used only to show simple example, for better performance
'you should declare as 'type' for events and subs other than variant

Event TaskComplete(ByVal FunctionName As String, Data As Variant)

'local variable(s) to hold property value(s)
Private mvarArg1 As String 'local copy
Private mvarArg2 As String 'local copy

'=========================================================
'IMPLEMENTATION CODE
'=========================================================
'note that each object in this activex exe takes up one thread
'because it is set in project properties to use "Thread per object"
'since the timer is an object, it also requires a thread, therefore
'each 'thread' object in your app will take up two threads a piece

Private WithEvents Synchronizer As CLiteTimer
Attribute Synchronizer.VB_VarHelpID = -1

'this is a list to hold tasks to be performed
Private Tasks As Collection

Function StartTask(ByVal strFunctionName As String, ByVal blnAsync As Boolean, ByVal enmCallType As VbCallType, ParamArray varData() As Variant) As Variant

     'if this (blnAsync) is true, this function will return immediately
     'leaving your main program free to do other things while processing this call
     ' when finished your function will raise an event
     'letting the main program know that it is done working
    If blnAsync Then
        
        'add it to the task list
        Tasks.Add strFunctionName
        'check to make sure timer is running
        If Not Synchronizer.Enabled Then
            'set timer to respond/tick within 1 millisecond
            Synchronizer.Interval = 1
            Synchronizer.Enabled = True
        End If

    Else
        
        'if not asynchronous then we act like a normal function call
        If enmCallType = VbMethod Then
            StartTask = CallByName(Me, strFunctionName, enmCallType, varData(0))
        Else
            CallByName Me, strFunctionName, enmCallType, varData(0)
        End If
        
    End If

End Function

Private Sub Class_Initialize()

 'initialize the tasklist
 Set Tasks = New Collection
 
 'initialize the timer
 Set Synchronizer = New CLiteTimer
 
End Sub

Private Sub Synchronizer_Timer()
   
   'the timer has fired so we must check the tasklist to see if any tasks
   'are waiting to be done
    If Tasks.Count > 0 Then
        Synchronizer.Enabled = False
        'do the first added task in the list
        CallByName Me, Tasks(1), VbMethod
        'remove it
        Tasks.Remove 1
    End If
    
    'if no tasks left then disable timer
    If Tasks.Count > 0 Then
        Synchronizer.Enabled = True
    End If
   
End Sub


'======================================================================
'WORKER FUNCTIONS
' write all your code functions and properties below!
'======================================================================

'properties are used to persist data needed by worker subs, not ideal
'but it works. It is a pain in the neck to persist ParamArray data otherwise.
'you can of course, change the StartTask() function to be specialized to each
'worker sub you add below

Public Property Let Arg1(ByVal vData As String)
    mvarArg1 = vData
End Property
Public Property Let Arg2(ByVal vData As String)
    mvarArg2 = vData
End Property

Public Sub DoSomeLongWork()

  Dim X As Long
    
    'this function uses property "Arg1" and "Arg2" as its parameters (or arguments)
    'this would be similar to calling it like so: DoSomeLongWork Arg1, Arg2
    
    'lets do some heavy duty string processing.
    'this is what "multithreading" is all about!
    mvarArg1 = Left$(mvarArg1, Len(mvarArg2))

    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    X = 0
    Do While X < Len(mvarArg1)
        X = X + 1
        Mid(mvarArg1, X, 1) = Mid$(mvarArg2, X, 1)
    Loop
    
    'raise event that we are done and show how
    'many characters were in the string
    RaiseEvent TaskComplete("DoSomeLongWork", X)

End Sub

'here is an example of not having to persist data for a call, you can execute this all in one call
'however COM will synchronize the call an your app will appear frozen if it takes along time
'even tho it is in another thread
Public Function DoMoreWork(Value As String) As String

  On Error Resume Next
    'shows/simulates a synchronous method that returns data
    DoMoreWork = "Thread Has Processed: " & CStr(Value)

End Function

