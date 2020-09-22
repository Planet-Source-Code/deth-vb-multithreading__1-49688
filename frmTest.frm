VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Testing VB Multi Threading"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Simple Thread"
      Height          =   2265
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   6675
      Begin VB.CommandButton cmdExample2 
         Caption         =   "Example 2"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   270
         Width           =   1275
      End
      Begin VB.ListBox lstStatus 
         Height          =   1425
         Left            =   90
         TabIndex        =   2
         Top             =   720
         Width           =   6495
      End
      Begin VB.CommandButton cmdExample1 
         Caption         =   "Example 1"
         Height          =   375
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Caption         =   "Test"
         Height          =   285
         Left            =   3150
         TabIndex        =   4
         Top             =   360
         Width           =   3300
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: Active X Exe asynchronous multithread example
'Author: Lewis Miller (aka deth)
'Notes: read included readme.txt for testing instructions

'this is a thread object variable to hold a thread
'that is able to have events
Private WithEvents Thread As IThread
Attribute Thread.VB_VarHelpID = -1
'you can have as many threads as you want!
'Private WithEvents Thread2 As IThread
'Private WithEvents Thread3 As IThread

Dim Working As Boolean

Private Sub Form_Load()

    'initialize the thread
    Set Thread = New IThread

End Sub

Private Sub cmdExample1_Click()
    'this is an example of calling another thread synchronously
    'this isnt ideal because COM will block your app's main thread till
    'the method returns, example 2 is the ideal method
  Dim Temp As String
    Temp = "testing synchronous thread function"
    ThreadStatus "Calling Function 'DoMoreWork'"
    Temp = Thread.StartTask("DoMoreWork", False, VbMethod, Temp)
    ThreadStatus Temp

End Sub
Private Sub cmdExample2_Click()
Dim X As Long

    'This example shows the true power of asynchronous "threading" in ActiveX Exe.
    'Normally in vb your code that calls a function (even most event driven code)
    'will stop at the line of code that makes the call until that function is complete.
    'You will see here that the code does not stop at the line that makes the call.
    'The code can continue on, your program remains responsive to user input while the seperate
    'thread processes your work

    'set properties, blnAsync should be false
    Thread.StartTask "Arg1", False, VbLet, Space$(999999)
    Thread.StartTask "Arg2", False, VbLet, String$(999999, "Z")

    'start task, set blnAsync to true for it to do a task asynchronously
    Thread.StartTask "DoSomeLongWork", True, VbMethod
    
    'show status
    ThreadStatus "Separate Thread Is Now Busy Working."
    ThreadStatus "Notice That Program Is Not Frozen!!"

    'we can even do some work in this app while waiting on the thread to be done!!
    Working = True
    DoEvents
    Do While Working
        lblTest = "Waiting On Thread To Finish Working " & CStr(X)
        X = X + 1
        If X Mod 100 Then DoEvents
    Loop
    lblTest = "Two jobs at once in VB!!!"

End Sub

'this is an event from the active x exe, you can add any events you want similar to an ocx
Private Sub Thread_TaskComplete(ByVal FunctionName As String, Data As Variant)
  
    'the thread is done working!
    If FunctionName = "DoSomeLongWork" Then
      
        Working = False
      
        'process the finished data
        ThreadStatus "'" & FunctionName & "' has completed its task!!"
        ThreadStatus "Separate Thread Processed " & CStr(Data) & " Characters!"
  
    End If
  
    'add all the function names here that are in your "threaded" activeX exe control
  
  
End Sub


'this just adds text to the listbox for display purposes
Sub ThreadStatus(ByVal StatusText As String)

    lstStatus.AddItem StatusText
    lstStatus.ListIndex = lstStatus.NewIndex

End Sub

