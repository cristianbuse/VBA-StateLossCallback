VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StateLossCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA StateLossCallback
''' -----------------------------------------------------
''' https://github.com/cristianbuse/VBA-StateLossCallback
''' -----------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2023 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit

'*******************************************************************************
'' Each instance of this class will make a call back to the provided macro even
''   if state is lost. A 'fake' interface is used and an unmanaged instance of
''   that interface (m_data) will do the call when VB calls it's Release method.
'' Reference count is not affected and there are no memory leaks. Furthermore,
''   there are no crashes either as the extra interface calls 'wraps' the calls.
'' The callback can be:
'' - a macro name passed to 'InitByMacroName'. Arguments are also allowed
'' - an AddressOf pointer passed to 'InitByAddress'. One text argument allowed
'' For a more detailed explanation read:
'' - https://github.com/cristianbuse/VBA-StateLossCallback/blob/master/README.md
'*******************************************************************************

'Turn the following compiler constant to True if already using:
'https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/LibMemory.bas
#Const USE_LIB_MEMORY = False

#If USE_LIB_MEMORY = 0 Then
#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
#If Win64 Then
    Const PTR_SIZE As Long = 8
#Else
    Const PTR_SIZE As Long = 4
#End If
#End If

Private Enum CallType
    ctNone = 0
    ctAddress = 1
    ctName = 2
End Enum

Implements MSForms.DataObject 'Used to trap state loss
Implements stdole.IFontEventsDisp 'Used to call back by method memory address

Private m_data As MSForms.DataObject
Private m_callbackPtr As LongPtr
Private m_callbackArg As String
Private m_callbackName As String
Private m_args() As Variant
Private m_callType As CallType

#If USE_LIB_MEMORY = 0 Then
Private Property Get MemLongPtr(ByVal memAddress As LongPtr) As LongPtr
    CopyMemory MemLongPtr, ByVal memAddress, PTR_SIZE
End Property
Private Property Let MemLongPtr(ByVal memAddress As LongPtr, ByVal newValue As LongPtr)
    CopyMemory ByVal memAddress, newValue, PTR_SIZE
End Property
#End If

'*******************************************************************************
'Callback should be of the following form:
' - Private/Public Sub MethodName(ByVal instancePtr As LongPtr, ByVal argText As String)
' - Private/Public Sub MethodName(ByVal instancePtr As LongPtr)
' - Private/Public Sub MethodName()
'Init like:
' - .InitByAddress AddressOf MethodName, "Optional Text"
'*******************************************************************************
#If Win64 Then
Public Sub InitByAddress(ByVal memAddress As LongLong, Optional ByVal argText As String)
#Else
Public Sub InitByAddress(ByVal memAddress As Long, Optional ByVal argText As String)
#End If
    If memAddress = 0 Then Exit Sub
    m_callbackPtr = memAddress
    m_callbackArg = argText
    Init ctAddress
End Sub

'*******************************************************************************
'Init like:
' - .InitByMacroName callbackName, arg1, arg2, ...
'The number of arguments passed to this method needs to match the number of
'  arguments expected by the callback method or the call will fail
'Important! Try to avoid passing Objects as arguments as there is no guarantee
'  that the instances are still alive by the time the call back is made
'*******************************************************************************
Public Sub InitByMacroName(ByVal macroName As String, ParamArray args() As Variant)
    If LenB(macroName) = 0 Then Exit Sub
    m_callbackName = macroName
    m_args = args
    Init ctName
End Sub

'*******************************************************************************
'Initialize the unmanaged interface instance responsible for the trigger
'*******************************************************************************
Private Sub Init(ByVal cType As CallType)
    If m_callType <> ctNone Then GoTo AfterInit
    '
    Dim dataInterfacePtr As LongPtr:  dataInterfacePtr = VarPtr(DataObject) + PTR_SIZE
    Dim dataVirtTablePtr As LongPtr:  dataVirtTablePtr = MemLongPtr(dataInterfacePtr)
    Dim dataReleaseAddr As LongPtr:   dataReleaseAddr = dataVirtTablePtr + PTR_SIZE * 2
    Dim dataClearAddr As LongPtr:     dataClearAddr = dataVirtTablePtr + PTR_SIZE * 7
    '
    MemLongPtr(VarPtr(m_data)) = dataInterfacePtr 'Unmanaged reference
    MemLongPtr(dataReleaseAddr) = MemLongPtr(dataClearAddr)
AfterInit:
    m_callType = cType
End Sub

'*******************************************************************************
'Useful when used in With New StateLossCallback ... End With constructs
'*******************************************************************************
Public Function Self() As StateLossCallback
    Set Self = Me
End Function

'*******************************************************************************
'Called when the m_data instance is being cleared by VB
'*******************************************************************************
Private Sub DataObject_Clear()
    If Not m_data Is Nothing Then Exit Sub 'Ignore external logic attempts
    Select Case m_callType
    Case CallType.ctAddress
        Dim eFont As stdole.IFontEventsDisp: Set eFont = Me
        MemLongPtr(MemLongPtr(ObjPtr(eFont)) + PTR_SIZE * 7) = m_callbackPtr
        eFont.FontChanged m_callbackArg
    Case CallType.ctName
        RunMacro 'No need for Error Handling
    End Select
End Sub
Private Sub RunMacro(Optional ByVal Missing As Variant)
    Const maxRunArgs As Long = 30
    Dim argsCount As Long: argsCount = UBound(m_args) + 1
    Dim i As Long
    '
    ReDim Preserve m_args(0 To maxRunArgs - 1)
    For i = argsCount To UBound(m_args)
        m_args(i) = Missing
    Next i
    '
    Application.Run m_callbackName _
        , m_args(0), m_args(1), m_args(2), m_args(3), m_args(4) _
        , m_args(5), m_args(6), m_args(7), m_args(8), m_args(9) _
        , m_args(10), m_args(11), m_args(12), m_args(13), m_args(14) _
        , m_args(15), m_args(16), m_args(17), m_args(18), m_args(19) _
        , m_args(20), m_args(21), m_args(22), m_args(23), m_args(24) _
        , m_args(25), m_args(26), m_args(27), m_args(28), m_args(29)
End Sub

'The following methods are required by the 2 implemented interfaces but are not used
Private Function DataObject_GetFormat(ByVal Format As Variant) As Boolean: End Function
Private Sub DataObject_GetFromClipboard(): End Sub
Private Function DataObject_GetText(Optional ByVal Format As Variant) As String: End Function
Private Sub DataObject_PutInClipboard(): End Sub
Private Sub DataObject_SetText(ByVal Text As String, Optional ByVal Format As Variant): End Sub
Private Function DataObject_StartDrag(Optional ByVal OKEffect As Variant) As MSForms.fmDropEffect: End Function
Private Sub IFontEventsDisp_FontChanged(ByVal PropertyName As String): End Sub
