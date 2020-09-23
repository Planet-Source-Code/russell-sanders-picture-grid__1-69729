VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Private Type UndoElement
    spX                 As Long 'we need the left and top of the shape control
    spY                 As Long
    spX2                As Long 'Width And Height Of the Selection Square
    spY2                As Long
    x                   As Long 'point on the source
    Y                   As Long
    X2                  As Long 'left and top of the source picture at time of the edit
    Y2                  As Long
    Color               As Long
End Type
Private stackValue()    As UndoElement
Private level           As Double
Public WithEvents desp As Menu

Public Property Get stackLevel() As Double
    On Error Resume Next
        If level < 1 Then
            stackLevel = 0
        Else
            stackLevel = UBound(stackValue) 'set the stack level to the greatest index
        End If
End Property

Private Sub Class_Initialize()
    'level = 0          'set the stacklevel to nothing
    ClearUndo
End Sub

Public Sub push(spX As Long, spY As Long, spX2 As Long, spY2 As Long, x As Long, Y As Long, X2 As Long, Y2 As Long, Color As Long)
    level = level + 1         ' add one to the stack level
    'FrmGrid.UndoClear
    If Not desp Is Nothing Then desp.Caption = "Undo Buffer Level: " & level & " Select to Clear."
    ReDim Preserve stackValue(level) 'redim the array to the new level
    stackValue(level).Color = Color  'add info to the array veriables
    stackValue(level).spX = spX
    stackValue(level).spY = spY
    stackValue(level).spX2 = spX2
    stackValue(level).spY2 = spY2
    stackValue(level).x = x
    stackValue(level).Y = Y
    stackValue(level).X2 = X2
    stackValue(level).Y2 = Y2
End Sub

Public Function pop() As Long()
Dim res As Variant
Dim p() As Long
    If level > 0 Then               'is any thing in the stack(array)
        ReDim p(8)
        p(0) = stackValue(level).Color      'set the verable to the item with the greatest index
        p(1) = stackValue(level).Y
        p(2) = stackValue(level).x
        p(3) = stackValue(level).Y2
        p(4) = stackValue(level).X2
        p(5) = stackValue(level).spY2
        p(6) = stackValue(level).spX2
        p(7) = stackValue(level).spY
        p(8) = stackValue(level).spX
        pop = p
        level = level - 1           'take one from the stack level
        ReDim Preserve stackValue(level)  'remove that item from the array
    End If
    If Not desp Is Nothing Then desp.Caption = "Undo Buffer Level: " & level & " Select to Clear."
End Function

Public Sub ClearUndo()
    ReDim stackValue(0)
    level = 0
    If Not desp Is Nothing Then desp.Caption = "Undo Buffer Level: " & level & " Select to Clear."
End Sub

Private Sub desp_Click()
    ClearUndo
End Sub
