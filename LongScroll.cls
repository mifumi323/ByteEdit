VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLongScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'HScroll‚É‚àVScroll‚É‚à‘Î‰ž‚³‚¹‚é‚½‚ßObject‚È‚Ì‚¾
Dim m_Target As Object

Dim m_Value As Long
Dim m_Max As Long

Public Property Get ScrollBar() As Object
    Set ScrollBar = m_Target
End Property

Public Property Let ScrollBar(ByVal NewValue As Object)
    Set m_Target = NewValue
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal NewValue As Long)
    m_Target.Max = IIf(NewValue < 32767, NewValue, 32767)
    m_Max = NewValue
End Property

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Long)
    If ToIntValue(m_Value) <> m_Target.Value Then m_Target.Value = ToIntValue(NewValue)
    m_Value = NewValue
End Property

Public Sub Update()
    If ToIntValue(m_Value) <> m_Target.Value Then m_Value = ToLongValue(m_Target.Value)
End Sub

Public Function ToLongValue(ByVal IntValue As Integer) As Long
    If m_Target.Max Then
        ToLongValue = CLng(Int(CDbl(m_Max) / m_Target.Max * IntValue))
    Else
        ToLongValue = 0
    End If
End Function

Public Function ToIntValue(ByVal LongValue As Long) As Integer
    If m_Max Then
        ToIntValue = CLng(Int(CDbl(m_Target.Max) / m_Max * LongValue))
    Else
        ToIntValue = 0
    End If
End Function

