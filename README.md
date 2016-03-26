Excel VBA ListBox Drag & Drop
=============================

![Preview](https://github.com/buntatsu/Excel-VBA_ListBox_DragAndDrop/blob/master/ScreenShot.png)

```vbnet
Option Explicit
Option Base 0

Private Const MOUSEBUTTON_LEFT As Long = 1
Private g_dragFrom As MSForms.ListBox

Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)

    If ListBox1.ListIndex < 0 Then Exit Sub
    If Button = MOUSEBUTTON_LEFT Then
        Call SetDragItem(ListBox1)
    End If
End Sub

Private Sub ListBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)

    If ListBox2.ListIndex < 0 Then Exit Sub
    If Button = MOUSEBUTTON_LEFT Then
        Call SetDragItem(ListBox2)
    End If
End Sub

Private Sub ListBox1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal DragState As Long, ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer)

    Cancel = True
    If g_dragFrom = ListBox1 Then
        Effect = fmDropEffectNone
    Else
        Effect = fmDropEffectMove
    End If
End Sub

Private Sub ListBox2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal DragState As Long, ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer)

    Cancel = True
    Effect = fmDropEffectMove
End Sub

Private Sub ListBox1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
    ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

    Cancel = True
    Effect = fmDropEffectMove

    Call AddDropItem(ListBox1, Data, Y)
    Call DeleteDragItem(g_dragFrom)
End Sub

Private Sub ListBox2_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, _
    ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

    Cancel = True
    Effect = fmDropEffectMove

    Call AddDropItem(ListBox2, Data, Y)
    Call DeleteDragItem(g_dragFrom)
End Sub

Private Sub SetDragItem(lb As MSForms.ListBox)
    Set g_dragFrom = lb
    Dim dataObj As New DataObject
    dataObj.SetText lb.Text
    Call dataObj.StartDrag(fmDropEffectMove)
End Sub

Private Sub AddDropItem(lb As MSForms.ListBox, dataObj As DataObject, Y As Single)
    lb.AddItem dataObj.GetText, FixDropIndex(lb, Y)
End Sub

Private Function FixDropIndex(lb As MSForms.ListBox, Y As Single) As Long
    Dim toIndex As Long

    With lb

    Select Case .Font.Name
    Case "Meiryo UI"
        toIndex = .TopIndex + Int(Y / (.Font.Size + 2.25))
    Case "MS UI Gothic"
        toIndex = .TopIndex + Int(Y / .Font.Size)
    Case Else
        toIndex = .TopIndex + Int(Y * 0.85 / .Font.Size)
    End Select

    If toIndex < 0 Then toIndex = 0
    If toIndex >= .ListCount Then toIndex = .ListCount

    End With

    FixDropIndex = toIndex
End Function

Private Sub DeleteDragItem(lb As MSForms.ListBox)
    Dim selIndex As Long
    With lb
    selIndex = .ListIndex
    .selected(selIndex) = False
    .RemoveItem selIndex
    End With
    Set g_dragFrom = Nothing
End Sub

Private Sub UserForm_Initialize()
    ListBox1.list = Array("Item1", "Item2", "Item3", "Item4", "Item5", "Item6", "Item7")
End Sub
```

