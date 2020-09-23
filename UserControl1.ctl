VERSION 5.00
Begin VB.UserControl LoSlider 
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   ScaleHeight     =   885
   ScaleWidth      =   3465
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2760
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "LoSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Change()
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If X > 10 And X < (UserControl.Width - 10) Then
Line2.X1 = X
Line2.X2 = X
RaiseEvent Change
End If
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 10 And X < (UserControl.Width - 10) Then
If Button = 1 Then
Line2.X1 = X
Line2.X2 = X
RaiseEvent Change
End If
End If
End Sub
Private Sub UserControl_Resize()
Line1.X2 = UserControl.Width
End Sub
Property Get Value()
Value = Int((Line2.X2) / (UserControl.Width - 15) * 100)
End Property
Property Set Value(a)
Line2.X2 = a + 15
Line2.X1 = a + 15
End Property

Property Let Value(a)
Line2.X2 = a + 15
Line2.X1 = a + 15
End Property

