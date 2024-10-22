Attribute VB_Name = "modLvTimer"
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before making use of this code!
'Disclaimer: This is illegal if executed on real victims and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education purpose only.
'For more educational source codes please visit us http://www.digi77.com
'Author of this code W. Al Maawali Founder of  Eagle Eye Digital Solutions and Oman0.net can be reached via warith@digi77.com .

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

Option Explicit
' REQUIRED: copy & paste these few lines in any module of your project
' This is used by every lvButtons control as a replacement of the Timer control
' By doing it this way, each button control does NOT need an individual timer control
' The timer function is primarily used to determine when the mouse enters/leaves
' the button's physical region on the screen. See compiling information notes below.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, _
                                                                ByVal lpString As String) As Long

Public Function lv_TimerCallBack(ByVal hWnd As Long, _
                                 ByVal Message As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long
                                 
  Dim tgtButton As lvButtons_H

    ' when timer was intialized, the button control's hWnd
    ' had property set to the handle of the control itself
    ' and the timer ID was also set as a window property
    CopyMemory tgtButton, GetProp(hWnd, "lv_ClassID"), &H4
    Call tgtButton.TimerUpdate(GetProp(hWnd, "lv_TimerID"))  ' fire the button's event
    CopyMemory tgtButton, 0&, &H4                                    ' erase this instance

End Function


