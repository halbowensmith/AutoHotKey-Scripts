

copyIntegration()
{
WinActivate, ahk_class #32770
WinWaitActive, ahk_class #32770
ControlGetText, topIntegration, Edit3 ; Use the window found above.
Return topIntegration
}