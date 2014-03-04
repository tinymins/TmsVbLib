Attribute VB_Name = "modGetWindowHandle"
Option Explicit
Public Declare Function GetForegroundWindow Lib "user32" () As Long
