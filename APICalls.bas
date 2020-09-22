Attribute VB_Name = "APICalls"
Public Type INITCOMMONCONTROLSEX_TYPE
    dwSize As Long
    dwICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As _
    INITCOMMONCONTROLSEX_TYPE) As Long
Public Const ICC_INTERNET_CLASSES = &H800
