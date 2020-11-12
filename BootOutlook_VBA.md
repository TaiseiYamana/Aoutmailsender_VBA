VBA Code
```
Option Explicit

Sub Startoutlook()
    Dim tol As Outlook.Application
    Dim tns As Namespace
    
    'Outlookのインスタンスを生成する
    Set tol = CreateObject("Outlook.Application")
    '受信トレイを開く
    Set tns = tol.GetNamespace("MAPI")
    tns.GetDefaultFolder(olFolderInbox).Display
    
    Set tns = Nothing
    Set tol = Nothing
End Sub

```
