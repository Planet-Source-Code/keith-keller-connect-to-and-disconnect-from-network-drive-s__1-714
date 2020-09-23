<div align="center">

## Connect To and Disconnect From Network Drive\(s\)


</div>

### Description

Being an NT network administrator and software engineer sure has its advantages.

Visual Basic 4.0 has afforded me the opportunity to create useful apps that

greatly reduce the amount of time it takes to perform those tasks that many of us

perform often. This little app simply uses the Windows 32 API (Win95 or NT 4.0 only)

to open the network resource browse list. You can map network resources or disconnect

from network resources.

Enjoy the code! We've been using it for months in several VB apps on our network

and it works GREAT!
 
### More Info
 
Some knowledge of the Windows API would help.

Opens the respective (Connect To) dialog box or (Disconnect From) dialog box!

not aware of any


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Keith Keller](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/keith-keller.md)
**Level**          |Unknown
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/keith-keller-connect-to-and-disconnect-from-network-drive-s__1-714/archive/master.zip)

### API Declarations

```
Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Public Const RESOURCETYPE_DISK = &H1, RESOURCETYPE_PRINT = 0
```


### Source Code

```
Create a Form with 4 command buttons.
Name the first three buttons: 'Command1' (This will create a Control Array)
Label the first button: 'Connect Drive'
Label the second button: 'Disconnect Drive'
Label the third button: 'End Capture'
Label the fourth button: 'Quit'
Double-Click on one the button labelled "Connect Drive" and enter the following:
Private Sub Command1_Click(Index As Integer) <<== You won't need this line
  Dim x As Long
  If Index = 0 Then  'Connect
    x = WNetConnectionDialog(Me.hwnd, RESOURCETYPE_DISK)
  ElseIf Index = 1 Then 'Disconnect
    x = WNetDisconnectDialog(Me.hwnd, RESOURCETYPE_DISK)
  Else
    End
  End If
End Sub <<== You won't need this line either.
Name the fourth button 'printerbutton'. Double-Click it and enter the following:
Private Sub printerbutton_Click()
  Dim x As Long
  x = WNetDisconnectDialog(Me.hwnd, RESOURCETYPE_PRINT)
End Sub
Run the app and click each of the buttons to see what happens!
Hope you find it useful!
If you're interested in trading VB code tips, email me at: kkeller@1stnet.com
```

