<div align="center">

## Mailto: \(98/NT/2000 compatible\)


</div>

### Description

A very simple program that opens a new message in Outlook/Outlook Express (may work on others but not tested) addressed to whatever you enter in the text box. Submitted because other examples didn't work with NT/2000. This one does
 
### More Info
 
Used the shell32 api call


<span>             |<span>
---                |---
**Submitted On**   |2000-07-30 01:05:16
**By**             |[Tim Kent](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-kent.md)
**Level**          |Beginner
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD83407292000\.zip](https://github.com/Planet-Source-Code/tim-kent-mailto-98-nt-2000-compatible__1-10194/archive/master.zip)

### API Declarations

```
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
```





