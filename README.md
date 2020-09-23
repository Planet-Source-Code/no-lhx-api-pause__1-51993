<div align="center">

## API Pause


</div>

### Description

Demonstrates how to pause in VB, without winding up the CPU, as While/Doevents/Wend code will do.
 
### More Info
 
Milliseconds to pause

This uses the timeout feature of the WaitForSingleObject API call to break up a pause into smaller millisecond time chunks, allowing for a more responsive pause function. It also eliminates the CPU usage issue when using While/Doevents/Wend loops that make it hard to catchs problems. Since it does not do a comparison to the system clock, it mearly counts up milliseconds, there are no issues with midnight or clocks. You can change the resolution to suit your system. I have not submitted in a long time, so let me know if you would like more like this.

Noel.


<span>             |<span>
---                |---
**Submitted On**   |2004-02-25 14:16:00
**By**             |[Noëlhx](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/no-lhx.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[API\_Pause1713292252004\.zip](https://github.com/Planet-Source-Code/no-lhx-api-pause__1-51993/archive/master.zip)

### API Declarations

```
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpname As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
```





