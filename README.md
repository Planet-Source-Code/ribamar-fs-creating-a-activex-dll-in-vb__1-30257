<div align="center">

## Creating a ActiveX DLL in VB


</div>

### Description

Creating a ActiveX DLL in VB5 or VB6.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-12-30 12:49:56
**By**             |[Ribamar FS](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ribamar-fs.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Creating\_a4541312312001\.zip](https://github.com/Planet-Source-Code/ribamar-fs-creating-a-activex-dll-in-vb__1-30257/archive/master.zip)





### Source Code

<h2><b>CRIATING ACTIVEXDLL WITH VB5 OR VB6</h2>
<br>
<br>
1 &#8211; Creating a new project ActiveX DLL<br></b>
<br>
- Open a new project in VB<br>
- Select a type ActiveX DLL<br>
- Press Ctrl+S to save Class with name clsMath<br>
- Change Instancing propertie to GlobalMultiUse. <br>
- Select Project menu &#8211; Properties and change only: <br>
	- Change Project Name in General to DLLMath<br>
	- Change Project Description to My DLL Math<br>
- In File menu - Save Project (DLLMath.vbp or other) <br><br>
<b>2 &#8211; WRITING CODE TO CLASS</b><br>
<br>
Now writing functions of clsMath class. <br>
Double click in class clsMath and paste the code below: <br>
<br>
Public Function fSum(ByVal X As Long, ByVal Y As Long) As Long<br>
  fSum = X + Y<br>
End Function<br>
<br>
Public Function fSub(ByVal X As Long, Y As Long) As Long<br>
  fSub = X - Y<br>
End Function<br>
<br>
Public Function fMult(ByVal X As Long, Y As Long) As Long<br>
  fMult = X * Y<br>
End Function<br>
<br>
Public Function fDiv(ByVal X As Long, Y As Long) As Long<br>
  If Y <> 0 Then<br>
    fDiv = X / Y<br>
  Else<br>
    MsgBox " The divider must be different of zero.!" <br>
  End If<br>
End Function<br>
<br>
Press Ctrl+S to save. <br>
<br>
<b>3 &#8211; ADD A NEW CLASS</b><br>
<br>
- In Project menu &#8211; Add Class Module<br>
- Click in Class Module and in Open<br>
- Rename this Class to clsTrig<br>
- Change Instancing propertie of this Class to GlobalMultiUse. <br>
- Double click in this Class to open<br>
<br>
<b>4 &#8211; WRITING CODE OF THIS NEW CLASS</b><br>
<br>
Paste this code in clsTrig Class: <br>
<br>
Public Function fSin(X As Double) <br>
  fSin = Sin(X) <br>
End Function<br>
<br>
Public Function fCos(X As Double) <br>
  fCos = Cos(X) <br>
End Function<br>
<br>
Press Ctrl+S to save this class with name clsTrig or other<br>
<br>
<b>5 &#8211; COMPILING THE DLL</b><br>
<br>
In File menu - Make DLLMath and wait. <br>
After VB compile then register automaticaly the DLL. <br>
<br>
<b>6 &#8211; TRY DLL</b><br>
<br>
- Open a new project Standard EXE<br>
- In Project menu &#8211; References - Browse<br>
- Select DLLMath in your folder and OK. <br>
<br>
Double click in form and paste in Load event. <br>
<br>
Private Sub Form_Load()<br>
Dim objMath As DLLMath.clsMath<br>
' DLLMath here is name in Project - Properties - Project Name<br>
<br>
Set objMath = New DLLMath.clsMath<br>
<br>
MsgBox objMath.fSum(2, 6) <br>
<br>
End Sub<br>
<br>
Press F5 to run. <br>
<br>
<b>7 &#8211; HOW TO REFERENCE DLL AND CLASS</b><br>
<br>
To make reference for all classes of a DLL<br>
<br>
Dim NameVariableObject As NameOfDLL.NameOfClass<br>
<br>
To make reference for a unique class of a DLL <br>
<br>
Dim NameVariableObject As NameOfClass<br>
<br><br>
Forgive my English. I studied the English little.

