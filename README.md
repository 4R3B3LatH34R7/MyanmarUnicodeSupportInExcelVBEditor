# Myanmar Unicode in Excel VBEditor
## A Brief Review of Approaches to Visualize Unicode in VBIDE in Microsoft Excel
This is going to be just discussing ways to show Unicode characters in Excel's VBEditor.\
The premises I shall base all this upon, are, <b>MS Excel 2010, Myanmar Unicode font: Pyidaungsu</b> only.\
For any other situation, I have no guarantee that this discussion could be applicable. So, YMMV.

### I. Cause of Problem
Why do I even have to discuss this?\
There are many Myanmar people who find this situation frustrating, including me, when we tried to write Myanmar words in Myanmar Unicode fonts like, Pyidaungsu.\
While it is generally easy to find out the cause, most people who started writing VBA code doe not even have the necessary googling skills just find out nor the basic English language skills, nor the coding knowledge to understand that, the VBE(VBIDE/Visual Basic Editor) only works with ANSI codepage, thus, it cannot show Unicode characters. Period. Microsoft will not fix this, as of right now, 11DEC2023 09:30AM MMR STD time.\
And all that we type in VBE in Myanmar Unicode font will show up like "???".

### II. How do we approach this problem?
#### II.1.Referencing a value on another worksheet
This is the simplest method.\
I was a n00b once but even back then, if I wanted to check whether a variable is equal to item(s) written in Myanmar Unicode font,
```vba
If sProdName = "ရှမ်းခေါက်ဆွဲ" Then
```
I realized that we can just employ a simple method like below.\
I just need to put the word ရှမ်းခေါက်ဆွဲ in Cell A1 of worksheet Sheet1 and reference from the code like:
```vba
If sProdName = Sheet1.Range("A1").Value then
```
While there's nothing wrong with this approach, there are a few pros and cons about this.
|No.|pros|cons|
|---|---|---|
|1|Easy to accomplish|FileSize bigger because of extra sheet|
|2|End-User could update values per their requirements|End-User could mess up|
|3|Could be a feature|Remedy above with changing worksheet visibility|
|4|Does not affect code size (lines count)|Can get corrupted|
|5|No need to (write functions to) convert values|Savvy users can find out values even inside veryHidden sheets|
|6|N00B friendly|Too simple|

#### II.2.Convert Myanmar Unicode String to Unicode number and convert it back to Unicode String at run-time
This may seem like hard for a n00b but the concept is still simple.\
Basically, we just need some Sub/Function to convet a list of Myanmar Unicode text values on a worksheet to get their respective Unicode character codes printed out to Immediate window.\
This is a one-time process. However, we could repeat this as much as we need/want.\
The proposed function could be as simple as what's outlined below:
```vba
'to collect unicode values to be used as arrays or string in subs/functions
Sub convertMMRtoUnicodeArray(Optional theSeparator As String = "|") 
  Dim oneCell, i As Integer, st As String
  For Each oneCell In Sheet1.Range("C2:C4") 'change range as required
    st = ""
    For i = 1 To Len(oneCell)
      st = st & IIf(st = "", "", IIf(theSeparator <> "|", theSeparator, "|")) & AscW(Mid(oneCell, i, 1))
    Next i
    Debug.Print oneCell.Address(False, False) & " = " & st
  Next oneCell
End Sub
```
Above function could be easily shortened to become a one-liner we can run inside the Immediate window, without needing writing a function in CodePane inside of VBE:
```vba
for each oneCell in range("C2:C4"):st="":for i=1 to len(oneCell):st=st &iif(st="","","|") &ascw(mid(oneCell,i,1)) :next i:?st:next oneCell
```
The output of above function can be observed below:
![output_convertMMRtoUnicodeNumber](images/convertingMMRtoUnicodeNumber.png)
