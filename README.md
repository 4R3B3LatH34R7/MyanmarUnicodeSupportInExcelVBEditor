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
I was a n00b once but even back then, if I wanted to check whether a variable is equal to items written in Myanmar Unicode font:\
```vba
If sProdName = "ရှမ်းခေါက်ဆွဲ" Then
```
I realized that we can just employ a simple method like below.\
I just need to put the word ရှမ်းခေါက်ဆွဲ in Cell A1 of worksheet Sheet1 and reference from the code like:
```vba
If sProdName = Sheet1.Range("A1") then
```
While there's nothing wrong with this approach, there are a few pros and cons about this.\
|No.|pros|cons|
|---|---|---|
|1|End-User could update values per their requirements|End-User could mess up|
|2|Could be a feature|Remedy above with changing worksheet visibility|

