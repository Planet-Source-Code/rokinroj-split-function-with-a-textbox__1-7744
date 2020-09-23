<div align="center">

## split function with a textbox


</div>

### Description

Good example of the split and randomize commands.

lets you place a long list of text in a textbox and randomly grab a line. My use for this was I had about 500 cd keys to issue to customers as they would callin. I would run this script, give them a number, remove from my list later, etc.

Possible uses might be for numbers, tip of the day quotes, etc.

Rokinroj
 
### More Info
 
text1.text needs to contain a list of text that is seperated by a return


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rokinroj ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rokinroj.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rokinroj-split-function-with-a-textbox__1-7744/archive/master.zip)





### Source Code

```
Private sArray() As String
Private Sub cmdGetKey_Click()
Dim RandNum As Long
  Randomize
  RandNum = Int(Rnd * 1446) + 1
  Text1.Text = sArray(RandNum)
End Sub
Private Sub Form_Load()
   sArray() = Split(txtKeys.Text, vbCrLf)
End Sub
```

