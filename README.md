# ExcelVBA_insert-blank-line
When you use microsoft forms,This code can insert a row in Excel when a consecutive number is missing.
# How to
1. you can use developer tab in Excel
2. include bas file your excel file(import from microsoft forms)
3. check as follow list and correct [form.bas]
  - if you use class number(1,2,3・・・),you change A,B,C,D・・・ to 1,2,3,4,・・・
  - if class number is not column I,you change [Select Case gakuen(i,9)] -> [Select Case gakuen(i,"your class column number")]
  - if attendance number is not column L,you change [If Cells(i,12).Value] -> [If Cells(i,"your attendance column number").Value](all change class A ~ classIB)
# Attention
If your student missed input attendance number and same attendance number is exsisted, once run delite.bas and modified attendance number.  
