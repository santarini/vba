Select Text After Character (,)

`=RIGHT(A1,LEN(A1)-FIND(",",A1))`

Select Text Before Character

`=LEFT(A1,FIND(",",A1)-1)`


Select Test Between Characters



Select Text After Second occurence of space


`=MID(A1, FIND(" ", A1, FIND(" ", A1)+1)+1,256)`
