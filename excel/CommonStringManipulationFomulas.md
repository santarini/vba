<b>Select Text After Character (,)</b>

  `=RIGHT(A1,LEN(A1)-FIND(",",A1))`

<b>Select Text Before Character</b>

  `=LEFT(A1,FIND(",",A1)-1)`


<b>Select Test Between Characters</b>



<b>Select Text After Second occurence of space</b>


  `=MID(A1, FIND(" ", A1, FIND(" ", A1)+1)+1,256)`
