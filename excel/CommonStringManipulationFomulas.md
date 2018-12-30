<b>Extract 5 Characters from Beginning of Cell</b>
  
  `=LEFT(A1,5)`
  
<b>Extract 5 Characters from End of Cell</b>
  
  `=Right(A1,5)`
  
<b>Extract 10 charactes after first 5 characters</b>

  `=MID(A1,5,10)`
  
<b>Extract Text from Left After First Occurence of Character</b>

  `=RIGHT(A1,LEN(A1)-FIND("Character",A1))`

<b>Extract Text from Left Before First Occurence of Character</b>

  `=LEFT(A1,FIND("Character",A1)-1)`
  
<b>Extract Text from Right Before First Occurence of Character</b>
  
  `=TRIM(RIGHT(SUBSTITUTE(D2,"Character",REPT(" ",255)),255))`

<b>Extract Test Between Two Characters</b>

  `=MID(A1,SEARCH("FirstChar",A1)+3,SEARCH("LastChar",A1)-SEARCH("FirstChar",A1)-4)`

<b>Extract text after first occurence of second character, after first occurence of first character</b>

  `=MID(A1, FIND("SecondChar", A1, FIND("FirstChar", A1)+1)+1,256)`
 
