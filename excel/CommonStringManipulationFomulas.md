<b>Extract Text After Character</b>

  `=RIGHT(A1,LEN(A1)-FIND("Character",A1))`

<b>Extract Text Before Character</b>

  `=LEFT(A1,FIND("Character",A1)-1)`

<b>Extract Test Between Characters</b>

  `=MID(A1,SEARCH("FirstChar",A1)+3,SEARCH("LastChar",A1)-SEARCH("FirstChar",A1)-4)`

<b>Extracts text after first occurence of second character, after first occurence of first character</b>

  `=MID(A1, FIND("SecondChar", A1, FIND("FirstChar", A1)+1)+1,256)`
