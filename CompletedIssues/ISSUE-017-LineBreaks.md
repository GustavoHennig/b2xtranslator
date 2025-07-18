
# Line-break very specific scenario

Please check the file `samples\Bug44292.doc`, a specific line break is missing.  
99% of the line breaks in other files work; only special ones like this are not working.


The expected result is:
```txt
	One paragraph is ok	First para is ok
Second paragraph is skipped	One paragraph is ok
```

But the current result is:
```txt
	One paragraph is ok	First para is okSecond paragraph is skipped	One paragraph is ok
```


## To run:

`dotnet run --project Shell/doc2text/doc2text.csproj -- samples/Bug44292.doc Bug44292.txt`
