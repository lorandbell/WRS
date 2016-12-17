toTransFileName = InputBox( "Enter the file you wish to translate:",,"text.txt" )
'"Life_and_Ministry.htm" is an example
dictFile = InputBox( "Enter your dictionary file:",, "dictionary.txt")

Set fso=CreateObject("Scripting.FileSystemObject")
dictFile = fso.OpenTextFile(dictFile).ReadAll
listLines = Split(dictFile, vbCrLf)

'File to translate
toTransData = fso.OpenTextFile(toTransFileName).ReadAll

'Output file
Set outFile = fso.CreateTextFile("Translated " + toTransFileName , True)

For Each line In listLines
   data = Split(line, ";")
   'WScript.Echo "from " + data(0) + " to " + data(1) 
   toTransData = Replace(toTransData, data(0), data(1))
Next

outFile.Write toTransData

WScript.Echo "Document translated"
