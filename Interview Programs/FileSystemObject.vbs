FSO

'Build Path :Appends a name to an existing path  FSO.BuildPath("","")
'CopyFile 
'GetAbsolutePathName :Returns the complete path from the root of the drive for the specified path
'					  Syntax FSO.GetAbsolutePathName(path)

'GetFile:Returns the file name or folder name for the last component in a specified path
			'FSO.GetFileName(path)
'GetFileName:


'OpenTextFile :Opens a file and returns a TextStream object that can be used to access the file
'				FSO.OpenTextFile(fname,mode,create,format)
'				mode:	Optional. How to open the file
'				1=ForReading - Open a file for reading. You cannot write to this file.
'				2=ForWriting - Open a file for writing.
'				8=ForAppending - Open a file and write to the end of the file.'
'AtEndOfLine  : Returns true if the file pointer is positioned immediately before the end-of-line marker in a TextStream file, and false if not
'AtEndOfStream:	Returns true if the file pointer is at the end of a TextStream file, and false if not
'Column:	  : Returns the column number of the current character position in an input stream
'line		  : Returns the current line number in a TextStream file

'Skip: 		  : Skips a specified number of characters when reading a TextStream file
'SkipLine     :	Writes a specified text and a new-line character to a TextStream file
'WriteBlankLines :Writes a specified number of new-line character to a TextStream file