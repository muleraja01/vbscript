Dim Message, ArrayList

' Initializes the message.
Message = ""

' Creates a list.
Set ArrayList = CreateObject("System.Collections.ArrayList")

' Fills the list.
ArrayList.Add 6
ArrayList.Add 8
ArrayList.Add 2
ArrayList.Add 4
ArrayList.Add 1
ArrayList.Add 5
ArrayList.Add 3
ArrayList.Add 3
ArrayList.Add 7

' Converts the list into an array.
Message = Message & "1: " & Join(ArrayList.ToArray, ",") & vbCrLf

' Returns the element at the specified index.
Message = Message & "2: " & ArrayList(0) & vbCrLf

' Returns the element count.
Message = Message & "3: " & ArrayList.Count & vbCrLf

' Finds out whether an element is contained in the list.
Message = Message & "4: " & ArrayList.Contains(5) & vbCrLf

' Sorts the list.
ArrayList.Sort
Message = Message & "5: " & Join(ArrayList.ToArray, ",") & vbCrLf

' Reverses the sorting direction.
ArrayList.Reverse
Message = Message & "6: " & Join(ArrayList.ToArray, ",") & vbCrLf

' Removes the element.
ArrayList.Remove 8

' Removes the element at the specified index.
ArrayList.RemoveAt 6

' Removes all elements.
ArrayList.Clear

' Releases the list.
Set ArrayList = Nothing

' Displays the result.
MsgBox Message, vbSystemModal