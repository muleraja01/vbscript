
The XML DOM (XML Document Object Model) defines methods for accessing and manipulating XML documents.This topic will discuss how to manipulate  XML  using  vbscript and accessing information from XML.

Methods to Access XML DOM:

1. Creating an object for xmlDOM
set  xmldom = Createobject("MSXML.DOMDocument")
xmldom.async = “False”

2. Loading a file using xml DOM

Xmldom.Load(FileName)

3. To Find number of parent nodes in a xml
Intlen = Xmldom.childnodes.length

4. To find names of parent nodes
For i=0 to intlen-1
      strParNodeName = xmldom.childnodes(i).Name
Next

5. How to search a node with tagName
Set nodelist = xmldom.getElementsByTagName(strTagName)

This gives a collection of nodes with tag name as strTagName.
The tagName value "*" will returns all elements in the document.

6. How to extract information from nodelist using GetElementbyTagName
Set nodelist = xmldom.getElementsByTagName(strTagName)
For i= 0 to nodelist.length – 1
       strText = nodelist.item(i).xml;
       msgbox strText
Next

So points to remember in this are as follows:

We can get all elements with the tagname using xmldom.getElementsByTagName
Once we have the list, we can access each of the nodes in the list using item(index).property
To access attribute property value for a node , use strText = nodelist.item(i).getattribute("category")
To get xml content of a node, use strText = nodelist.item(i).xml
Use strText = nodelist.item(i).nodeName to get the name of required node.
Use strText = nodelist.item(i).text to get value in a particular node

Sample Code explaining the concept

FileName = "d:\xmldom.xml"
''Create an instance of xml dom
set  xmldom = Createobject("MSXML.DOMDocument")
''set async = "False"
xmldom.async = "False"
xmldom.Load(FileName)
intlen = xmldom.childnodes.length
msgbox intlen
Set nodelist = xmldom.getElementsByTagName("title")
msgbox nodelist.length
If nodelist.length > 0 then
    For each x in nodelist
        AttName=x.getattribute("lang")
        myname = x.Text
        msgbox AttName +" "+ myname
    Next
Else
    msgbox " field not found."
End If

7. Using Selectnodes method to extract information

FileName = "d:\xmldom.xml"
''Create an instance of xml dom
set  xmldom = Createobject("MSXML.DOMDocument")
''set async = "False"
xmldom.async = "False"
xmldom.Load(FileName)
intlen = xmldom.childnodes.length
msgbox intlen
Set nodelist = xmldom.selectNodes("/bookstore/book/title")
msgbox nodelist.length
If nodelist.length > 0 then
    For each x in nodelist
        AttName=x.getattribute("lang")
        myname = x.Text
        msgbox AttName +" "+ myname
    Next
Else
    msgbox " field not found."
End If

8. Using SelectSingleNode to extract information

FileName = "d:\xmldom.xml"
''Create an instance of xml dom
set  xmldom = Createobject("MSXML.DOMDocument")
''set async = "False"
xmldom.async = "False"
xmldom.Load(FileName)
''This will give 1st item property matching the node structure
Msgbox xmldom.selectSingleNode("/bookstore/book/title").Text
''This will give 1st item property matching the node structure with filter as attribute value
Msgbox xmldom.selectSingleNode("/bookstore/book/title[@lang = 'ke']").Text
''This will give 1st item property matching the node structure with filter as text  value
Msgbox xmldom.selectSingleNode("/bookstore/book/title[.= 'Learning XML']").Text