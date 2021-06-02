str= "The New Order with :Ord-1234 was Successfully Displayed"

str=Replace(str,"The New Order with :","")
str=Replace(str," was Successfully Displayed","")

MsgBox str