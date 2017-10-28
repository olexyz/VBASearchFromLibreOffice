sub Search
dim lancher As object
dim aWebPage as String
dim searchText as String
dim selection as object
luncher= createUnoService("com.sun.star.system.SystemShellExecute")
selection=ThisComponent.CurrentSelection
if selection.count = 1 then
 searchText=selection.getByIndex(0).String
 aWebPage="https://www.google.ru/search?q="+searchText
else
 for i=1 to  selection.count - 1
 searchText=searchText+selection.getByIndex(i).String+chr(13)
 next
 aWebPage="https://www.google.ru/search?q="+searchText
endif
luncher.execute(aWebPage, "", 0)
end sub
