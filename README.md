# Xlsx-Erlang
Reader &amp; writer for Xlsx in Erlang 


#read(XlsxFile, RowHandler)

##XlsxFile : only support higher than excel 2003 ,could be .xlsx or .xlsxm
##RowHandler: callback function
###fun(SheetName,RowData)
###SheetName:Sheet Name
###RowData ： true text list

so you can write such code:
```
test()->
    XlsxFile = "xlsx/theme_1003_new.xlsx",
    case xlsx_reader:read(XlsxFile,fun(SheetName,[Line|Row])-> io:format("~ts=====>~p | ~ts~n",[SheetName,Line,Row]) end) of
        {error,Reason}-> io:format("read error:~p~n",[Reason]);
        ok-> ok
    end.
```