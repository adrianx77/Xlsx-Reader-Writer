# Xlsx-Reader-Writer
Reader &amp; Writer for Xlsx in Erlang

## USAGE:
```
shell>git clone git@github.com:adrianx77/Xlsx-Reader-Writer.git
shell>./rebar co
shell>erl xlsxio -pa ebin -s xlsxio test
```


```
read(XlsxFile, RowHandler)

    XlsxFile  : only support higher than excel 2003 ,could be .xlsx or .xlsxm
    RowHandler: callback function =>fun(SheetName,RowData)
        SheetName:Sheet Name
        RowData  : true text list
        return   : next_sheet | next_row | break
            next_sheet  : ignor left rows ,continue process next sheet
            next_row    : continue next row
            break       : ignor next rows and sheets
```
so you can write such code:
```
test()->
    XlsxFile = "xlsx/t.xlsx",
    RowHandler =
        fun(SheetName, [1 | Row]) ->
            io:format("~n~ts=====>~p | ", [SheetName, 1]),
            lists:foreach(fun(R)-> io:format("\t~ts",[R]) end,Row),
            next_row;
            (SheetName, [Line | Row]) ->
                io:format("~n~ts=====>~p | ", [SheetName, Line]),
                lists:foreach(fun(R)-> io:format("\t~ts",[R]) end,Row),
            next_sheet
        end,

    case xlsx_reader:read(XlsxFile,RowHandler) of
        {error,Reason}-> io:format("read error:~p~n",[Reason]);
        ok-> io:format("~n")
    end.
```

## Future
I will Add support for writer.