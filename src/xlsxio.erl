-module(xlsxio).

-behaviour(application).

%% Application callbacks
-export([start/2, stop/1]).
-export([start/0]).
-export([test/0]).
-export([test2/0]).
%% ===================================================================
%% Application callbacks
%% ===================================================================
-include("xlsxio.hrl").
start(_StartType, _StartArgs) ->
    xlsxio_sup:start_link().

stop(_State) ->
    ok.

start()->
    application:start(?MODULE).

test2()->
    XlsxFile = "xlsx/t2.xlsx",
    XlsxHandle = xlsx_writer:create(XlsxFile),
    SheetHandle = xlsx_writer:create_sheet("Hello",[[1,1,1,2],[1,2,2]]),
    xlsx_writer:add_sheet(XlsxHandle,SheetHandle),
    xlsx_writer:close(XlsxHandle).


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