-module(xlsxio).

-behaviour(application).

%% Application callbacks
-export([start/2, stop/1]).
-export([start/0]).
-export([test/0]).
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