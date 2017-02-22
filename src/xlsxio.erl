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
    case xlsx_reader:read(XlsxFile,fun(SheetName,[Line|Row])-> io:format("~ts=====>~p | ~ts~n",[SheetName,Line,Row]) end) of
        {error,Reason}-> io:format("read error:~p~n",[Reason]);
        ok-> ok
    end.