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
    XlsxFile = "xlsx/theme_1003_new.xlsx",
    case xlsx_reader:read(XlsxFile,fun(SheetName,Row)-> io:format(" Sheet ~ts=> ~p~n",[SheetName,Row]) end) of
        {error,Reason}-> io:format("read error:~p~n",[Reason]);
        ok-> ok
    end.