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
    DictKey = XlsxFile,
    ReadOp = #xlsx_read{
        xlsx_context = DictKey,
        sheet_handler =
        fun(Context, SheetInfo) ->
            case get({Context, sheet}) of
                undefined -> put({Context, sheet}, SheetInfo);
                OldSheetInfo -> put({Context, sheet}, [SheetInfo | OldSheetInfo])
            end
        end,

        share_putter =
        fun(Context, ShareString) ->
            Table = case get({Context, share}) of
                        undefined ->
                            Tab = ets:new(share_table, [set, {keypos, #xlsx_share.id}]),
                            put({Context, share}, Tab),
                            Tab;
                        Tab -> Tab
                    end,
            io:format("ShareString:~ts~n",[ShareString#xlsx_share.string]),
            ets:insert(Table, ShareString)
        end
    },
    case xlsx_reader:read(XlsxFile,ReadOp) of
        {error,Reason}-> io:format("read error:~p~n",[Reason]);
        ok-> ok
    end.