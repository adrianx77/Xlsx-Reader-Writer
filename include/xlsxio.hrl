%%%-------------------------------------------------------------------
%%% @author Adrianx Lau <adrianx.lau@gmail.com>
%%% @copyright (C) 2017, 
%%% @doc
%%%
%%% @end
%%% Created : 20. 二月 2017 下午1:03
%%%-------------------------------------------------------------------

-author("Adrianx Lau <adrianx.lau@gmail.com>").
-record(xlsx_sheet,{id,name}).
-record(xlsx_share,{id,string}).

-record(xlsx_read,{xlsx_context,sheet_handler,share_putter,share_getter,row_handler}).
