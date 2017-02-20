%%%-------------------------------------------------------------------
%%% @author Adrianx Lau <adrianx.lau@gmail.com>
%%% @copyright (C) 2017, 
%%% @doc
%%%
%%% @end
%%% Created : 20. 二月 2017 下午12:14
%%%-------------------------------------------------------------------
-module(xlsx_reader).
-author("Adrianx Lau <adrianx.lau@gmail.com>").
-include_lib("xmerl/include/xmerl.hrl").
-include("xlsxio.hrl").

%% API
-export([read/2]).

-spec(read(XlsxFile :: string(),ReadOps :: [tuple()]) ->
	{error, Resaon :: atom()} | {error, Resaon :: string()} |
	{ok}).
read(XlsxFile,ReadOps)->
	case zip:zip_open(XlsxFile, [memory]) of
		{error,Reason}-> {error,Reason};
		{ok,ZipHandle}->
			read_memory(ZipHandle,ReadOps),
			zip:zip_close(ZipHandle)
	end.

read_memory(XlsxZipHandle,ReadOps)->
	case process_sheetinfos(XlsxZipHandle,ReadOps) of
		ok->
			case process_sharestring(XlsxZipHandle,ReadOps) of
				ok-> ok;
				{error,Reason}->{error,Reason}
			end;
		{error,Reason}->{error,Reason}
	end.

process_sharestring(ZipHandle,ReadOps)->
	ShareStringFile = "xl/sharedStrings.xml",
	case zip:zip_get(ShareStringFile, ZipHandle) of
		{error, Reason}->
			{error,lists:flatten(io_lib:format("zip:zip_get ~p error:~p", [ShareStringFile, Reason]))};
		{ok, ShareStrings}->
			{_File, Binary} = ShareStrings,
			SharePutter = ReadOps#xlsx_read.share_putter,
			XlsxContext = ReadOps#xlsx_read.xlsx_context,
			do_put_shareString(SharePutter,XlsxContext, Binary),
			ok
	end.

do_put_shareString(SharePutter, XlsxContext, BinaryString) ->
	case xmerl_scan:string(binary_to_list(BinaryString)) of
		{ParsedDocRootEl, _Rest} ->
			XMLS = ParsedDocRootEl#xmlElement.content,
			FltNodes = lists:filter(
				fun(Node) ->
					element(#xmlElement.name, Node) =:= 'si'
				end, XMLS),
			TextList = lists:map(
				fun(Node) ->
					Elements = Node#xmlElement.content,
					lists:foldl(
						fun(Element, AccTxt) ->
							case is_record_ex(Element, xmlElement) of
								false -> AccTxt;
								true ->
									case Element#xmlElement.name of
										t -> get_element_text(Element);
										r ->
											TextNodes = Element#xmlElement.content,
											[FltTextNode | _] = lists:filter(
												fun(TextNode) ->
													IsRecord = is_record_ex(TextNode, xmlElement),
													if not IsRecord -> false;
														true ->
															TextNode#xmlElement.name =:= 't'
													end
												end, TextNodes),
											AccTxt ++ get_element_text(FltTextNode);
										_ ->
											AccTxt
									end
							end
						end, [], Elements)
				end, FltNodes),
			lists:foldl(
				fun(Txt, Id) ->
					SharePutter(XlsxContext, #xlsx_share{id=Id, string = Txt}),
					Id + 1
				end, 0, TextList),
			ok; %% return ok
		ExceptResult -> {error, lists:flatten(io_lib:format("xmerl_scan:string error:~p", ExceptResult))}
	end.


is_record_ex(Term, RecordTag)->
	IsTerm = erlang:is_tuple(Term),
	if (not IsTerm)
		->false;
		true->
			element(1, Term) =:= RecordTag
	end.


get_element_text(Element)->
	IsRecord = is_record_ex(Element, xmlElement),
	if not IsRecord->[];
		true->
			if element(#xmlElement.name, Element) =:= 't'->

				TextNodes = lists:filter(
					fun(Text) ->
						is_record_ex(Text, xmlText)
					end,
					Element#xmlElement.content),

				case TextNodes of
					[]->[];
					_->lists:flatmap(
						fun(T) ->
							Text1 = element(#xmlText.value, T),
							Text1
						end,
						TextNodes)
				end
			end
	end.
process_sheetinfos(ZipHandle,ReadOps)->
	SheetFile = "xl/workbook.xml",
	case zip:zip_get(SheetFile, ZipHandle) of
		{error, Reason}->
			{error,Reason};
		{ok, WorkBook}->
			{_File, Binary} = WorkBook,
			case xmerl_scan:string(binary_to_list(Binary)) of
				{ParsedDocRootEl, _Rest}->
					SheetInfos = xmerl_xpath:string("//sheet", ParsedDocRootEl),
					SheetPutter = ReadOps#xlsx_read.sheet_handler,
					XlsxContext = ReadOps#xlsx_read.xlsx_context,
					lists:foreach(
						fun(SheetInfoRaw)->
							SheetInfo = sheetinfo_from_workbook(SheetInfoRaw),
							SheetPutter(XlsxContext,SheetInfo)
						end, SheetInfos),
					ok;
				ExceptResult->{error,lists:flatten(io_lib:format("xmerl_scan:string error:~p", ExceptResult))}
			end
	end.


sheetinfo_from_workbook(SheetInfo)->
	Attributes = SheetInfo#xmlElement.attributes,
	lists:foldl(
		fun(Attr, Acc)->
			if Attr#xmlAttribute.name =:= 'r:id'->
				Acc#xlsx_sheet{id = Attr#xmlAttribute.value};
				Attr#xmlAttribute.name =:= 'name'->
					Acc#xlsx_sheet{name = Attr#xmlAttribute.value};
				true->
					Acc
			end
		end, #xlsx_sheet{}, Attributes).

