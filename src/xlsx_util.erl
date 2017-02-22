%%%-------------------------------------------------------------------
%%% @author Adrianx Lau <adrianx.lau@gmail.com>
%%% @copyright (C) 2017, 
%%% @doc
%%%
%%% @end
%%% Created : 22. 二月 2017 上午10:38
%%%-------------------------------------------------------------------
-module(xlsx_util).
-author("Adrianx Lau <adrianx.lau@gmail.com>").
-include_lib("xmerl/include/xmerl.hrl").
%% API
-compile(export_all).

sprintf(Format,Args)->
	lists:flatten(io_lib:format(Format,Args)).

is_record(Term, RecordTag) when is_tuple(Term)  ->
	element(1, Term) =:= RecordTag;
is_record(_Term, _RecordTag)-> false.



has_attribute_value( AttrName, AttValue,XmlNode) ->
	lists:any(fun(Attr) -> (Attr#xmlAttribute.name =:= AttrName) and
		(Attr#xmlAttribute.value =:= AttValue)
			  end, XmlNode#xmlElement.attributes).

xmlattribute_value_from_name(Name, XmlNode) ->
	case lists:keyfind(Name,#xmlAttribute.name,XmlNode#xmlElement.attributes) of
		false-> error(xlsx_util:sprintf("xmlattribute_value_from_name exception: ~p ~p",[Name,XmlNode]));
		Attribute-> Attribute#xmlAttribute.value
	end.

xmlElement_from_name(Name,XmlNode)->
	case lists:keyfind(Name,#xmlElement.name,XmlNode) of
		false-> error(xlsx_util:sprintf("xmlElement_from_name exception: ~p ~p",[Name,XmlNode]));
		Element-> Element
	end.



get_column_count(DimString) ->
	[BeginStr, EndStr] = case string:tokens(DimString, ":") of
							 [BeginString, EndString] -> [BeginString, EndString];
							 [BeginString] -> [BeginString, BeginString]
						 end,
	[Begin] = string:tokens(BeginStr, "0123456789"),
	[End] = string:tokens(EndStr, "0123456789"),
	field_to_num(End) - field_to_num(Begin) + 1.

integer_pow(N, L) ->
	if L >= 1 ->
		N * integer_pow(N, L - 1);
		true ->
			1
	end.
field_to_num(NumStr) ->
	{_, Num} = lists:foldr(
		fun(C, {I, N}) ->
			New = (C - $A + 1) * integer_pow($Z - $A + 1, I) + N,
			{I + 1, New}
		end, {0, 0}, NumStr),
	Num.
get_field_number(FieldString)->
	[FieldName] = string:tokens(FieldString, "0123456789"),
	PostNumString = string:right(FieldString, length(FieldString) - length(FieldName)),
	{field_to_num(FieldName), list_to_integer(PostNumString)}.

take_nth_list(N, List, Value)->
	{NewList, _Acc} = lists:mapfoldl(
		fun(I, Index)->
			if Index =:= N->
				{Value, Index + 1};
				true->
					{I, Index + 1}
			end
		end, 1, List),
	NewList.