class ZCL_XLOM__EX_UT_PARSER_ITEM definition
  public
  final
  create private
  global friends zif_xlom__ut_all_friends.

public section.
protected section.
private section.

  types:
    tt_item TYPE STANDARD TABLE OF REF TO zcl_xlom__ex_ut_parser_item WITH EMPTY KEY .

  data TYPE type ZCL_XLOM__EX_UT_LEXER=>TS_TOKEN-TYPE .
  data VALUE type ZCL_XLOM__EX_UT_LEXER=>TS_TOKEN-VALUE .
  data EXPRESSION type ref to ZIF_XLOM__EX .
  data SUBITEMS type TT_ITEM .

  class-methods CREATE
    importing
      !TYPE type ZCL_XLOM__EX_UT_LEXER=>TS_TOKEN-TYPE
    returning
      value(RESULT) type ref to ZCL_XLOM__EX_UT_PARSER_ITEM .
ENDCLASS.



CLASS ZCL_XLOM__EX_UT_PARSER_ITEM IMPLEMENTATION.


  method CREATE.

    RESULT = new ZCL_xlom__ex_ut_parser_item( ).
    result->type = type.

  endmethod.
ENDCLASS.
