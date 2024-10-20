CLASS zcl_xlom__ex_ut_parser_item DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PRIVATE SECTION.
    TYPES tt_item TYPE STANDARD TABLE OF REF TO zcl_xlom__ex_ut_parser_item WITH EMPTY KEY.

    DATA type       TYPE zcl_xlom__ex_ut_lexer=>ts_token-type.
    DATA value      TYPE zcl_xlom__ex_ut_lexer=>ts_token-value.
    DATA expression TYPE REF TO zif_xlom__ex.
    DATA subitems   TYPE tt_item.

    CLASS-METHODS create
      IMPORTING !type         TYPE zcl_xlom__ex_ut_lexer=>ts_token-type
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_ut_parser_item.
ENDCLASS.


CLASS zcl_xlom__ex_ut_parser_item IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_ut_parser_item( ).
    result->type = type.
  ENDMETHOD.
ENDCLASS.
