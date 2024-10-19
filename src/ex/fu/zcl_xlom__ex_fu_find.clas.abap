"! FIND(find_text, within_text, [start_num])
"! https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628
CLASS zcl_xlom__ex_fu_find DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING find_text     TYPE REF TO zif_xlom__ex
                within_text   TYPE REF TO zif_xlom__ex
                start_num     TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_find.

  PRIVATE SECTION.
    DATA find_text   TYPE REF TO zif_xlom__ex.
    DATA within_text TYPE REF TO zif_xlom__ex.
    DATA start_num   TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_find IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_find( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-find.
    result->find_text         = find_text.
    result->within_text       = within_text.
    result->start_num         = start_num.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = me
                                 context    = context
                                 operands   = VALUE #( ( name = 'FIND_TEXT  ' object = find_text   )
                                                       ( name = 'WITHIN_TEXT' object = within_text )
                                                       ( name = 'START_NUM  ' object = start_num   ) ) ).
    result = zif_xlom__ex~set_result(
                 COND #( WHEN array_evaluation-result IS BOUND
                         THEN array_evaluation-result
                         ELSE zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                                            context   = context ) ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(result_of_find_text) = zcl_xlom__va=>to_string( arguments[ name = 'FIND_TEXT' ]-object )->get_string( ).
        DATA(result_of_within_text) = zcl_xlom__va=>to_string( arguments[
                                                                   name = 'WITHIN_TEXT' ]-object )->get_string( ).
        DATA(result_of_start_num) = CAST zcl_xlom__va_number( arguments[ name = 'START_NUM' ]-object ).
        DATA(start_offset) = COND i( WHEN result_of_start_num IS BOUND THEN result_of_start_num->get_number( ) ).
        IF start_offset > strlen( result_of_within_text ).
          result = zcl_xlom__va_error=>value_cannot_be_calculated.
        ELSE.
          DATA(result_offset) = COND #( WHEN result_of_find_text IS INITIAL
                                        THEN 1
                                        ELSE find( val = result_of_within_text
                                                   sub = result_of_find_text
                                                   off = start_offset ) + 1 ).
          IF result_offset = 0.
            result = zcl_xlom__va_error=>value_cannot_be_calculated.
          ELSE.
            result = zcl_xlom__va_number=>create( CONV #( result_offset ) ).
          ENDIF.
        ENDIF.
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
