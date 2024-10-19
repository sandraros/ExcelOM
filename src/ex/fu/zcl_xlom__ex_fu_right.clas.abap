"! RIGHT(text,[num_chars])
"! A1=RIGHT({"hello","world"},{2,3}) -> A1="lo", B1="rld"
"! https://support.microsoft.com/en-us/office/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f
CLASS zcl_xlom__ex_fu_right DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !text         TYPE REF TO zif_xlom__ex
                num_chars     TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_right.

  PRIVATE SECTION.
    DATA text      TYPE REF TO zif_xlom__ex.
    DATA num_chars TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_right IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_right( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-right.
    result->text              = text.
    result->num_chars         = num_chars.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = me
                                 context    = context
                                 operands   = VALUE #( ( name = 'TEXT'      object = text )
                                                       ( name = 'NUM_CHARS' object = num_chars ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    DATA right       TYPE string.
    DATA temp_result TYPE REF TO zif_xlom__va.

    TRY.
        DATA(text) = zcl_xlom__va=>to_string( arguments[ name = 'TEXT' ]-object )->get_string( ).
        DATA(result_num_chars) = arguments[ name = 'NUM_CHARS' ]-object.
        DATA(number_num_chars) = COND #( WHEN result_num_chars       IS BOUND
                                          AND result_num_chars->type <> result_num_chars->c_type-empty
                                         THEN zcl_xlom__va=>to_number( result_num_chars )->get_number( ) ).
        IF number_num_chars < 0.
          temp_result = zcl_xlom__va_error=>value_cannot_be_calculated.
        ELSE.
          IF text = ''.
            right = ``.
          ELSE.
            DATA(off) = COND i( " Get the last character
                                WHEN result_num_chars IS NOT BOUND     THEN strlen( text ) - 1
                                " Get the whole text
                                WHEN number_num_chars > strlen( text ) THEN 0
                                " Get exactly as many characters as defined in NUM_CHARS
                                " (note that if NUM_CHARS = STRLEN( text ), the result is the empty string "")
                                ELSE                                        strlen( text ) - number_num_chars ).
            right = substring( val = text
                               off = off ).
          ENDIF.
          temp_result = zcl_xlom__va_string=>create( right ).
        ENDIF.
        result = zif_xlom__ex~set_result( temp_result ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    IF expression->type <> zif_xlom__ex=>c_type-function-right.
      RETURN.
    ENDIF.
    DATA(compare_right) = CAST zcl_xlom__ex_fu_right( expression ).

    result = xsdbool(     text->is_equal( compare_right->text )
                      AND zcl_xlom__ex_ut=>are_equal( expression_1 = num_chars
                                                      expression_2 = compare_right->num_chars ) ).
*                      AND (    (     num_chars                IS NOT BOUND
*                                 AND  IS NOT BOUND )
*                            OR (     num_chars                IS BOUND
*                                 AND compare_right->num_chars IS BOUND
*                                 AND num_chars->is_equal( compare_right->num_chars ) ) ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
