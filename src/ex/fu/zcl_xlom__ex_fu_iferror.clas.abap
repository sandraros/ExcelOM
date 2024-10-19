"! IFERROR(value, value_if_error)
"! IFERROR(#N/A,"1") returns "1"
"! https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611
CLASS zcl_xlom__ex_fu_iferror DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !value         TYPE REF TO zif_xlom__ex
                value_if_error TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result)  TYPE REF TO zcl_xlom__ex_fu_iferror.

  PRIVATE SECTION.
    DATA value          TYPE REF TO zif_xlom__ex.
    DATA value_if_error TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_iferror IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_iferror( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-iferror.
    result->value             = value.
    result->value_if_error    = value_if_error.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = me
                                 context    = context
                                 operands   = VALUE #( ( name = 'VALUE         ' object = value )
                                                       ( name = 'VALUE_IF_ERROR' object = value_if_error ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    result = zif_xlom__ex~set_result( COND #( LET value_result = arguments[ name = 'VALUE' ]-object
                                                  IN
                                              WHEN value_result->type = zif_xlom__va=>c_type-error
                                              THEN arguments[ name = 'VALUE_IF_ERROR' ]-object
                                              ELSE value_result ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
