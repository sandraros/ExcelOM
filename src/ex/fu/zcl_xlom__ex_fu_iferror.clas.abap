"! IFERROR(value, value_if_error)
"! IFERROR(#N/A,"1") returns "1"
"! https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611
CLASS zcl_xlom__ex_fu_iferror DEFINITION
  PUBLIC FINAL
  INHERITING FROM zcl_xlom__ex_fu
*  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_fu.

  PUBLIC SECTION.
    CLASS-METHODS create
      IMPORTING !value         TYPE REF TO zif_xlom__ex
                value_if_error TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result)  TYPE REF TO zcl_xlom__ex_fu_iferror.

    METHODs zif_xlom__ex~evaluate_single REDEFINITION.

  PROTECTED SECTION.
    METHODS constructor.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        value          TYPE i VALUE 1,
        value_if_error TYPE i VALUE 2,
      END OF c_arg.
*    DATA value          TYPE REF TO zif_xlom__ex.
*    DATA value_if_error TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_iferror IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    zif_xlom__ex~type = zif_xlom__ex=>c_type-function-iferror.
    zif_xlom__ex~parameters = VALUE #( ( name = 'VALUE         ' )
                                       ( name = 'VALUE_IF_ERROR' ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_fu_iferror( ).
    result->zif_xlom__ex~arguments_or_operands = VALUE #( ( value          )
                                                          ( value_if_error ) ).
    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = result->zif_xlom__ex~arguments_or_operands ).
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-iferror.
*    result->value             = value.
*    result->value_if_error    = value_if_error.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'VALUE         ' object = value )
*                                                       ( name = 'VALUE_IF_ERROR' object = value_if_error ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                             context   = context ).
*    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    result = COND #( LET value_result = arguments[ c_arg-VALUE ]
                                                  IN
                                              WHEN value_result->type = zif_xlom__va=>c_type-error
                                              THEN arguments[ c_arg-VALUE_IF_ERROR ]
                                              ELSE value_result ).
    zif_xlom__ex~result_of_evaluation = result.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE zcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~set_result.
*    zif_xlom__ex~result_of_evaluation = value.
*    result = value.
  ENDMETHOD.
ENDCLASS.
