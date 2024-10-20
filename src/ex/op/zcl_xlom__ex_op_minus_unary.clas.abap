"! negative numbers like in OFFSET(B2,-1)
CLASS zcl_xlom__ex_op_minus_unary DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_op.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING operand       TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_op_minus_unary.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        operand TYPE i VALUE 1,
      END OF c_arg.

    METHODS constructor.
*    DATA operand TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_op_minus_unary IMPLEMENTATION.
  METHOD constructor.
    zif_xlom__ex~type = zif_xlom__ex=>c_type-operation-minus_unary.
    zif_xlom__ex~parameters = VALUE #( ( name = 'OPERAND' ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_op_minus_unary( ).
*    result->operand           = operand.
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-operation-minus_unary.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'OPERAND' object = operand ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                             context   = context ).
*    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(operand_result) = arguments[ c_arg-OPERAND ].
        result = zcl_xlom__va_number=>create( -
                                              ( zcl_xlom__va=>to_number( operand_result )->get_number( ) ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
    zif_xlom__ex~result_of_evaluation = result.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    IF     expression       IS BOUND
*       AND expression->type  = zif_xlom__ex=>c_type-operation-minus_unary
*       AND operand->is_equal( CAST zcl_xlom__ex_op_minus_unary( expression )->operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
