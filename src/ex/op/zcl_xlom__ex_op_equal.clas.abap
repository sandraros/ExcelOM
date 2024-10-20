CLASS zcl_xlom__ex_op_equal DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_op.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING left_operand  TYPE REF TO zif_xlom__ex
                right_operand TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_op_equal.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        left_operand  TYPE i VALUE 1,
        right_operand TYPE i VALUE 2,
      END OF c_arg.

    METHODS constructor.
*    DATA left_operand  TYPE REF TO zif_xlom__ex.
*    DATA right_operand TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_op_equal IMPLEMENTATION.
  METHOD constructor.
    zif_xlom__ex~type = zif_xlom__ex=>c_type-operation-equal.
    zif_xlom__ex~parameters = VALUE #( ( name = 'LEFT_OPERAND'  )
                                       ( name = 'RIGHT_OPERAND' ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_op_equal( ).
*    result->left_operand      = left_operand.
*    result->right_operand     = right_operand.
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-operation-equal.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = me
                                 context    = context
                                 operands   = VALUE #( ( name = 'LEFT'  object = left_operand )
                                                       ( name = 'RIGHT' object = right_operand ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    DATA(left_result) = arguments[ c_arg-left_operand ].
    DATA(right_result) = arguments[ c_arg-right_operand ].
    IF left_result->type <> right_result->type.
      result = zcl_xlom__va_boolean=>false.
    ELSE.
      DATA(ref_to_left_operand_value) = left_result->get_value( ).
      DATA(ref_to_right_operand_value) = right_result->get_value( ).
      ASSIGN ref_to_left_operand_value->* TO FIELD-SYMBOL(<left_operand_value>).
      ASSIGN ref_to_right_operand_value->* TO FIELD-SYMBOL(<right_operand_value>).
      result = zcl_xlom__va_boolean=>get( xsdbool( <left_operand_value> = <right_operand_value> ) ).
    ENDIF.
    zif_xlom__ex~result_of_evaluation = result.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    IF     expression       IS BOUND
*       AND expression->type  = zif_xlom__ex=>c_type-operation-equal
*       AND left_operand->is_equal( CAST zcl_xlom__ex_op_equal( expression )->left_operand )
*       AND right_operand->is_equal( CAST zcl_xlom__ex_op_equal( expression )->right_operand ).
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
