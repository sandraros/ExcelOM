CLASS zcl_xlom__ex_op_ampersand DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING left_operand  TYPE REF TO zif_xlom__ex
                right_operand TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_op_ampersand.

  PRIVATE SECTION.
    DATA left_operand  TYPE REF TO zif_xlom__ex.
    DATA right_operand TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_op_ampersand IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_op_ampersand( ).
    result->left_operand      = left_operand.
    result->right_operand     = right_operand.
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-operation-ampersand.
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
    TRY.
        result = zif_xlom__ex~set_result(
                     zcl_xlom__va_string=>create(
                         zcl_xlom__va=>to_string( arguments[ name = 'LEFT' ]-object )->get_string( )
                      && zcl_xlom__va=>to_string( arguments[ name = 'RIGHT' ]-object )->get_string( ) ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    IF     expression       IS BOUND
       AND expression->type  = zif_xlom__ex=>c_type-operation-ampersand
       AND left_operand->is_equal( CAST zcl_xlom__ex_op_ampersand( expression )->left_operand )
       AND right_operand->is_equal( CAST zcl_xlom__ex_op_ampersand( expression )->right_operand ).
      result = abap_true.
    ELSE.
      result = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
