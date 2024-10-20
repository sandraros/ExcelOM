"! Operator colon (e.g. A1:A2, OFFSET(A1,1,1):OFFSET(A1,2,2), my.B1:my.C1 (range names), etc.)
CLASS zcl_xlom__ex_op_colon DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_op.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.
    INTERFACES zif_xlom__ex.
    INTERFACES zif_xlom__ex_array.

    CLASS-METHODS create
      IMPORTING left_operand  TYPE REF TO zif_xlom__ex
                right_operand TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_op_colon.

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


CLASS zcl_xlom__ex_op_colon IMPLEMENTATION.
  METHOD constructor.
    zif_xlom__ex~type = zif_xlom__ex=>c_type-operation-colon.
    zif_xlom__ex~parameters = VALUE #( ( name = 'LEFT_OPERAND'  )
                                       ( name = 'RIGHT_OPERAND' ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_op_colon( ).
    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = result->zif_xlom__ex~arguments_or_operands ).
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-operation-plus.
*    result->left_operand      = left_operand.
*    result->right_operand     = right_operand.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    DATA(cell1) = SWITCH #( left_operand->type
*                            WHEN left_operand->c_type-number THEN
*                              |{ CAST zcl_xlom__va_number( left_operand->evaluate( context ) )->get_integer( ) }|
*                            WHEN left_operand->c_type-array
*                              OR left_operand->c_type-range THEN
*                              CAST zcl_xlom__ex_el_range( left_operand )->_address_or_name
*                            ELSE
*                              THROW zcx_xlom_todo( ) ).
*    DATA(cell2) = SWITCH #( right_operand->type
*                            WHEN right_operand->c_type-number THEN
*                              |{ CAST zcl_xlom__va_number( right_operand->evaluate( context ) )->get_integer( ) }|
*                            WHEN left_operand->c_type-array
*                              OR left_operand->c_type-range THEN
*                              CAST zcl_xlom__ex_el_range( right_operand )->_address_or_name
*                            ELSE
*                              THROW zcx_xlom_todo( ) ).
*    TRY.
*        result = zif_xlom__ex~set_result( zcl_xlom_range=>create( zcl_xlom_range=>create_from_address_or_name(
*                                                                      address     = |{ cell1 }:{ cell2 }|
*                                                                      relative_to = context->worksheet ) ) ).
*      CATCH zcx_xlom__va INTO DATA(error).
*        result = error->result_error.
*    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    DATA(left_operand) = zif_xlom__ex~arguments_or_operands[ c_arg-left_operand ].
    DATA(right_operand) = zif_xlom__ex~arguments_or_operands[ c_arg-right_operand ].
    DATA(cell1) = SWITCH #( left_operand->type
                            WHEN left_operand->c_type-number THEN
                              |{ CAST zcl_xlom__va_number( zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                                             expression = left_operand
                                                             context    = context ) )->get_integer( ) }|
                            WHEN left_operand->c_type-array
                              OR left_operand->c_type-range THEN
                              CAST zcl_xlom__ex_el_range( left_operand )->_address_or_name
                            ELSE
                              THROW zcx_xlom_todo( ) ).
    DATA(cell2) = SWITCH #( right_operand->type
                            WHEN right_operand->c_type-number THEN
                              |{ CAST zcl_xlom__va_number( zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                                             expression = right_operand
                                                             context    = context ) )->get_integer( ) }|
                            WHEN left_operand->c_type-array
                              OR left_operand->c_type-range THEN
                              CAST zcl_xlom__ex_el_range( right_operand )->_address_or_name
                            ELSE
                              THROW zcx_xlom_todo( ) ).
    TRY.
        result = zcl_xlom_range=>create( zcl_xlom_range=>create_from_address_or_name(
                                                                      address     = |{ cell1 }:{ cell2 }|
                                                                      relative_to = context->worksheet ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
    zif_xlom__ex~result_of_evaluation = result.
*    RAISE EXCEPTION TYPE zcx_xlom_unexpected.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    IF     expression       IS BOUND
*       AND expression->type  = zif_xlom__ex=>c_type-operation-plus
*       AND left_operand->is_equal( CAST zcl_xlom__ex_op_colon( expression )->left_operand )
*       AND right_operand->is_equal( CAST zcl_xlom__ex_op_colon( expression )->right_operand ).
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
