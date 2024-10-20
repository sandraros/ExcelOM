"! ROW([reference])
"! https://support.microsoft.com/en-us/office/row-function-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d
CLASS zcl_xlom__ex_fu_row DEFINITION
  PUBLIC FINAL
  INHERITING FROM zcl_xlom__ex_fu
*  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_fu.

  PUBLIC SECTION.
    CLASS-METHODS create
      IMPORTING !reference    TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_row.

    METHODs zif_xlom__ex~evaluate_single REDEFINITION.

  PROTECTED SECTION.
    METHODS constructor.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        reference TYPE i VALUE 1,
      END OF c_arg.
*    DATA reference TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_row IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    zif_xlom__ex~type = zif_xlom__ex=>c_type-function-row.
    zif_xlom__ex~parameters = VALUE #( ( name = 'REFERENCE' ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_fu_row( ).
    result->zif_xlom__ex~arguments_or_operands = VALUE #( ( REFERENCE ) ).
    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = result->zif_xlom__ex~arguments_or_operands ).
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-row.
*    result->reference         = reference.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'REFERENCE' object = reference ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                             context   = context ).
*    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    DATA(reference) = arguments[ c_arg-reference ].
    IF reference->type = reference->c_type-empty.
      result = zcl_xlom__va_number=>create( EXACT #( context->containing_cell-row ) ).
    ELSE.
      result = zcl_xlom__va_number=>create( ( CAST zcl_xlom_range( reference )->row( ) ) ).
    ENDIF.
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
