"! ROW([reference])
"! https://support.microsoft.com/en-us/office/row-function-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d
CLASS zcl_xlom__ex_fu_row DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !reference    TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_row.

  PRIVATE SECTION.
    DATA reference TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_row IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_row( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-row.
    result->reference         = reference.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = me
                                 context    = context
                                 operands   = VALUE #( ( name = 'REFERENCE' object = reference ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    DATA temp_result TYPE REF TO zif_xlom__va.

    IF reference IS NOT BOUND.
      temp_result = zcl_xlom__va_number=>create( EXACT #( context->containing_cell-row ) ).
    ELSE.
      DATA(reference_result) = CAST zcl_xlom_range( arguments[ name = 'REFERENCE' ]-object ).
      temp_result = zcl_xlom__va_number=>create( reference_result->row( ) ).
    ENDIF.
    result = zif_xlom__ex~set_result( temp_result ).
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
