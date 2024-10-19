"! INDEX(array, row_num, [column_num])
"! If row_num is omitted, column_num is required.
"! If column_num is omitted, row_num is required.
"! row_num = 0 is interpreted the same way as row_num = 1. Same remark for column_num.
"! row_num < 0 or column_num < 0 lead to #VALUE!
"! row_num and column_num with values outside the array lead to #REF!
"! https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
CLASS zcl_xlom__ex_fu_index DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING array         TYPE REF TO zif_xlom__ex
                row_num       TYPE REF TO zif_xlom__ex OPTIONAL
                column_num    TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_index.

  PRIVATE SECTION.
    DATA array      TYPE REF TO zif_xlom__ex.
    DATA row_num    TYPE REF TO zif_xlom__ex.
    DATA column_num TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_index IMPLEMENTATION.
  METHOD create.
    IF    array IS NOT BOUND
       OR (     row_num    IS NOT BOUND
            AND column_num IS NOT BOUND ).
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    result = NEW zcl_xlom__ex_fu_index( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-index.
    result->array             = array.
    result->row_num           = row_num.
    result->column_num        = column_num.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
        expression = me
        context    = context
        operands   = VALUE #( ( name = 'ARRAY     ' object = array      not_part_of_result_array = abap_true )
                              ( name = 'ROW_NUM   ' object = row_num    )
                              ( name = 'COLUMN_NUM' object = column_num ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(array_or_range) = zcl_xlom__va=>to_array( arguments[ name = 'ARRAY' ]-object ).
        DATA(row) = zcl_xlom__va=>to_number( arguments[ name = 'ROW_NUM' ]-object )->get_number( ).
        DATA(column) = zcl_xlom__va=>to_number( arguments[ name = 'COLUMN_NUM' ]-object )->get_number( ).
        result = zif_xlom__ex~set_result( array_or_range->get_cell_value( column = EXACT #( column )
                                                                          row    = EXACT #( row ) ) ).
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
