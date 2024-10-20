"! OFFSET(reference, rows, cols, [height], [width])
"! OFFSET($A$1,0,0,5,0) is equivalent to $A$1:$A$5
"! https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66
CLASS zcl_xlom__ex_fu_offset DEFINITION
  PUBLIC FINAL
  INHERITING FROM zcl_xlom__ex_fu
*  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_fu.

  PUBLIC SECTION.
    CLASS-METHODS create
      IMPORTING !reference    TYPE REF TO zif_xlom__ex
                !rows         TYPE REF TO zif_xlom__ex
                cols          TYPE REF TO zif_xlom__ex
                height        TYPE REF TO zif_xlom__ex OPTIONAL
                !width        TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_offset.

    METHODs zif_xlom__ex~evaluate_single REDEFINITION.

  PROTECTED SECTION.
    METHODS constructor.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        reference TYPE i VALUE 1,
        rows      TYPE i VALUE 2,
        cols      TYPE i VALUE 3,
        height    TYPE i VALUE 4,
        width     TYPE i VALUE 5,
      END OF c_arg.
*    DATA reference TYPE REF TO zif_xlom__ex.
*    DATA rows      TYPE REF TO zif_xlom__ex.
*    DATA cols      TYPE REF TO zif_xlom__ex.
*    DATA height    TYPE REF TO zif_xlom__ex.
*    DATA width     TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_offset IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    zif_xlom__ex~type = zif_xlom__ex=>c_type-function-offset.
    zif_xlom__ex~parameters = VALUE #( ( name = 'REFERENCE' )
                                       ( name = 'ROWS     ' )
                                       ( name = 'COLS     ' default = zcl_xlom__ex_el_number=>create( 1 ) )
                                       ( name = 'HEIGHT   ' default = zcl_xlom__ex_el_boolean=>true )
                                       ( name = 'WIDTH    ' default = zcl_xlom__ex_el_string=>create( '' ) ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_fu_offset( ).
    result->zif_xlom__ex~arguments_or_operands = VALUE #( ( reference )
                                                          ( rows      )
                                                          ( cols      )
                                                          ( height    )
                                                          ( width     ) ).
    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = result->zif_xlom__ex~arguments_or_operands ).
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-offset.
*    result->reference         = reference.
*    result->rows              = rows.
*    result->cols              = cols.
*    result->height            = height.
*    result->width             = width.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
*        expression = me
*        context    = context
*        operands   = VALUE #( ( name = 'REFERENCE' object = reference not_part_of_result_array = abap_true )
*                              ( name = 'ROWS     ' object = rows      )
*                              ( name = 'COLS     ' object = cols      )
*                              ( name = 'HEIGHT   ' object = height    )
*                              ( name = 'WIDTH    ' object = width     ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                             context   = context ).
*    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(rows_result) = zcl_xlom__va=>to_number( arguments[ c_arg-ROWS ] )->get_integer( ).
        DATA(cols_result) = zcl_xlom__va=>to_number( arguments[ c_arg-COLS ] )->get_integer( ).
        DATA(reference_result) = zcl_xlom__va=>to_range( input = arguments[ c_arg-REFERENCE ] ).
        DATA(height) = arguments[ c_arg-height ].
        DATA(width) = arguments[ c_arg-width ].
        DATA(height_result) = COND #( WHEN height       IS BOUND
                                       AND height->type <> height->c_type-empty
                                      THEN zcl_xlom__va=>to_number( arguments[
                                                                        c_arg-HEIGHT ] )->get_integer( )
                                      ELSE reference_result->rows( )->count( ) ).
        DATA(width_result) = COND #( WHEN width       IS BOUND
                                      AND width->type <> width->c_type-empty
                                     THEN zcl_xlom__va=>to_number( arguments[ c_arg-WIDTH ] )->get_integer( )
                                     ELSE reference_result->columns( )->count( ) ).
        result = reference_result->offset( row_offset    = rows_result
                                           column_offset = cols_result
                                         )->resize( row_size    = height_result
                                                    column_size = width_result ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
    zif_xlom__ex~result_of_evaluation = result.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    IF expression->type <> zif_xlom__ex=>c_type-function-offset.
*      RETURN.
*    ENDIF.
*    DATA(compare_offset) = CAST zcl_xlom__ex_fu_offset( expression ).
*
*    result = xsdbool(     reference->is_equal( compare_offset->reference )
*                      AND rows->is_equal( compare_offset->rows )
*                      AND cols->is_equal( compare_offset->cols )
*                      AND zcl_xlom__ex_ut=>are_equal( expression_1 = height
*                                                      expression_2 = compare_offset->height )
*                      AND zcl_xlom__ex_ut=>are_equal( expression_1 = width
*                                                      expression_2 = compare_offset->width ) ).
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~set_result.
*    zif_xlom__ex~result_of_evaluation = value.
*    result = value.
  ENDMETHOD.
ENDCLASS.
