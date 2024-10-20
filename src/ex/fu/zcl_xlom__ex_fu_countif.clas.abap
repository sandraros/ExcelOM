"! COUNTIF(range, criteria)
"! https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34
CLASS zcl_xlom__ex_fu_countif DEFINITION
  PUBLIC FINAL
  INHERITING FROM zcl_xlom__ex_fu
*  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_fu.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    CLASS-METHODS create
      IMPORTING !range        TYPE REF TO zif_xlom__ex
                criteria      TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_countif.

    METHODs zif_xlom__ex~evaluate_single REDEFINITION.

  PROTECTED SECTION.
    METHODS constructor.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        RANGE    TYPE i VALUE 1,
        CRITERIA TYPE i VALUE 2,
      END OF c_arg.
*    DATA range    TYPE REF TO zif_xlom__ex.
*    DATA criteria TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_countif IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    zif_xlom__ex~type = zif_xlom__ex=>c_type-function-countif.
    zif_xlom__ex~parameters = VALUE #( ( name = 'RANGE   ' )
                                       ( name = 'CRITERIA' ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_fu_countif( ).
    result->zif_xlom__ex~arguments_or_operands = VALUE #( ( RANGE    )
                                                          ( CRITERIA ) ).
    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = result->zif_xlom__ex~arguments_or_operands ).
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-countif.
*    result->range             = range.
*    result->criteria          = criteria.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
*        expression = me
*        context    = context
*        operands   = VALUE #( ( name = 'RANGE   ' object = range    not_part_of_result_array = abap_true )
*                              ( name = 'CRITERIA' object = criteria ) ) ).
*    result = zif_xlom__ex~set_result(
*                 COND #( WHEN array_evaluation-result IS BOUND
*                         THEN array_evaluation-result
*                         ELSE zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                            context   = context ) ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(result_of_range) = CAST zif_xlom__va_array( arguments[ C_ARG-RANGE ] ).
        DATA(result_of_criteria) = zcl_xlom__va=>to_string( arguments[ C_ARG-CRITERIA ] )->get_string( ).
        " criteria may start with the comparison operator =, <, <=, >, >=, <>.
        " It's followed by the string to compare.
        " Anything else will be equivalent to the same string prefixed with the = comparison operator.
        IF strlen( result_of_criteria ) >= 1.
          CASE substring( val = result_of_criteria
                          len = 1 ).
            WHEN '=' OR '<' OR '>'.
              RAISE EXCEPTION TYPE zcx_xlom_todo.
          ENDCASE.
        ENDIF.
        DATA(count_cells) = 0.
        DATA(row_number) = 1.
        WHILE row_number <= result_of_range->row_count.
          DATA(column_number) = 1.
          WHILE column_number <= result_of_range->column_count.
            DATA(cell_value_string) = zcl_xlom__va=>to_string( result_of_range->get_cell_value(
                                                                   column = column_number
                                                                   row    = row_number ) )->get_string( ).
            IF cell_value_string CP result_of_criteria.
              count_cells = count_cells + 1.
            ENDIF.
            column_number = column_number + 1.
          ENDWHILE.
          row_number = row_number + 1.
        ENDWHILE.
        result = zcl_xlom__va_number=>create( CONV #( count_cells ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
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
