CLASS zcl_xlom__ex_fu DEFINITION
  PUBLIC
*  FINAL
  CREATE PROTECTED .

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create_dynamic
      IMPORTING function_name TYPE csequence
                arguments     TYPE zif_xlom__ex=>tt_argument_or_operand
      RETURNING VALUE(result) TYPE REF TO zif_xlom__ex.

  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_xlom__ex_fu IMPLEMENTATION.
  METHOD create_dynamic.
    data function TYPE REF TO ZCL_xlom__ex_fu.

    CASE function_name.
      WHEN 'ADDRESS'.  result = NEW zcl_xlom__ex_fu_address( ).
*      WHEN 'CELL'.     result = NEW zcl_xlom__ex_fu_cell( ).
*      WHEN 'COUNTIF'.  result = NEW zcl_xlom__ex_fu_countif( ).
*      WHEN 'FIND'.     result = NEW zcl_xlom__ex_fu_find( ).
*      WHEN 'IF'.       result = NEW zcl_xlom__ex_fu_if( ).
*      WHEN 'IFERROR'.  result = NEW zcl_xlom__ex_fu_iferror( ).
*      WHEN 'INDEX'.    result = NEW zcl_xlom__ex_fu_index( ).
*      WHEN 'INDIRECT'. result = NEW zcl_xlom__ex_fu_indirect( ).
*      WHEN 'LEN'.      result = NEW zcl_xlom__ex_fu_len( ).
*      WHEN 'MATCH'.    result = NEW zcl_xlom__ex_fu_match( ).
*      WHEN 'OFFSET'.   result = NEW zcl_xlom__ex_fu_offset( ).
*      WHEN 'RIGHT'.    result = NEW zcl_xlom__ex_fu_right( ).
*      WHEN 'ROW'.      result = NEW zcl_xlom__ex_fu_row( ).
*      WHEN 'T'.        result = NEW zcl_xlom__ex_fu_t( ).
      WHEN OTHERS.
        TRY.
            DATA(function_class_name) = |ZCL_XLOM__EX_FU_{ function_name }|.
            CREATE OBJECT result TYPE (function_class_name).
          CATCH cx_root.
            RAISE EXCEPTION TYPE zcx_xlom_todo.
        ENDTRY.
    ENDCASE.

    function ?= result.
    function->zif_xlom__ex~arguments_or_operands = arguments.

    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = function->zif_xlom__ex~arguments_or_operands ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.
ENDCLASS.
