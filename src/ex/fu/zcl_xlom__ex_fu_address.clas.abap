"! ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
"! https://support.microsoft.com/en-us/office/address-function-d0c26c0d-3991-446b-8de4-ab46431d4f89
CLASS zcl_xlom__ex_fu_address DEFINITION
  PUBLIC FINAL
  INHERITING FROM zcl_xlom__ex_fu
*  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_fu.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.
*    INTERFACES zif_xlom__ex.

    "! @parameter row_num    | Required. A numeric value that specifies the row number to use in the cell reference.
    "! @parameter column_num    | Required. A numeric value that specifies the column number to use in the cell reference.
    "! @parameter abs_num    | Optional. A numeric value that specifies the type of reference to return.
    "!                       <ul>
    "!                       <li>1 or omitted: Absolute          </li>
    "!                       <li>2: Absolute row; relative column</li>
    "!                       <li>3: Relative row; absolute column</li>
    "!                       <li>4: Relative                     </li>
    "!                       </ul>
    "! @parameter A1    | Optional. A logical value that specifies the A1 or R1C1 reference style. In A1 style, columns
    "!                    are labeled alphabetically, and rows are labeled numerically. In R1C1 reference style, both
    "!                    columns and rows are labeled numerically.
    "!                       <ul>
    "!                       <li>TRUE or omitted: the ADDRESS function returns an A1-style reference</li>
    "!                       <li>FALSE: the ADDRESS function returns an R1C1-style reference</li>
    "!                       </ul>
    "!                    Note: To change the reference style that Excel uses, click the File tab, click Options, and then click
    "!                          Formulas. Under Working with formulas, select or clear the R1C1 reference style check box.
    "! @parameter sheet_text    | Optional. A text value that specifies the name of the worksheet to be used as the external
    "!                            reference. For example, the formula =ADDRESS(1,1,,,"Sheet2") returns Sheet2!$A$1. If the
    "!                            sheet_text argument is omitted, no sheet name is used, and the address returned by the
    "!                            function refers to a cell on the current sheet.
    CLASS-METHODS create
      IMPORTING row_num       TYPE REF TO zif_xlom__ex
                column_num    TYPE REF TO zif_xlom__ex
                abs_num       TYPE REF TO zif_xlom__ex OPTIONAL
                a1            TYPE REF TO zif_xlom__ex OPTIONAL
                sheet_text    TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_address.

    METHODs zif_xlom__ex~evaluate_single REDEFINITION.

  PROTECTED SECTION.
    METHODS constructor.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        row_num    TYPE i VALUE 1,
        column_num TYPE i VALUE 2,
        abs_num    TYPE i VALUE 3,
        a1         TYPE i VALUE 4,
        sheet_text TYPE i VALUE 5,
      END OF c_arg.

    CONSTANTS:
      BEGIN OF c_abs,
        absolute                     TYPE i VALUE 1,
        absolute_row_relative_column TYPE i VALUE 2,
        relative_row_absolute_column TYPE i VALUE 3,
        relative                     TYPE i VALUE 4,
      END OF c_abs.

*    METHODS constructor.
*    DATA row_num    TYPE REF TO zif_xlom__ex.
*    DATA column_num TYPE REF TO zif_xlom__ex.
*    DATA abs_num    TYPE REF TO zif_xlom__ex.
*    DATA a1         TYPE REF TO zif_xlom__ex.
*    DATA sheet_text TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_address IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    zif_xlom__ex~type = zif_xlom__ex=>c_type-function-address.
    zif_xlom__ex~parameters = VALUE #( ( name = 'ROW_NUM   ' )
                                       ( name = 'COLUMN_NUM' )
                                       ( name = 'ABS_NUM   ' default = zcl_xlom__ex_el_number=>create( 1 ) )
                                       ( name = 'A1        ' default = zcl_xlom__ex_el_boolean=>true )
                                       ( name = 'SHEET_TEXT' default = zcl_xlom__ex_el_string=>create( '' ) ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_fu_address( ).
    result->zif_xlom__ex~arguments_or_operands = VALUE #( ( row_num    )
                                                          ( column_num )
                                                          ( abs_num    )
                                                          ( a1         )
                                                          ( sheet_text ) ).
    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = result->zif_xlom__ex~arguments_or_operands ).
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-address.
*    result->row_num           = row_num.
*    result->column_num        = column_num.
*    result->abs_num           = abs_num.
*    result->a1                = a1.
*    result->sheet_text        = sheet_text.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    " ADDRESS({1;5};{2;5};{1;2};{0;1};{"Sheet1";"Sheet2"})
*    " returns
*    " Sheet1!R1C2 (ADDRESS(1;2;1;0;"Sheet1"))
*    " Sheet2!E$5  (ADDRESS(5;5;2;1;"Sheet2"))
*    " TODO: variable is assigned but never used (ABAP cleaner)
*    result = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
*                 expression = me
*                 context    = context
*                 operands   = VALUE #( ( name = 'ROW_NUM   ' object = row_num    )
*                                       ( name = 'COLUMN_NUM' object = column_num )
*                                       ( name = 'ABS_NUM   ' object = abs_num    )
*                                       ( name = 'A1        ' object = a1         )
*                                       ( name = 'SHEET_TEXT' object = sheet_text ) ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(row_num) = zcl_xlom__va=>to_number( arguments[ c_arg-ROW_NUM ] )->get_integer( ).
        DATA(column_num) = zcl_xlom__va=>to_number( arguments[ c_arg-COLUMN_NUM ] )->get_integer( ).
        DATA(abs_num) = zcl_xlom__va=>to_number( arguments[ c_arg-ABS_NUM ] )->get_integer( ).
        DATA(a1) = zcl_xlom__va=>to_boolean( arguments[ c_arg-A1 ] )->boolean_value.
        DATA(sheet_text) = zcl_xlom__va=>to_string( arguments[ c_arg-SHEET_TEXT ] )->get_string( ).

        IF    row_num NOT    BETWEEN 1 AND zcl_xlom_worksheet=>max_rows
           OR column_num NOT BETWEEN 1 AND zcl_xlom_worksheet=>max_columns
           OR abs_num NOT    BETWEEN 0 AND 4
           OR abs_num NOT    BETWEEN 0 AND 1.
          result = zcl_xlom__va_error=>value_cannot_be_calculated.
          RETURN.
        ENDIF.

        case abs_num.
          when c_abs-absolute.
            result = zcl_xlom__va_string=>get( |${ zcl_xlom_range=>convert_column_number_to_a_xfd( column_num ) }${ row_num }| ).
          when others.
            raise EXCEPTION type zcx_xlom_todo.
        ENDCASE.

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
*    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.
ENDCLASS.
