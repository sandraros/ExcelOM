*"* use this source file for your ABAP unit test classes

CLASS ltc_lexer DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS arithmetic                     FOR TESTING RAISING cx_static_check.
    METHODS array                          FOR TESTING RAISING cx_static_check.
    METHODS array_two_rows                 FOR TESTING RAISING cx_static_check.
    METHODS error_name                     FOR TESTING RAISING cx_static_check.
    METHODS function                       FOR TESTING RAISING cx_static_check.
    METHODS function_function              FOR TESTING RAISING cx_static_check.
    METHODS function_optional_argument     FOR TESTING RAISING cx_static_check.
    METHODS number                         FOR TESTING RAISING cx_static_check.
    METHODS operator_function              FOR TESTING RAISING cx_static_check.
    METHODS range                          FOR TESTING RAISING cx_static_check.
    METHODS smart_table                    FOR TESTING RAISING cx_static_check.
    METHODS smart_table_all                FOR TESTING RAISING cx_static_check.
    METHODS smart_table_column             FOR TESTING RAISING cx_static_check.
    METHODS smart_table_no_space           FOR TESTING RAISING cx_static_check.
    METHODS smart_table_space_separator    FOR TESTING RAISING cx_static_check.
    METHODS smart_table_space_boundaries   FOR TESTING RAISING cx_static_check.
    METHODS smart_table_space_all          FOR TESTING RAISING cx_static_check.
    METHODS text_literal                   FOR TESTING RAISING cx_static_check.
    METHODS text_literal_with_double_quote FOR TESTING RAISING cx_static_check.
    METHODS very_long                      FOR TESTING RAISING cx_static_check.

    TYPES tt_parenthesis_group TYPE zcl_xlom__ex_ut_lexer=>tt_parenthesis_group.
    TYPES tt_token             TYPE zcl_xlom__ex_ut_lexer=>tt_token.
    TYPES ts_result_lexe       TYPE zcl_xlom__ex_ut_lexer=>ts_result_lexe.

    CONSTANTS c_type LIKE zcl_xlom__ex_ut_lexer=>c_type VALUE zcl_xlom__ex_ut_lexer=>c_type.

    METHODS lexe
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE tt_token. " ts_result_lexe.

ENDCLASS.


CLASS ltc_lexer IMPLEMENTATION.
  METHOD arithmetic.
    cl_abap_unit_assert=>assert_equals( act = lexe( '2*(1+3*(5+1))' )
                                        exp = VALUE tt_token( ( value = `2`  type = c_type-number )
                                                              ( value = `*`  type = c_type-operator )
                                                              ( value = `(`  type = c_type-parenthesis_open )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `+`  type = c_type-operator )
                                                              ( value = `3`  type = c_type-number )
                                                              ( value = `*`  type = c_type-operator )
                                                              ( value = `(`  type = c_type-parenthesis_open )
                                                              ( value = `5`  type = c_type-number )
                                                              ( value = `+`  type = c_type-operator )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `)`  type = c_type-parenthesis_close )
                                                              ( value = `)`  type = c_type-parenthesis_close ) ) ).
  ENDMETHOD.

  METHOD array.
    cl_abap_unit_assert=>assert_equals( act = lexe( '{1,2}' )
                                        exp = VALUE tt_token( ( value = `{` type = c_type-curly_bracket_open )
                                                              ( value = `1` type = c_type-number )
                                                              ( value = `,` type = c_type-comma )
                                                              ( value = `2` type = c_type-number )
                                                              ( value = `}` type = c_type-curly_bracket_close ) ) ).
  ENDMETHOD.

  METHOD array_two_rows.
    cl_abap_unit_assert=>assert_equals( act = lexe( '{1,2;3,4}' )
                                        exp = VALUE tt_token( ( value = `{` type = c_type-curly_bracket_open )
                                                              ( value = `1` type = c_type-number )
                                                              ( value = `,` type = c_type-comma )
                                                              ( value = `2` type = c_type-number )
                                                              ( value = `;` type = c_type-semicolon )
                                                              ( value = `3` type = c_type-number )
                                                              ( value = `,` type = c_type-comma )
                                                              ( value = `4` type = c_type-number )
                                                              ( value = `}` type = c_type-curly_bracket_close ) ) ).
  ENDMETHOD.

  METHOD error_name.
    cl_abap_unit_assert=>assert_equals( act = lexe( '#N/A!' )
                                        exp = VALUE tt_token( ( value = `#N/A!` type = c_type-error_name ) ) ).
  ENDMETHOD.

  METHOD function.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'IF(1=1,0,1)' )
                                        exp = VALUE tt_token( ( value = `IF` type = c_type-function_name )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `=`  type = c_type-operator )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `0`  type = c_type-number )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `)`  type = ')' ) ) ).
  ENDMETHOD.

  METHOD function_function.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'LEN(RIGHT("text",2))' )
                                        exp = VALUE tt_token( ( value = `LEN`   type = c_type-function_name )
                                                              ( value = `RIGHT` type = c_type-function_name )
                                                              ( value = `text`  type = c_type-text_literal )
                                                              ( value = `,`     type = c_type-comma )
                                                              ( value = `2`     type = c_type-number )
                                                              ( value = `)`     type = c_type-parenthesis_close )
                                                              ( value = `)`     type = c_type-parenthesis_close ) ) ).
  ENDMETHOD.

  METHOD function_optional_argument.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'RIGHT("text",)' )
                                        exp = VALUE tt_token( ( value = `RIGHT` type = c_type-function_name )
                                                              ( value = `text`  type = c_type-text_literal )
                                                              ( value = `,`     type = c_type-comma )
                                                              ( value = `)`     type = c_type-parenthesis_close ) ) ).
  ENDMETHOD.

  METHOD lexe.
    result = zcl_xlom__ex_ut_lexer=>create( )->lexe( text ).
  ENDMETHOD.

  METHOD number.
    cl_abap_unit_assert=>assert_equals( act = lexe( '25' )
                                        exp = VALUE tt_token( ( value = `25` type = c_type-number ) ) ).

    cl_abap_unit_assert=>assert_equals( act = lexe( '-1' )
                                        exp = VALUE tt_token( ( value = `-` type = c_type-operator )
                                                              ( value = `1` type = c_type-number ) ) ).
  ENDMETHOD.

  METHOD operator_function.
    cl_abap_unit_assert=>assert_equals( act = lexe( '1+LEN("text")' )
                                        exp = VALUE tt_token( ( value = `1`    type = c_type-number )
                                                              ( value = `+`    type = c_type-operator )
                                                              ( value = `LEN`  type = c_type-function_name )
                                                              ( value = `text` type = c_type-text_literal )
                                                              ( value = `)`    type = c_type-parenthesis_close ) ) ).
  ENDMETHOD.

  METHOD range.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Sheet1!$A$1' )
                                        exp = VALUE tt_token( ( value = `Sheet1!$A$1` type = 'W' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( `'Sheet 1'!$A$1` )
                                        exp = VALUE tt_token( ( value = `'Sheet 1'!$A$1` type = 'W' ) ) ).
  ENDMETHOD.

  METHOD smart_table.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[]' )
                                        exp = VALUE tt_token( ( value = `Table1` type = c_type-table_name )
                                                              ( value = `]`      type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_all.
    cl_abap_unit_assert=>assert_equals(
        act = lexe( 'Table1[[#All]]' )
        exp = VALUE tt_token( ( value = `Table1` type = c_type-table_name )
                              ( value = `[#All]` type = c_type-square_bracket_open )
                              ( value = `]`      type = c_type-square_bracket_close ) ) ).
  ENDMETHOD.

  METHOD smart_table_column.
    cl_abap_unit_assert=>assert_equals(
        act = lexe( 'Table1[Column1]' )
        exp = VALUE tt_token( ( value = `Table1`  type = c_type-table_name )
                              ( value = `[Column1]` type = c_type-square_bracket_open )
                              ( value = `]`      type = c_type-square_bracket_close ) ) ).
  ENDMETHOD.

  METHOD smart_table_no_space.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    cl_abap_unit_assert=>assert_equals(
        act = lexe( `DeptSales[[#Headers],[#Data],[% Commission]]` )
        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[#Data]`        type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[% Commission]` type = c_type-square_bracket_open )
                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_space_all.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    DATA(act) = lexe( `DeptSales[ [#Headers], [#Data], [% Commission] ]` ).
    cl_abap_unit_assert=>assert_equals(
        act = act
        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[#Data]`        type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[% Commission]` type = c_type-square_bracket_open )
                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_space_boundaries.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    DATA(act) = lexe( `DeptSales[ [#Headers],[#Data],[% Commission] ]` ).
    cl_abap_unit_assert=>assert_equals(
        act = act
        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[#Data]`        type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[% Commission]` type = c_type-square_bracket_open )
                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_space_separator.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    cl_abap_unit_assert=>assert_equals(
        act = lexe( `DeptSales[[#Headers], [#Data], [% Commission]]` )
        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[#Data]`        type = c_type-square_bracket_open )
                              ( value = `,`              type = `,` )
                              ( value = `[% Commission]` type = c_type-square_bracket_open )
                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD text_literal.
    cl_abap_unit_assert=>assert_equals( act = lexe( '"IF(1=1,0,1)"' )
                                        exp = VALUE tt_token( ( value = `IF(1=1,0,1)` type = c_type-text_literal ) ) ).
  ENDMETHOD.

  METHOD text_literal_with_double_quote.
    cl_abap_unit_assert=>assert_equals(
        act = lexe( '"IF(A1=""X"",0,1)"' )
        exp = VALUE tt_token( ( value = `IF(A1="X",0,1)` type = c_type-text_literal ) ) ).
  ENDMETHOD.

  METHOD very_long.
    cl_abap_unit_assert=>assert_equals( act = lexe( |(a{ repeat( val = ',a'
                                                                 occ = 5000 )
                                                    })| )
                                        exp = VALUE tt_token( ( value = `(` type = '(' )
                                                              ( value = `a` type = 'W' )
                                                              ( LINES OF VALUE
                                                                tt_token( FOR i = 1 WHILE i <= 5000
                                                                          ( value = `,` type = ',' )
                                                                          ( value = `a` type = 'W' ) ) )
                                                              ( value = `)` type = ')' ) ) ).
  ENDMETHOD.
ENDCLASS.
