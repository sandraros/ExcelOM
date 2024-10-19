*"* use this source file for your ABAP unit test classes

CLASS ltc_evaluate DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

  PRIVATE SECTION.
    METHODS ampersand                  FOR TESTING RAISING cx_static_check.
    METHODS array                      FOR TESTING RAISING cx_static_check.
    METHODS cell                       FOR TESTING RAISING cx_static_check.
    METHODS colon                      FOR TESTING RAISING cx_static_check.
    METHODS complex_1                  FOR TESTING RAISING cx_static_check.
    METHODS complex_2                  FOR TESTING RAISING cx_static_check.
    METHODS countif                    FOR TESTING RAISING cx_static_check.
    METHODS equal                      FOR TESTING RAISING cx_static_check.
    METHODS error                      FOR TESTING RAISING cx_static_check.
    METHODS find                       FOR TESTING RAISING cx_static_check.
    METHODS function_optional_argument FOR TESTING RAISING cx_static_check.
    METHODS if                         FOR TESTING RAISING cx_static_check.
    METHODS iferror                    FOR TESTING RAISING cx_static_check.
    METHODS index                      FOR TESTING RAISING cx_static_check.
    "! Array evaluation, e.g. INDEX(A1:B2,{2,1},{2,1})
    METHODS index_ae                   FOR TESTING RAISING cx_static_check.
    METHODS indirect                   FOR TESTING RAISING cx_static_check.
    METHODS len                        FOR TESTING RAISING cx_static_check.
    METHODS len_a1_a2                  FOR TESTING RAISING cx_static_check.
    METHODS match                      FOR TESTING RAISING cx_static_check.
    METHODS match_2                    FOR TESTING RAISING cx_static_check.
    METHODS minus                      FOR TESTING RAISING cx_static_check.
    METHODS mult                       FOR TESTING RAISING cx_static_check.
    METHODS number                     FOR TESTING RAISING cx_static_check.
    METHODS offset                     FOR TESTING RAISING cx_static_check.
    METHODS plus                       FOR TESTING RAISING cx_static_check.
    METHODS range_a1_plus_one          FOR TESTING RAISING cx_static_check.
    METHODS range_two_sheets           FOR TESTING RAISING cx_static_check.
    METHODS right                      FOR TESTING RAISING cx_static_check.
    METHODS right_2                    FOR TESTING RAISING cx_static_check.
    METHODS row                        FOR TESTING RAISING cx_static_check.
    METHODS string                     FOR TESTING RAISING cx_static_check.
    METHODS t                          FOR TESTING RAISING cx_static_check.

    TYPES tt_parenthesis_group TYPE zcl_xlom__ex_ut_lexer=>tt_parenthesis_group.
    TYPES tt_token             TYPE zcl_xlom__ex_ut_lexer=>tt_token.
    TYPES ts_result_lexe       TYPE zcl_xlom__ex_ut_lexer=>ts_result_lexe.

    DATA worksheet   TYPE REF TO zcl_xlom_worksheet.
    DATA range_a1    TYPE REF TO zcl_xlom_range.
    DATA range_a2    TYPE REF TO zcl_xlom_range.
    DATA range_b1    TYPE REF TO zcl_xlom_range.
    DATA range_b2    TYPE REF TO zcl_xlom_range.
    DATA range_c1    TYPE REF TO zcl_xlom_range.
    DATA range_d1    TYPE REF TO zcl_xlom_range.
    DATA application TYPE REF TO zcl_xlom_application.
    DATA workbook    TYPE REF TO zcl_xlom_workbook.

    METHODS assert_equals
      IMPORTING act           TYPE REF TO zif_xlom__va
                exp           TYPE REF TO zif_xlom__va
      RETURNING VALUE(result) TYPE abap_bool.

    METHODS setup.
ENDCLASS.


CLASS ltc_evaluate IMPLEMENTATION.
  METHOD ampersand.
    range_a1->set_formula2( value = `"hello "&"world"` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1 )->get_string( )
                                        exp = `hello world` ).
    range_a1->set_formula2( value = `"hello "&"new "&"world"` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = `hello new world` ).
  ENDMETHOD.

  METHOD array.
    range_a1->set_formula2( value = `{1,2}` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1 )->get_number( )
                                        exp = 1 ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_b1 )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD assert_equals.
    cl_abap_unit_assert=>assert_true( xsdbool( exp->is_equal( act ) ) ).
  ENDMETHOD.

  METHOD cell.
* TODO not very clear what CELL should do without the Reference argument...
*    range_a1->set_formula2( value = `CELL("filename")` ).
*    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va_converter=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = `\[]Sheet1` ).
    range_a2->set_value( zcl_xlom__va_string=>create( '' ) ).
    range_a1->set_formula2( value = `CELL("filename",A2)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = `\[]Sheet1` ).
  ENDMETHOD.

  METHOD colon.
    DATA dummy_ref_to_offset TYPE REF TO zcl_xlom__ex_op_colon ##NEEDED.
    DATA result              TYPE REF TO zif_xlom__va.

    result = application->evaluate( `3:3` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$3:$3' ).

    result = application->evaluate( `C:C` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$C:$C' ).
  ENDMETHOD.

  METHOD complex_1.
    range_a2->set_formula2(
        value = `"'"&RIGHT(CELL("filename",A1),LEN(CELL("filename",A1))-FIND("]",CELL("filename",A1)))&" (2)'!$1:$1"` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a2->value( ) )->get_string( )
                                        exp = `'Sheet1 (2)'!$1:$1` ).
  ENDMETHOD.

  METHOD complex_2.
    DATA dummy_ref_to_iferror   TYPE REF TO zcl_xlom__ex_fu_iferror ##NEEDED.
    DATA dummy_ref_to_t         TYPE REF TO zcl_xlom__ex_fu_t ##NEEDED.
    DATA dummy_ref_to_ampersand TYPE REF TO zcl_xlom__ex_op_ampersand ##NEEDED.
    DATA dummy_ref_to_index     TYPE REF TO zcl_xlom__ex_fu_index ##NEEDED.
    DATA dummy_ref_to_offset    TYPE REF TO zcl_xlom__ex_fu_offset ##NEEDED.
    DATA dummy_ref_to_indirect  TYPE REF TO zcl_xlom__ex_fu_indirect ##NEEDED.
    DATA dummy_ref_to_minus     TYPE REF TO zcl_xlom__ex_op_minus ##NEEDED.
    DATA dummy_ref_to_row       TYPE REF TO zcl_xlom__ex_fu_row ##NEEDED.
    DATA dummy_ref_to_match     TYPE REF TO zcl_xlom__ex_fu_match ##NEEDED.

    DATA(worksheet_bkpf) = workbook->worksheets->add( 'BKPF' ).
    worksheet_bkpf->range_from_address( 'A1' )->set_value( zcl_xlom__va_string=>create( 'ID_REF_TEST' ) ).
    worksheet_bkpf->range_from_address( 'B2' )->set_value( zcl_xlom__va_string=>create( `'BKPF (2)'!$1:$1` ) ).

    DATA(worksheet_bkpf_2) = workbook->worksheets->add( 'BKPF (2)' ).
    worksheet_bkpf_2->range_from_address( 'A1' )->set_value( zcl_xlom__va_string=>create( 'ID_REF_TEST' ) ).
    worksheet_bkpf_2->range_from_address( 'A3' )->set_value( zcl_xlom__va_string=>create( 'MY_TEST' ) ).

    DATA(range_bkpf_a3) = worksheet_bkpf->range_from_address( 'A3' ).
    range_bkpf_a3->set_formula2(
        value = `IFERROR(T(""&INDEX(OFFSET(INDIRECT(BKPF!$B$2),ROW()-1,0),1,MATCH(A$1,INDIRECT(BKPF!$B$2),0))),"")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_bkpf_a3->value( ) )->get_string( )
                                        exp = `MY_TEST` ).
  ENDMETHOD.

  METHOD countif.
    range_a1->set_value( zcl_xlom__va_string=>create( `Hello` ) ).
    range_a2->set_value( zcl_xlom__va_string=>create( `world` ) ).
    range_b1->set_value( zcl_xlom__va_string=>create( `peace` ) ).
    range_b2->set_value( zcl_xlom__va_string=>create( `love` ) ).
    range_c1->set_formula2( value = `COUNTIF(A1:B2,"*e*")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_c1->value( ) )->get_integer( )
                                        exp = 3 ).
  ENDMETHOD.

  METHOD equal.
    range_a1->set_formula2( value = `1=1` ).
    cl_abap_unit_assert=>assert_true( zcl_xlom__va=>to_boolean( range_a1->value( ) )->boolean_value ).
  ENDMETHOD.

  METHOD error.
    range_a1->set_formula2( value = `#N/A` ).
    cl_abap_unit_assert=>assert_equals( act = range_a1->value( )->type
                                        exp = zif_xlom__va=>c_type-error ).
  ENDMETHOD.

  METHOD find.
    range_a1->set_formula2( value = `FIND("b","abc")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD function_optional_argument.
    range_a1->set_formula2( value = `RIGHT("hello",0)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '' ).
    range_a1->set_formula2( value = `RIGHT("hello",)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '' ).
  ENDMETHOD.

  METHOD if.
    range_a1->set_formula2( value = `IF(0=1,2,4)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 4 ).
  ENDMETHOD.

  METHOD iferror.
    range_a1->set_formula2( value = `IFERROR(#N/A,1)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 1 ).
    range_a1->set_formula2( value = `IFERROR(2,1)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD index.
    range_b2->set_value( zcl_xlom__va_string=>create( `Hello` ) ).
    range_a1->set_formula2( value = `INDEX(A1:C3,2,2)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = `Hello` ).
  ENDMETHOD.

  METHOD index_ae.
    range_a1->set_value( zcl_xlom__va_string=>create( `Hello ` ) ).
    range_b2->set_value( zcl_xlom__va_string=>create( `world` ) ).
    range_c1->set_formula2( value = `INDEX(A1:B2,{2,1},{2,1})` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_c1->value( ) )->get_string( )
                                        exp = `world` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_d1->value( ) )->get_string( )
                                        exp = `Hello ` ).

    range_a1->set_formula2( value = `INDEX({"a","b";"c","d"},{2,1},{2,1})` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = `d` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_b1->value( ) )->get_string( )
                                        exp = `a` ).
  ENDMETHOD.

  METHOD indirect.
    range_a1->set_value( zcl_xlom__va_string=>create( `Hello` ) ).
    range_a2->set_formula2( value = `INDIRECT("A1")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = `Hello` ).
  ENDMETHOD.

  METHOD len.
    range_a1->set_formula2( value = `LEN("ABC")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 3 ).
    range_a1->set_formula2( value = `LEN("ABC ")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 4 ).
    range_a1->set_formula2( value = `LEN("")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 0 ).
  ENDMETHOD.

  METHOD len_a1_a2.
    range_a1->set_value( zcl_xlom__va_string=>create( `Hello ` ) ).
    range_a2->set_value( zcl_xlom__va_string=>create( `world` ) ).
    range_b1->set_formula2( value = `LEN(A1:A2)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_b1->value( ) )->get_number( )
                                        exp = 6 ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_b2->value( ) )->get_number( )
                                        exp = 5 ).
  ENDMETHOD.

  METHOD match.
    range_a1->set_value( zcl_xlom__va_string=>create( `Hello ` ) ).
    range_a2->set_value( zcl_xlom__va_string=>create( `world` ) ).
    range_b1->set_formula2( value = `MATCH("world",A1:A2,0)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_b1->value( ) )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD match_2.
    range_a1->set_value( zcl_xlom__va_string=>create( `Hello ` ) ).
    range_a2->set_value( zcl_xlom__va_string=>create( `world` ) ).
    range_b1->set_formula2( value = `MATCH("world",A:A,0)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_b1->value( ) )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD minus.
    range_a1->set_formula2( value = `5-3` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD mult.
    range_a1->set_formula2( value = `2*3*4` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 24 ).
  ENDMETHOD.

  METHOD number.
    range_a1->set_formula2( value = `1` ).
    assert_equals( act = range_a1->value( )
                   exp = zcl_xlom__va_number=>get( 1 ) ).

    range_a1->set_formula2( value = `-1` ).
    assert_equals( act = range_a1->value( )
                   exp = zcl_xlom__va_number=>get( -1 ) ).
  ENDMETHOD.

  METHOD offset.
    DATA dummy_ref_to_offset TYPE REF TO zcl_xlom__ex_fu_offset ##NEEDED.
    DATA result              TYPE REF TO zif_xlom__va.

    result = application->evaluate( `OFFSET(A1,1,1)` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$B$2' ).

    result = application->evaluate( `OFFSET(A1,2,0)` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$A$3' ).

    result = application->evaluate( `OFFSET(A1,2,2)` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$C$3' ).

    result = application->evaluate( `OFFSET(C2,-1,-2)` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$A$1' ).

    result = application->evaluate( `OFFSET(A1,1,1,,)` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$B$2' ).

    result = application->evaluate( `OFFSET(A1,1,1,2,2)` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$B$2:$C$3' ).

    result = application->evaluate( `OFFSET(1:1,1,0)` ).
    cl_abap_unit_assert=>assert_equals( act = CAST zcl_xlom_range( result )->address( )
                                        exp = '$2:$2' ).
  ENDMETHOD.

  METHOD plus.
    range_a1->set_formula2( value = `1+1` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD range_a1_plus_one.
    range_a1->set_value( zcl_xlom__va_number=>create( 10 ) ).
    DATA(range_a2) = worksheet->range_from_address( 'A2' ).
    range_a2->set_formula2( 'A1+1' ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_number( range_a2->value( ) )->get_number( )
                                        exp = 11 ).
  ENDMETHOD.

  METHOD range_two_sheets.
    DATA dummy_ref_to_offset TYPE REF TO zcl_xlom__ex_fu_offset ##NEEDED.

    range_a1->set_value( zcl_xlom__va_string=>create( `Hello` ) ).

    DATA(worksheet_2) = workbook->worksheets->add( 'Sheet2' ).
    DATA(range_sheet2_b2) = worksheet_2->range_from_address( 'B2' ).
    range_sheet2_b2->set_formula2( |"C"&Sheet1!A1| ).

    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_sheet2_b2->value( ) )->get_string( )
                                        exp = `CHello` ).
  ENDMETHOD.

  METHOD right.
    range_a1->set_formula2( value = `RIGHT("Hello",2)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = 'lo' ).
    range_a1->set_formula2( value = `RIGHT(25,1)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '5' ).
    range_a1->set_formula2( value = `RIGHT("hello")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = 'o' ).
  ENDMETHOD.

  METHOD right_2.
    range_a1->set_formula2( value = `RIGHT("hello")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = 'o' ).
    range_a1->set_formula2( value = `RIGHT("hello",0)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '' ).
    range_a1->set_formula2( value = `RIGHT("hello",)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '' ).
  ENDMETHOD.

  METHOD row.
    range_a1->set_formula2( value = `ROW(B2)` ).
    assert_equals( act = range_a1->value( )
                   exp = zcl_xlom__va_number=>create( 2 ) ).
    range_a1->set_formula2( value = `ROW()` ).
    assert_equals( act = range_a1->value( )
                   exp = zcl_xlom__va_number=>create( 1 ) ).
  ENDMETHOD.

  METHOD setup.
    application = zcl_xlom_application=>create( ).
    workbook = application->workbooks->add( ).
    TRY.
        worksheet = workbook->worksheets->item( 'Sheet1' ).
        range_a1 = worksheet->range_from_address( 'A1' ).
        range_a2 = worksheet->range_from_address( 'A2' ).
        range_b1 = worksheet->range_from_address( 'B1' ).
        range_b2 = worksheet->range_from_address( 'B2' ).
        range_c1 = worksheet->range_from_address( 'C1' ).
        range_d1 = worksheet->range_from_address( 'D1' ).
      CATCH zcx_xlom__va INTO DATA(error). " TODO: variable is assigned but never used (ABAP cleaner)
        cl_abap_unit_assert=>fail( 'unexpected' ).
    ENDTRY.
  ENDMETHOD.

  METHOD string.
    range_a1->set_formula2( value = `"1"` ).
    assert_equals( act = range_a1->value( )
                   exp = zcl_xlom__va_string=>create( '1' ) ).
  ENDMETHOD.

  METHOD t.
    range_a1->set_formula2( value = `T("1")` ).
    assert_equals( act = range_a1->value( )
                   exp = zcl_xlom__va_string=>create( '1' ) ).

    range_a1->set_formula2( value = `T(1)` ).
    assert_equals( act = range_a1->value( )
                   exp = zcl_xlom__va_string=>create( '' ) ).
  ENDMETHOD.
ENDCLASS.
