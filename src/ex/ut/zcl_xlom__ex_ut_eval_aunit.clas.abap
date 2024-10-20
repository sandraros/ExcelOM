CLASS zcl_xlom__ex_ut_eval_aunit DEFINITION
  PUBLIC
  CREATE PUBLIC
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PROTECTED SECTION.
    DATA application TYPE REF TO zcl_xlom_application.
    DATA workbook    TYPE REF TO zcl_xlom_workbook.
    DATA worksheet   TYPE REF TO zcl_xlom_worksheet.
    DATA range_a1    TYPE REF TO zcl_xlom_range.
    DATA range_a2    TYPE REF TO zcl_xlom_range.
    DATA range_b1    TYPE REF TO zcl_xlom_range.
    DATA range_b2    TYPE REF TO zcl_xlom_range.
    DATA range_c1    TYPE REF TO zcl_xlom_range.
    DATA range_d1    TYPE REF TO zcl_xlom_range.

    METHODS assert_equals
      IMPORTING act           TYPE REF TO zif_xlom__va
                exp           TYPE REF TO zif_xlom__va
      RETURNING VALUE(result) TYPE abap_bool.

    METHODS setup_default_xlom_objects.
ENDCLASS.


CLASS zcl_xlom__ex_ut_eval_aunit IMPLEMENTATION.
  METHOD assert_equals.
    cl_abap_unit_assert=>assert_true( xsdbool( exp->is_equal( act ) ) ).
  ENDMETHOD.

  METHOD setup_default_xlom_objects.
    application = zcl_xlom_application=>create( ).
    workbook = application->workbooks->add( ).
    TRY.
        worksheet = workbook->worksheets->item( 'Sheet1' ).
        range_a1 = worksheet->range( cell1_string = 'A1' ).
        range_a2 = worksheet->range( cell1_string = 'A2' ).
        range_b1 = worksheet->range( cell1_string = 'B1' ).
        range_b2 = worksheet->range( cell1_string = 'B2' ).
        range_c1 = worksheet->range( cell1_string = 'C1' ).
        range_d1 = worksheet->range( cell1_string = 'D1' ).
      CATCH zcx_xlom__va.
        cl_abap_unit_assert=>fail( 'unexpected' ).
    ENDTRY.
  ENDMETHOD.
ENDCLASS.
