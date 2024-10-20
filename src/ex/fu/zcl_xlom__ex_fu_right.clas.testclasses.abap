*"* use this source file for your ABAP unit test classes

CLASS ltc_app DEFINITION
  INHERITING FROM zcl_xlom__ex_ut_eval_aunit FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS test FOR TESTING RAISING cx_static_check.
ENDCLASS.

CLASS ltc_app IMPLEMENTATION.
  METHOD test.
    setup_default_xlom_objects( ).

    range_a1->set_formula2( value = `RIGHT("Hello",2)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = 'lo' ).

    range_a1->set_formula2( value = `RIGHT(25,1)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '5' ).

    range_a1->set_formula2( value = `RIGHT("hello")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = 'o' ).

    range_a1->set_formula2( value = `RIGHT("hello")` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = 'o' ).

    range_a1->set_formula2( value = `RIGHT("hello",0)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '' ).

    " RIGHT("hello",) is the same result as RIGHT("hello",0), it differs from RIGHT("hello"), which is the same as RIGHT("hello",1).
    range_a1->set_formula2( value = `RIGHT("hello",)` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
                                        exp = '' ).
  ENDMETHOD.
ENDCLASS.
