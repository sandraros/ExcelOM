*"* use this source file for your ABAP unit test classes

CLASS ltc_range DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

  PRIVATE SECTION.
    METHODS convert_column_a_xfd_to_number FOR TESTING RAISING cx_static_check.
*    METHODS decode_range_address_a1_invali FOR TESTING RAISING cx_static_check.
    METHODS decode_range_address_a1_valid  FOR TESTING RAISING cx_static_check.
*    METHODS decode_range_address_sh_invali FOR TESTING RAISING cx_static_check.
    METHODS decode_range_address_sh_valid  FOR TESTING RAISING cx_static_check.
    METHODS convert_column_number_to_a_xfd FOR TESTING RAISING cx_static_check.

    TYPES ty_address TYPE zif_xlom__va_array=>ts_address.
ENDCLASS.


CLASS ltc_range IMPLEMENTATION.
  METHOD convert_column_a_xfd_to_number.
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = 'XFD' )
                                        exp = 16384 ).

    TRY.
        zcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = 'XFE' ).
        cl_abap_unit_assert=>fail( msg = 'Exception expected for XFE - Column does not exist' ).
      CATCH cx_root ##NO_HANDLER.
    ENDTRY.

    TRY.
        zcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = 'ZZZZ' ).
        cl_abap_unit_assert=>fail( msg = 'Exception expected for ZZZZ - Column does not exist' ).
      CATCH cx_root ##NO_HANDLER.
    ENDTRY.

    TRY.
        zcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = '1' ).
        cl_abap_unit_assert=>fail( msg = 'Exception expected for 1 - Invalid column ID' ).
      CATCH cx_root ##NO_HANDLER.
    ENDTRY.
  ENDMETHOD.

  METHOD convert_column_number_to_a_xfd.
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>convert_column_number_to_a_xfd( 16384 )
                                        exp = 'XFD' ).

    TRY.
        zcl_xlom_range=>convert_column_number_to_a_xfd( 16385 ).
        cl_abap_unit_assert=>fail( msg = 'Exception expected for 16385 - Column does not exist' ).
      CATCH cx_root ##NO_HANDLER.
    ENDTRY.

    TRY.
        zcl_xlom_range=>convert_column_number_to_a_xfd( -1 ).
        cl_abap_unit_assert=>fail( msg = 'Exception expected for -1 - Column does not exist' ).
      CATCH cx_root ##NO_HANDLER.
    ENDTRY.
  ENDMETHOD.

*  METHOD decode_range_address_a1_invali.
*    LOOP AT VALUE string_table( ( `:` ) ( `` ) ( `$` ) ( `A` ) ( `A:` ) ( `$$A1` ) ( `A:A1` ) ( `B2:A1` ) ) INTO DATA(address).
*      TRY.
*          zcl_xlom_range=>decode_range_address_a1( address ).
*          cl_abap_unit_assert=>fail( msg = |Exception expected for address "{ address }"| ).
*        CATCH cx_root ##NO_HANDLER.
*      ENDTRY.
*    ENDLOOP.
*  ENDMETHOD.

  METHOD decode_range_address_a1_valid.
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( 'A1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1
                                                                                        row    = 1 )
                                                                bottom_right = VALUE #( column = 1
                                                                                        row    = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( 'A$1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column    = 1
                                                                                        row       = 1
                                                                                        row_fixed = abap_true )
                                                                bottom_right = VALUE #( column    = 1
                                                                                        row       = 1
                                                                                        row_fixed = abap_true ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( '$A1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1 )
                                                                bottom_right = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( '$A$1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1
                                                                                        row_fixed    = abap_true )
                                                                bottom_right = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1
                                                                                        row_fixed    = abap_true ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( 'A1:B1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1
                                                                                        row    = 1 )
                                                                bottom_right = VALUE #( column = 2
                                                                                        row    = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( 'A:A' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1 )
                                                                bottom_right = VALUE #( column = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( '1:1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( row = 1 )
                                                                bottom_right = VALUE #( row = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( 'Sheet1!A1' )
                                        exp = VALUE ty_address( worksheet_name = 'Sheet1'
                                                                top_left       = VALUE #( column = 1
                                                                                          row    = 1 )
                                                                bottom_right   = VALUE #( column = 1
                                                                                          row    = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( `'Sheet1 (2)'!A1` )
                                        exp = VALUE ty_address( worksheet_name = 'Sheet1 (2)'
                                                                top_left       = VALUE #( column = 1
                                                                                          row    = 1 )
                                                                bottom_right   = VALUE #( column = 1
                                                                                          row    = 1 ) ) ).
  ENDMETHOD.

*  METHOD decode_range_address_sh_invali.
*    LOOP AT VALUE string_table( ( `:` ) ( `` ) ( `$` ) ( `A` ) ( `A:` ) ( `$$A1` ) ( `A:A1` ) ( `B2:A1` ) ) INTO DATA(address).
*      TRY.
*          zcl_xlom_range=>decode_range_address_a1( address ).
*          cl_abap_unit_assert=>fail( msg = |Exception expected for address "{ address }"| ).
*        CATCH cx_root ##NO_HANDLER.
*      ENDTRY.
*    ENDLOOP.
*  ENDMETHOD.

  METHOD decode_range_address_sh_valid.
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( 'BKPF!A:A' )
                                        exp = VALUE ty_address( worksheet_name = 'BKPF'
                                                                top_left       = VALUE #( column = 1 )
                                                                bottom_right   = VALUE #( column = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = zcl_xlom_range=>decode_range_address_a1( `'BKPF (2)'!A:A` )
                                        exp = VALUE ty_address( worksheet_name = 'BKPF (2)'
                                                                top_left       = VALUE #( column = 1 )
                                                                bottom_right   = VALUE #( column = 1 ) ) ).
  ENDMETHOD.
ENDCLASS.
