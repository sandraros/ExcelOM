CLASS zcl_xlom_application DEFINITION
  PUBLIC
  CREATE PUBLIC
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    DATA active_sheet    TYPE REF TO zcl_xlom_sheet        READ-ONLY.
    DATA calculation     TYPE zcl_xlom=>ty_calculation     READ-ONLY VALUE zcl_xlom=>c_calculation-automatic ##NO_TEXT.
    DATA reference_style TYPE zcl_xlom=>ty_reference_style READ-ONLY VALUE zcl_xlom=>c_reference_style-a1 ##NO_TEXT.
    DATA workbooks       TYPE REF TO zcl_xlom_workbooks    READ-ONLY.

    METHODS calculate.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_application.

    METHODS evaluate
      IMPORTING !name         TYPE csequence
      RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

    METHODS international
      IMPORTING !index        TYPE zcl_xlom=>ty_application_international
      RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

    METHODS intersect
      IMPORTING arg1          TYPE REF TO zcl_xlom_range
                arg2          TYPE REF TO zcl_xlom_range
                arg3          TYPE REF TO zcl_xlom_range OPTIONAL
                arg4          TYPE REF TO zcl_xlom_range OPTIONAL
                arg5          TYPE REF TO zcl_xlom_range OPTIONAL
                arg6          TYPE REF TO zcl_xlom_range OPTIONAL
                arg7          TYPE REF TO zcl_xlom_range OPTIONAL
                arg8          TYPE REF TO zcl_xlom_range OPTIONAL
                arg9          TYPE REF TO zcl_xlom_range OPTIONAL
                arg10         TYPE REF TO zcl_xlom_range OPTIONAL
                arg11         TYPE REF TO zcl_xlom_range OPTIONAL
                arg12         TYPE REF TO zcl_xlom_range OPTIONAL
                arg13         TYPE REF TO zcl_xlom_range OPTIONAL
                arg14         TYPE REF TO zcl_xlom_range OPTIONAL
                arg15         TYPE REF TO zcl_xlom_range OPTIONAL
                arg16         TYPE REF TO zcl_xlom_range OPTIONAL
                arg17         TYPE REF TO zcl_xlom_range OPTIONAL
                arg18         TYPE REF TO zcl_xlom_range OPTIONAL
                arg19         TYPE REF TO zcl_xlom_range OPTIONAL
                arg20         TYPE REF TO zcl_xlom_range OPTIONAL
                arg21         TYPE REF TO zcl_xlom_range OPTIONAL
                arg22         TYPE REF TO zcl_xlom_range OPTIONAL
                arg23         TYPE REF TO zcl_xlom_range OPTIONAL
                arg24         TYPE REF TO zcl_xlom_range OPTIONAL
                arg25         TYPE REF TO zcl_xlom_range OPTIONAL
                arg26         TYPE REF TO zcl_xlom_range OPTIONAL
                arg27         TYPE REF TO zcl_xlom_range OPTIONAL
                arg28         TYPE REF TO zcl_xlom_range OPTIONAL
                arg29         TYPE REF TO zcl_xlom_range OPTIONAL
                arg30         TYPE REF TO zcl_xlom_range OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

    METHODS set_calculation
      IMPORTING !value TYPE zcl_xlom=>ty_calculation DEFAULT zcl_xlom=>c_calculation-automatic.

  PROTECTED SECTION.

  PRIVATE SECTION.
    DATA _country_code TYPE i VALUE 1 ##NO_TEXT.

    CLASS-METHODS _intersect_2
      IMPORTING arg1          TYPE REF TO zcl_xlom_range
                arg2          TYPE REF TO zcl_xlom_range
      RETURNING VALUE(result) TYPE zcl_xlom=>ts_range_address.

    CLASS-METHODS _intersect_2_basis
      IMPORTING arg1          TYPE zcl_xlom=>ts_range_address
                arg2          TYPE zcl_xlom=>ts_range_address
      RETURNING VALUE(result) TYPE zcl_xlom=>ts_range_address.

    CLASS-METHODS type
      IMPORTING any_data_object TYPE any
      RETURNING VALUE(result)   TYPE abap_typekind.
ENDCLASS.


CLASS zcl_xlom_application IMPLEMENTATION.
  METHOD calculate.
    DATA(workbook_number) = 1.
    WHILE workbook_number <= workbooks->count.
      DATA(workbook) = workbooks->item( workbook_number ).

      DATA(worksheet_number) = 1.
      WHILE worksheet_number <= workbook->worksheets->count.
        TRY.
            DATA(worksheet) = workbook->worksheets->item( worksheet_number ).
          CATCH zcx_xlom__va INTO DATA(error). " TODO: variable is assigned but never used (ABAP cleaner)
            RAISE EXCEPTION TYPE zcx_xlom_unexpected.
        ENDTRY.
        worksheet->calculate( ).
        worksheet_number = worksheet_number + 1.
      ENDWHILE.

      workbook_number = workbook_number + 1.
    ENDWHILE.
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom_application( ).
    result->workbooks = zcl_xlom_workbooks=>create( result ).
  ENDMETHOD.

  METHOD evaluate.
    DATA(lexer) = zcl_xlom__ex_ut_lexer=>create( ).
    DATA(lexer_tokens) = lexer->lexe( name ).
    DATA(parser) = zcl_xlom__ex_ut_parser=>create( ).

    TRY.
        DATA(expression) = parser->parse( lexer_tokens ).
      CATCH zcx_xlom__ex_ut_parser.
        result = zcl_xlom__va_error=>value_cannot_be_calculated.
        RETURN.
    ENDTRY.

    result = expression->evaluate(
                 context = zcl_xlom__ex_ut_eval_context=>create( worksheet       = CAST #( active_sheet )
                                                                 containing_cell = VALUE #( row    = 1
                                                                                            column = 1 ) ) ).
  ENDMETHOD.

  METHOD international.
    CASE index.
      WHEN zcl_xlom=>c_application_international-country_code.
        result = zcl_xlom__va_number=>get( EXACT #( _country_code ) ).
    ENDCASE.
  ENDMETHOD.

  METHOD intersect.
    TYPES tt_range TYPE STANDARD TABLE OF REF TO zcl_xlom_range WITH EMPTY KEY.

    DATA(args) = VALUE tt_range( ( arg1 )
                                 ( arg2 )
                                 ( arg3 )
                                 ( arg4 )
                                 ( arg5 )
                                 ( arg6 )
                                 ( arg7 )
                                 ( arg8 )
                                 ( arg9 )
                                 ( arg10 )
                                 ( arg11 )
                                 ( arg12 )
                                 ( arg13 )
                                 ( arg14 )
                                 ( arg15 )
                                 ( arg16 )
                                 ( arg17 )
                                 ( arg18 )
                                 ( arg19 )
                                 ( arg20 )
                                 ( arg21 )
                                 ( arg22 )
                                 ( arg23 )
                                 ( arg24 )
                                 ( arg25 )
                                 ( arg26 )
                                 ( arg27 )
                                 ( arg28 )
                                 ( arg29 )
                                 ( arg30 ) ).
*    DATA(args) = VALUE tt_range( ( arg1 ) ( arg2 ) ( arg3 ) ( arg4 ) ( arg5 ) ( arg6 ) ( arg7 ) ( arg8 ) ( arg9 ) ( arg10 )
*                                 ( arg11 ) ( arg12 ) ( arg13 ) ( arg14 ) ( arg15 ) ( arg16 ) ( arg17 ) ( arg18 ) ( arg19 ) ( arg20 )
*                                 ( arg21 ) ( arg22 ) ( arg23 ) ( arg24 ) ( arg25 ) ( arg26 ) ( arg27 ) ( arg28 ) ( arg29 ) ( arg30 ) ).

    DATA(temp_intersect_range_address) = VALUE zcl_xlom=>ts_range_address( ).
    LOOP AT args INTO DATA(arg)
         WHERE table_line IS BOUND.
      temp_intersect_range_address = _intersect_2_basis(
                                         arg1 = temp_intersect_range_address
                                         arg2 = VALUE #( top_left-column     = arg->_address-top_left-column
                                                         top_left-row        = arg->_address-top_left-row
                                                         bottom_right-column = arg->_address-bottom_right-column
                                                         bottom_right-row    = arg->_address-bottom_right-row ) ).
      IF temp_intersect_range_address IS INITIAL.
        " Empty intersection
        RETURN.
      ENDIF.
    ENDLOOP.

    result = zcl_xlom_range=>create_from_top_left_bottom_ri(
                 worksheet    = arg1->parent
                 top_left     = VALUE #( column = temp_intersect_range_address-top_left-column
                                         row    = temp_intersect_range_address-top_left-row )
                 bottom_right = VALUE #( column = temp_intersect_range_address-bottom_right-column
                                         row    = temp_intersect_range_address-bottom_right-row ) ).
  ENDMETHOD.

  METHOD set_calculation.
    calculation = value.
  ENDMETHOD.

  METHOD type.
    DESCRIBE FIELD any_data_object TYPE result.
  ENDMETHOD.

  METHOD _intersect_2.
    TYPES tt_range TYPE STANDARD TABLE OF REF TO zcl_xlom_range WITH EMPTY KEY.

    DATA(args) = VALUE tt_range( ( arg1 ) ( arg2 ) ).

    LOOP AT args INTO DATA(arg)
         WHERE table_line IS BOUND.
      result = _intersect_2_basis( arg1 = result
                                   arg2 = VALUE #( top_left-column     = arg->_address-top_left-column
                                                   top_left-row        = arg->_address-top_left-row
                                                   bottom_right-column = arg->_address-bottom_right-column
                                                   bottom_right-row    = arg->_address-bottom_right-row ) ).
      IF result IS INITIAL.
        " Empty intersection
        RETURN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD _intersect_2_basis.
    result = COND #( WHEN arg1 IS NOT INITIAL
                     THEN arg1
                     ELSE VALUE #( top_left-column     = 0
                                   top_left-row        = 0
                                   bottom_right-column = zcl_xlom_worksheet=>max_columns + 1
                                   bottom_right-row    = zcl_xlom_worksheet=>max_rows + 1 ) ).

    IF arg2-top_left-column > result-top_left-column.
      result-top_left-column = arg2-top_left-column.
    ENDIF.
    IF arg2-top_left-row > result-top_left-row.
      result-top_left-row = arg2-top_left-row.
    ENDIF.
    IF arg2-bottom_right-column < result-bottom_right-column.
      result-bottom_right-column = arg2-bottom_right-column.
    ENDIF.
    IF arg2-bottom_right-row < result-bottom_right-row.
      result-bottom_right-row = arg2-bottom_right-row.
    ENDIF.

    IF    result-top_left-column > result-bottom_right-column
       OR result-top_left-row    > result-bottom_right-row.
      " Empty intersection
      result = VALUE #( ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.
