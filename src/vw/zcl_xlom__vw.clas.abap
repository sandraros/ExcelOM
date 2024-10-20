CLASS zcl_xlom__vw DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    METHODS pai.

    METHODS pbo.

    CLASS-METHODS view
      IMPORTING !application TYPE REF TO zcl_xlom_application.

  PRIVATE SECTION.
    DATA application                    TYPE REF TO zcl_xlom_application.

    DATA pbo_already_executed           TYPE abap_bool VALUE abap_false.
    DATA worksheet_horizontal_toolbar_c TYPE REF TO cl_gui_custom_container.
    DATA worksheet_horizontal_toolbar   TYPE REF TO cl_gui_toolbar.
    DATA main_worksheet_alv_container   TYPE REF TO cl_gui_custom_container.
    DATA main_worksheet_alv             TYPE REF TO cl_gui_alv_grid.
    DATA ref_main_worksheet_alv_outtab  TYPE REF TO data.

    CLASS-METHODS create
      IMPORTING !application  TYPE REF TO zcl_xlom_application
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__vw.
ENDCLASS.


CLASS zcl_xlom__vw IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__vw( ).
    result->application = application.
  ENDMETHOD.

  METHOD pai.

  ENDMETHOD.

  METHOD pbo.
    IF pbo_already_executed = abap_true.
      RETURN.
    ENDIF.

    worksheet_horizontal_toolbar_c = NEW cl_gui_custom_container( container_name = 'WORKSHEET_HORIZONTAL_TOOLBAR' ).

    worksheet_horizontal_toolbar   = NEW cl_gui_toolbar( parent = worksheet_horizontal_toolbar_c ).
    worksheet_horizontal_toolbar->add_button( fcode     = 'A1'
                                              icon      = ''
                                              butn_type = cntb_btype_group
                                              is_checked = abap_true
                                              text      = 'BNKA' ).
    worksheet_horizontal_toolbar->add_button( fcode     = 'A2'
                                              icon      = ''
                                              butn_type = cntb_btype_group
                                              text      = 'BNKAIN' ).
    worksheet_horizontal_toolbar->add_button( fcode     = 'A3'
                                              icon      = ''
                                              butn_type = cntb_btype_group
                                              text      = 'ADRC' ).

    main_worksheet_alv_container = NEW cl_gui_custom_container( container_name = 'MAIN_WORKSHEET_ALV' ).

    main_worksheet_alv = NEW cl_gui_alv_grid( i_parent = main_worksheet_alv_container ).

    DATA(workbook) = application->workbooks->item( 1 ).

    DATA(worksheet) = workbook->worksheets->item( 1 ).

    DATA(used_range) = worksheet->used_range( ).
    DATA(row_count) = used_range->rows( )->count( ).
    DATA(column_count) = used_range->columns( )->count( ).

    DATA(rtts_structure) = cl_abap_structdescr=>get(
                               p_components = VALUE #( FOR aux_column = 1 WHILE aux_column <= column_count
                                                       ( name = |COMP_{ aux_column }|
                                                         type = cl_abap_elemdescr=>get_string( ) ) ) ).

    DATA(rtts_table) = cl_abap_tabledescr=>get( p_line_type = rtts_structure ).
    TYPES ty_ref_to_data TYPE REF TO data.
    FIELD-SYMBOLS <alv_outtab> TYPE STANDARD TABLE.
    FIELD-SYMBOLS <structure>  TYPE any.

    DATA(ref_to_data) = VALUE ty_ref_to_data( ).

    CREATE DATA ref_to_data TYPE HANDLE rtts_structure.
    ASSIGN ref_to_data->* TO <structure>.

    CREATE DATA ref_main_worksheet_alv_outtab TYPE HANDLE rtts_table.
    ASSIGN ref_main_worksheet_alv_outtab->* TO <alv_outtab>.

    DATA(row) = 1.
    WHILE row <= row_count.
      CLEAR <structure>.
      LOOP AT rtts_structure->components REFERENCE INTO DATA(component).
        DATA(column_2) = sy-tabix.
        ASSIGN COMPONENT component->name OF STRUCTURE <structure> TO FIELD-SYMBOL(<field>).
        DATA(xlom_cell) = worksheet->cells( row    = row
                                            column = column_2 )->value( ).
        CASE xlom_cell->type.
          WHEN xlom_cell->c_type-string.
            <field> = CAST zcl_xlom__va_string( xlom_cell )->get_string( ).
          WHEN xlom_cell->c_type-number.
            <field> = CAST zcl_xlom__va_number( xlom_cell )->get_number( ).
        ENDCASE.
      ENDLOOP.
      INSERT <structure> INTO TABLE <alv_outtab>.
      row = row + 1.
    ENDWHILE.

    DATA(field_catalog) = VALUE lvc_t_fcat(
                                    FOR <component> IN rtts_structure->components INDEX INTO component_number
                                    ( fieldname  = |COMP_{ component_number }|
                                      colddictxt = 'L'
                                      scrtext_l  = zcl_xlom_range=>convert_column_number_to_a_xfd( component_number ) ) ).
    main_worksheet_alv->set_table_for_first_display(
      EXPORTING  is_layout                     = VALUE #( cwidth_opt = 'X'
                                                          grid_title = EXACT #( worksheet->name )
                                                          no_toolbar = abap_true )
      CHANGING   it_outtab                     = <alv_outtab>
                 it_fieldcatalog               = field_catalog
      EXCEPTIONS invalid_parameter_combination = 1                " Wrong Parameter
                 program_error                 = 2                " Program Errors
                 too_many_lines                = 3                " Too many Rows in Ready for Input Grid
                 OTHERS                        = 4 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    pbo_already_executed = abap_true.
  ENDMETHOD.

  METHOD view.
    DATA(data_viewer) = create( application = application ).
    CALL FUNCTION 'Z_XLOM_UI_DATA_VIEWER'
      EXPORTING data_viewer = data_viewer.
  ENDMETHOD.
ENDCLASS.
