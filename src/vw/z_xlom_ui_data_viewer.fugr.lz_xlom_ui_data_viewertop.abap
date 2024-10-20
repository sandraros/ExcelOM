FUNCTION-POOL z_xlom_ui_data_viewer.        " MESSAGE-ID ..

" INCLUDE LZ_XLOM_UI_DATA_VIEWERD...         " Local class definition

DATA:
  BEGIN OF global_data,
    data_viewer                         TYPE REF TO zcl_xlom__vw,
*    "! container
*    worksheet_horizontal_toolbar_c TYPE REF TO cl_gui_custom_container,
*    worksheet_horizontal_toolbar   TYPE REF TO cl_gui_toolbar,
*    main_worksheet_alv_container   TYPE REF TO cl_gui_custom_container,
*    main_worksheet_alv             TYPE REF TO cl_gui_alv_grid,
*    ref_main_worksheet_alv_outtab  TYPE REF TO data,
  END OF global_data.
