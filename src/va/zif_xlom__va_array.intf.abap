INTERFACE zif_xlom__va_array
  PUBLIC.

  INTERFACES zif_xlom__va.

  TYPES tt_column TYPE STANDARD TABLE OF REF TO zif_xlom__va WITH EMPTY KEY.
  TYPES:
    BEGIN OF ts_row,
      columns_of_row TYPE tt_column,
    END OF ts_row.
  TYPES tt_row TYPE STANDARD TABLE OF ts_row WITH EMPTY KEY.
  TYPES:
    BEGIN OF ts_address_one_cell,
      "! 0 means that the address is the whole row defined in ROW
      column       TYPE i,
      column_fixed TYPE abap_bool,
      "! 0 means that the address is the whole column defined in COLUMN
      row          TYPE i,
      row_fixed    TYPE abap_bool,
    END OF ts_address_one_cell.
  TYPES:
    BEGIN OF ts_address,
      "! Can also be an internal ID like "1" ([1]Sheet1!A1)
      workbook_name  TYPE string,
      worksheet_name TYPE string,
      range_name     TYPE string,
      top_left       TYPE ts_address_one_cell,
      bottom_right   TYPE ts_address_one_cell,
    END OF ts_address.

  DATA row_count    TYPE i READ-ONLY.
  DATA column_count TYPE i READ-ONLY.

  METHODS get_array_value
    IMPORTING top_left      TYPE zcl_xlom=>ts_range_address_one_cell
              bottom_right  TYPE zcl_xlom=>ts_range_address_one_cell
    RETURNING VALUE(result) TYPE REF TO zif_xlom__va_array.

  "! @parameter column | Start from 1
  "! @parameter row    | Start from 1
  METHODS get_cell_value
    IMPORTING !column       TYPE i
              !row          TYPE i
    RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

  METHODS set_array_value
    IMPORTING !rows TYPE tt_row.

  METHODS set_cell_value
    IMPORTING !column    TYPE i
              !row       TYPE i
              !value     TYPE REF TO zif_xlom__va
              formula    TYPE REF TO zif_xlom__ex OPTIONAL
              calculated TYPE abap_bool           OPTIONAL.
ENDINTERFACE.
