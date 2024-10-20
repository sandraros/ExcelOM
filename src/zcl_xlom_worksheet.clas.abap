"! https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet
CLASS zcl_xlom_worksheet DEFINITION
  PUBLIC
  INHERITING FROM zcl_xlom_sheet FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    TYPES ty_name TYPE c LENGTH 31.

    DATA application TYPE REF TO zcl_xlom_application READ-ONLY.
    "! worksheet name
    DATA name        TYPE ty_name                     READ-ONLY.
    DATA parent      TYPE REF TO zcl_xlom_workbook    READ-ONLY.

    "! Worksheet.Calculate method (Excel).
    "! Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.
    "! <p>expression.Calculate</p>
    "! expression A variable that represents a Worksheet object.
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.calculate(method)
    METHODS calculate.

    "! Use either both row and column, or item alone.
    "! @parameter row | Start from 1
    "! @parameter column | Start from 1.
    "! @parameter item | Item number from 1, 16385 is the same as row = 2 column = 1.
    METHODS cells
      IMPORTING !row          TYPE i    OPTIONAL
                !column       TYPE i    OPTIONAL
                item          TYPE int8 OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    CLASS-METHODS create
      IMPORTING workbook      TYPE REF TO zcl_xlom_workbook
                !name         TYPE csequence
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_worksheet.

    METHODS range
      IMPORTING cell1_string  TYPE string                OPTIONAL
                cell2_string  TYPE string                OPTIONAL
                cell1_range   TYPE REF TO zcl_xlom_range OPTIONAL
                cell2_range   TYPE REF TO zcl_xlom_range OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range
      RAISING   zcx_xlom__va.

    METHODS used_range
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

  PROTECTED SECTION.

  PRIVATE SECTION.
    CONSTANTS max_rows    TYPE i VALUE 1048576 ##NO_TEXT.
    CONSTANTS max_columns TYPE i VALUE 16384 ##NO_TEXT.

    DATA _array TYPE REF TO zcl_xlom__va_array.

    "! Worksheet.Range property. Returns a Range object that represents a cell or a range of cells.
    "! <p>expression.Range (Cell1, Cell2)</p>
    "! expression A variable that represents a Worksheet object.
    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    METHODS range_from_address
      IMPORTING cell1         TYPE string
                cell2         TYPE string OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range
      RAISING   zcx_xlom__va.

    "! Worksheet.Range property. Returns a Range object that represents a cell or a range of cells.
    "! <p>expression.Range (Cell1, Cell2)</p>
    "! expression A variable that represents a Worksheet object.
    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    METHODS range_from_two_ranges
      IMPORTING cell1         TYPE REF TO zcl_xlom_range
                cell2         TYPE REF TO zcl_xlom_range
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.
ENDCLASS.


CLASS zcl_xlom_worksheet IMPLEMENTATION.
  METHOD calculate.
*    RAISE EXCEPTION TYPE ZCX_xlom_to_do.
    " TODO the containing cell shouldn't be A1, it should vary for
    "      each cell where the formula is to be calculated.
    "      Solution: maybe store the range object within each cell
    "                to not recalculate it (performance).
    DATA(context) = zcl_xlom__ex_ut_eval_context=>create( worksheet       = me
                                                          containing_cell = VALUE #( row    = 1
                                                                                     column = 1 ) ).
    LOOP AT _array->_cells REFERENCE INTO DATA(cell)
         WHERE     formula    IS BOUND
               AND calculated  = abap_false.
      context->set_containing_cell( VALUE #( row    = cell->row
                                             column = cell->column ) ).
      cell->value = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                      expression = cell->formula
                      context    = context ).
*      cell->formula->evaluate( context ).
    ENDLOOP.
  ENDMETHOD.

  METHOD cells.
    " TODO: parameter ITEM is never used (ABAP cleaner)

    " This will change Z20:
    " Range("Z20:AA25").Cells(1, 1) = "C"
    result = zcl_xlom_range=>create_from_row_column( worksheet = me
                                                     row       = row
                                                     column    = column ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom_worksheet( ).
    result->name        = name.
    result->parent      = workbook.
    result->application = workbook->application.
    result->_array      = zcl_xlom__va_array=>create_initial( row_count    = max_rows
                                                              column_count = max_columns ).
  ENDMETHOD.

  METHOD range.
    IF    (     cell1_string IS NOT INITIAL
            AND cell1_range  IS BOUND )
       OR (     cell1_string IS INITIAL
            AND cell1_range  IS NOT BOUND )
       OR (     cell1_string IS INITIAL
            AND cell2_string IS NOT INITIAL )
       OR (     cell1_range  IS NOT BOUND
            AND cell2_range  IS BOUND ).
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.

    IF cell1_string IS NOT INITIAL.
      result = range_from_address( cell1 = cell1_string
                                   cell2 = cell2_string ).
    ELSE.
      result = range_from_two_ranges( cell1 = cell1_range
                                      cell2 = cell2_range ).
    ENDIF.
  ENDMETHOD.

  METHOD range_from_address.
    DATA(range_1) = zcl_xlom_range=>create_from_address_or_name( address     = cell1
                                                                 relative_to = me ).
    IF cell2 IS INITIAL.
      result = range_1.
    ELSE.
      DATA(range_2) = zcl_xlom_range=>create_from_address_or_name( address     = cell2
                                                                   relative_to = me ).
      result = range_from_two_ranges( cell1 = range_1
                                      cell2 = range_2 ).
    ENDIF.
  ENDMETHOD.

  METHOD range_from_two_ranges.
    result = zcl_xlom_range=>create( cell1 = cell1
                                     cell2 = cell2 ).
  ENDMETHOD.

  METHOD used_range.
    result = zcl_xlom_range=>create_from_row_column(
                 worksheet   = me
                 row         = _array->used_range-top_left-row
                 column      = _array->used_range-top_left-column
                 row_size    = _array->used_range-bottom_right-row - _array->used_range-top_left-row + 1
                 column_size = _array->used_range-bottom_right-column - _array->used_range-top_left-column + 1 ).
  ENDMETHOD.
ENDCLASS.
