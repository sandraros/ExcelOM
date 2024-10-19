CLASS zcl_xlom__va_array DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__va.
    INTERFACES zif_xlom__va_array.
    INTERFACES zif_xlom__ut_all_friends.

    TYPES:
      BEGIN OF ts_used_range_one_cell,
        column TYPE i,
        row    TYPE i,
      END OF ts_used_range_one_cell.
    TYPES:
      BEGIN OF ts_used_range,
        top_left     TYPE ts_used_range_one_cell,
        bottom_right TYPE ts_used_range_one_cell,
      END OF ts_used_range.

    DATA used_range TYPE ts_used_range READ-ONLY.

    CLASS-METHODS class_constructor.

    CLASS-METHODS create_from_range
      IMPORTING !range        TYPE REF TO zcl_xlom_range
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_array.

    CLASS-METHODS create_initial
      IMPORTING row_count     TYPE i
                column_count  TYPE i
                !rows         TYPE zif_xlom__va_array=>tt_row OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_array.

  PRIVATE SECTION.
    "! Internal type of cell value (empty, number, string, boolean, error, array, compound data)
    TYPES ty_value_type TYPE i.
    TYPES:
      BEGIN OF ts_cell,
        "! Start from 1
        column     TYPE i,
        "! Start from 1
        row        TYPE i,
        formula    TYPE REF TO zif_xlom__ex,
        calculated TYPE abap_bool,
        value      TYPE REF TO zif_xlom__va,
      END OF ts_cell.
    TYPES tt_cell TYPE SORTED TABLE OF ts_cell WITH UNIQUE KEY row column.

    CONSTANTS:
      BEGIN OF c_value_type,
        "! Needed by ISBLANK formula function. IT CANNOT be replaced with "empty = xsdbool( not line_exists( _cells[ row = ... column = ... ] ) )"
        "! because a cell may exist for data other than value, like number format, background color, and so on. Optionally, there could be two
        "! internal tables, _cells only for values.
        empty         TYPE ty_value_type VALUE 1,
        number        TYPE ty_value_type VALUE 2,
        string        TYPE ty_value_type VALUE 3,
        "! Cell containing the value TRUE or FALSE.
        "! Needed by TYPE formula function (4 = logical value)
        boolean       TYPE ty_value_type VALUE 4,
        "! Needed by TYPE formula function (16 = error)
        error         TYPE ty_value_type VALUE 5,
        "! Needed by TYPE formula function (64 = array)
        array         TYPE ty_value_type VALUE 6,
        "! Needed by TYPE formula function (128 = compound data)
        compound_data TYPE ty_value_type VALUE 7,
      END OF c_value_type.

    CLASS-DATA initial_used_range TYPE zcl_xlom__va_array=>ts_used_range.

    DATA _cells TYPE tt_cell.

    "! @parameter row | Start from 1
    "! @parameter column | Start from 1
    METHODS set_cell_value_single
      IMPORTING !row       TYPE i
                !column    TYPE i
                !value     TYPE REF TO zif_xlom__va
                formula    TYPE REF TO zif_xlom__ex OPTIONAL
                calculated TYPE abap_bool           OPTIONAL.
ENDCLASS.


CLASS zcl_xlom__va_array IMPLEMENTATION.
  METHOD class_constructor.
    initial_used_range = VALUE #( top_left     = VALUE #( row    = 1
                                                          column = 1 )
                                  bottom_right = VALUE #( row    = 1
                                                          column = 1 ) ).
  ENDMETHOD.

  METHOD create_from_range.
    result = range->parent->_array.
  ENDMETHOD.

  METHOD create_initial.
    result = NEW zcl_xlom__va_array( ).
    result->zif_xlom__va~type               = zif_xlom__va=>c_type-array.
    result->zif_xlom__va_array~row_count    = row_count.
    result->zif_xlom__va_array~column_count = column_count.
    result->used_range                      = initial_used_range.
    result->zif_xlom__va_array~set_array_value( rows ).
  ENDMETHOD.

  METHOD set_cell_value_single.
    DATA(cell) = REF #( _cells[ row    = row
                                column = column ] OPTIONAL ).
    IF cell IS NOT BOUND.
      INSERT VALUE #( row    = row
                      column = column )
             INTO TABLE _cells
             REFERENCE INTO cell.
    ENDIF.
    cell->value      = value.
    cell->formula    = formula.
    cell->calculated = calculated.

    IF lines( _cells ) = 1.
      used_range = VALUE #( top_left     = VALUE #( row    = row
                                                    column = column )
                            bottom_right = VALUE #( row    = row
                                                    column = column ) ).
    ELSE.
      IF row < used_range-top_left-row.
        used_range-top_left-row = row.
      ENDIF.
      IF column < used_range-top_left-column.
        used_range-top_left-column = column.
      ENDIF.
      IF row > used_range-bottom_right-row.
        used_range-bottom_right-row = row.
      ENDIF.
      IF column > used_range-bottom_right-column.
        used_range-bottom_right-column = column.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__va_array~get_array_value.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
    DATA(row_count)    = bottom_right-row - top_left-row + 1.
    DATA(column_count) = bottom_right-column - top_left-column + 1.
    DATA(target_array) = create_initial( row_count    = row_count
                                         column_count = column_count ).
    DATA(row) = 1.
    WHILE row <= row_count.
      DATA(column) = 1.
      WHILE column <= column_count.
        target_array->zif_xlom__va_array~set_cell_value(
            column = column
            row    = row
            value  = zif_xlom__va_array~get_cell_value( column = top_left-column + column - 1
                                                        row    = top_left-row + row - 1 ) ).
        column = column + 1.
      ENDWHILE.
      row = row + 1.
    ENDWHILE.
    result = target_array.
  ENDMETHOD.

  METHOD zif_xlom__va_array~get_cell_value.
    IF    row    < used_range-top_left-row
       OR row    > used_range-bottom_right-row
       OR column < used_range-top_left-column
       OR column > used_range-bottom_right-column.
      result = zcl_xlom__va_empty=>get_singleton( ).
    ELSE.
      DATA(cell) = REF #( _cells[ row    = row
                                  column = column ] OPTIONAL ).
      IF cell IS NOT BOUND.
        " Empty/Blank - Its evaluation depends on its usage (zero or empty string)
        " =1+Empty gives 1, ="a"&Empty gives "a"
        result = zcl_xlom__va_empty=>get_singleton( ).
      ELSE.
        result = cell->value.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__va_array~set_array_value.
    DATA(row) = 1.
    LOOP AT rows REFERENCE INTO DATA(row2).

      DATA(column) = 1.
      LOOP AT row2->columns_of_row INTO DATA(column_value).
        zif_xlom__va_array~set_cell_value( column = column
                                           row    = row
                                           value  = column_value ).
        column = column + 1.
      ENDLOOP.

      row = row + 1.
    ENDLOOP.
  ENDMETHOD.

  METHOD zif_xlom__va_array~set_cell_value.
    IF    row    > zif_xlom__va_array~row_count
       OR row    < 1
       OR column > zif_xlom__va_array~column_count
       OR column < 1.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    IF value IS NOT BOUND.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.

    CASE value->type.

      WHEN value->c_type-array
        OR value->c_type-range.

        DATA(source_array) = CAST zif_xlom__va_array( value ).
        DATA(source_array_row) = 1.
        WHILE source_array_row <= source_array->row_count.
          DATA(source_array_column) = 1.
          WHILE source_array_column <= source_array->column_count.
            DATA(source_array_cell) = source_array->get_cell_value( column = source_array_column
                                                                    row    = source_array_row ).
            set_cell_value_single( row        = row + source_array_row - 1
                                   column     = column + source_array_column - 1
                                   value      = source_array_cell
                                   formula    = formula
                                   calculated = calculated ).

            source_array_column = source_array_column + 1.
          ENDWHILE.
          source_array_row = source_array_row + 1.
        ENDWHILE.

      WHEN OTHERS.

        set_cell_value_single( row        = row
                               column     = column
                               value      = value
                               formula    = formula
                               calculated = calculated ).
    ENDCASE.
  ENDMETHOD.

  METHOD zif_xlom__va~get_value.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_array.
    result = abap_true.
  ENDMETHOD.

  METHOD zif_xlom__va~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_equal.
    IF input_result->type = zif_xlom__va=>c_type-array.
      DATA(input_array) = CAST zcl_xlom__va_array( input_result ).
      IF     zif_xlom__va_array~column_count = input_array->zif_xlom__va_array~column_count
         AND zif_xlom__va_array~row_count    = input_array->zif_xlom__va_array~row_count
         AND me->_cells = input_array->_cells.
        result = abap_true.
      ELSE.
        result = abap_false.
      ENDIF.
    ELSE.
      result = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__va~is_error.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_string.
    result = abap_false.
  ENDMETHOD.
ENDCLASS.
