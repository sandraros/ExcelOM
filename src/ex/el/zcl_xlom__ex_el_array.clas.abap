CLASS zcl_xlom__ex_el_array DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.
    INTERFACES zif_xlom__ex_array.

    TYPES tt_column TYPE STANDARD TABLE OF REF TO zif_xlom__ex WITH EMPTY KEY.
    TYPES:
      BEGIN OF ts_row,
        columns_of_row TYPE tt_column,
      END OF ts_row.
    TYPES tt_row TYPE STANDARD TABLE OF ts_row WITH EMPTY KEY.

    CLASS-METHODS create
      IMPORTING !rows         TYPE tt_row
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_el_array.

  PRIVATE SECTION.
    DATA rows TYPE tt_row.
ENDCLASS.


CLASS zcl_xlom__ex_el_array IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_el_array( ).
    result->zif_xlom__ex~type = result->zif_xlom__ex~c_type-array.
    result->rows              = rows.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    result = zif_xlom__ex~set_result( zcl_xlom__va_array=>create_initial(
                                          row_count    = lines( rows )
                                          column_count = REDUCE #( INIT n = 0
                                                                        FOR <row> IN rows
                                                                        NEXT n = nmax(
                                                                            val1 = n
                                                                            val2 = lines( <row>-columns_of_row ) ) )
                                          rows         = VALUE #(
                                              FOR <row> IN rows
                                              ( columns_of_row = VALUE #( FOR <column> IN <row>-columns_of_row
                                                                          ( <column>->evaluate( context ) ) ) ) ) ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    DATA(array) = CAST zcl_xlom__ex_el_array( expression ).
    IF lines( rows ) <> lines( array->rows ).
      RETURN.
    ENDIF.

    DATA(row_tabix) = 1.
    WHILE row_tabix <= lines( rows ).

      DATA(ref_columns) = REF #( rows[ row_tabix ]-columns_of_row ).
      DATA(ref_array_columns) = REF #( array->rows[ row_tabix ]-columns_of_row ).
      IF lines( ref_columns->* ) <> lines( ref_array_columns->* ).
        result = abap_false.
        RETURN.
      ENDIF.

      DATA(column_tabix) = 1.
      WHILE column_tabix <= lines( ref_columns->* ).
        IF abap_false = ref_columns->*[ column_tabix ]->is_equal( ref_array_columns->*[ column_tabix ] ).
          RETURN.
        ENDIF.
        column_tabix = column_tabix + 1.
      ENDWHILE.

      row_tabix = row_tabix + 1.
    ENDWHILE.

    result = abap_true.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
