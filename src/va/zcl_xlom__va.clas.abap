CLASS zcl_xlom__va DEFINITION
  PUBLIC
  CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    CLASS-METHODS to_boolean
      IMPORTING !input        TYPE REF TO zif_xlom__va
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_boolean.

    CLASS-METHODS to_number
      IMPORTING !input        TYPE REF TO zif_xlom__va
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_number
      RAISING   zcx_xlom__va.

    CLASS-METHODS to_string
      IMPORTING !input        TYPE REF TO zif_xlom__va
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_string
      RAISING   zcx_xlom__va.

    CLASS-METHODS to_range
      IMPORTING !input        TYPE REF TO zif_xlom__va
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range
      RAISING   zcx_xlom__va.

    CLASS-METHODS to_array
      IMPORTING !input        TYPE REF TO zif_xlom__va
      RETURNING VALUE(result) TYPE REF TO zif_xlom__va_array
      RAISING   zcx_xlom__va.

  PROTECTED SECTION.

  PRIVATE SECTION.
ENDCLASS.


CLASS zcl_xlom__va IMPLEMENTATION.
  METHOD to_array.
    CASE input->type.
      WHEN input->c_type-error.
        " TODO I didn't check whether it should be #N/A, #REF! or #VALUE!
        RAISE EXCEPTION TYPE zcx_xlom__va
          EXPORTING result_error = zcl_xlom__va_error=>value_cannot_be_calculated.
      WHEN input->c_type-array
        OR input->c_type-range.
        result = CAST #( input ).
      WHEN OTHERS.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDCASE.
  ENDMETHOD.

  METHOD to_boolean.
    " If source is Number:
    "       FALSE if 0
    "       TRUE if not 0
    "
    "  If source is String:
    "       Language-dependent.
    "       In English:
    "       TRUE if "TRUE"
    "       FALSE if "FALSE"
    CASE input->type.
      WHEN input->c_type-array.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      WHEN input->c_type-boolean.
        result = CAST #( input ).
      WHEN input->c_type-error.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      WHEN input->c_type-number.
        CASE CAST zcl_xlom__va_number( input )->get_number( ).
          WHEN 0.
            result = zcl_xlom__va_boolean=>false.
          WHEN OTHERS.
            result = zcl_xlom__va_boolean=>true.
        ENDCASE.
      WHEN input->c_type-range.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      WHEN input->c_type-string.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      WHEN OTHERS.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDCASE.
  ENDMETHOD.

  METHOD to_number.
    " If source is Boolean:
    "      0 if FALSE
    "      1 if TRUE
    "
    " If source is String:
    "      Language-dependent for decimal separator.
    "      "." if English, "," if French, etc.
    "      Accepted: "-1", "+1", ".5", "1E1", "-.5"
    "      "-1E-1", "1e1", "1e307", "1e05", "1e-309"
    "      Invalid: "", "E1", "1e308", "1e-310"
    "      #VALUE! if invalid decimal separator
    "      #VALUE! if invalid number
    IF     input->type <> input->c_type-array
       AND input->type <> input->c_type-range.
      DATA(cell) = input.
    ELSE.
*      DATA(range) = CAST ZCL_xlom_range( input ).
      DATA(range) = CAST zif_xlom__va_array( input ).
      IF    range->column_count <> 1
         OR range->row_count    <> 1.
*      IF range->top_left <> range->_address-bottom_right.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      ENDIF.
      cell = range->get_cell_value( column = 1
                                    row    = 1 ).
    ENDIF.

    CASE cell->type.
      WHEN cell->c_type-array.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      WHEN cell->c_type-boolean.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      WHEN cell->c_type-empty.
        result = zcl_xlom__va_number=>get( 0 ).
      WHEN cell->c_type-error.
        RAISE EXCEPTION TYPE zcx_xlom__va
          EXPORTING result_error = CAST #( cell ).
      WHEN cell->c_type-number.
        result = CAST #( cell ).
      WHEN cell->c_type-range.
        " impossible because processed in the previous block.
        RAISE EXCEPTION TYPE zcx_xlom_unexpected.
      WHEN cell->c_type-string.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      WHEN OTHERS.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDCASE.
  ENDMETHOD.

  METHOD to_range.
    CASE input->type.
      WHEN input->c_type-error.
        " TODO I didn't check whether it should be #N/A, #REF! or #VALUE!
        RAISE EXCEPTION TYPE zcx_xlom__va
          EXPORTING result_error = zcl_xlom__va_error=>value_cannot_be_calculated.
      WHEN input->c_type-range.
        result = CAST #( input ).
      WHEN OTHERS.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDCASE.
  ENDMETHOD.

  METHOD to_string.
    CASE input->type.
      WHEN input->c_type-empty.
        result = zcl_xlom__va_string=>create( '' ).
      WHEN input->c_type-error.
        RAISE EXCEPTION TYPE zcx_xlom__va
          EXPORTING result_error = CAST #( input ).
      WHEN input->c_type-number.
        result = zcl_xlom__va_string=>create( |{ CAST zcl_xlom__va_number( input )->get_number( ) }| ).
      WHEN input->c_type-range.
        DATA(range) = CAST zcl_xlom_range( input ).
        IF range->_address-top_left <> range->_address-bottom_right.
          RAISE EXCEPTION TYPE zcx_xlom_todo.
        ENDIF.
        DATA(cell) = REF #( range->parent->_array->_cells[ row    = range->_address-top_left-row
                                                           column = range->_address-top_left-column ] OPTIONAL ).
        DATA(string) = COND string( WHEN cell IS BOUND
                                    THEN SWITCH #( cell->value->type
                                                   WHEN zif_xlom__va=>c_type-number THEN
                                                     |{ CAST zcl_xlom__va_number( cell->value )->get_number( ) }|
                                                   WHEN zif_xlom__va=>c_type-string THEN
                                                     CAST zcl_xlom__va_string( cell->value )->get_string( )
                                                   ELSE
                                                     THROW zcx_xlom_todo( ) ) ).
        result = zcl_xlom__va_string=>create( string ).
      WHEN input->c_type-string.
        result = CAST #( input ).
      WHEN OTHERS.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDCASE.
  ENDMETHOD.
ENDCLASS.
