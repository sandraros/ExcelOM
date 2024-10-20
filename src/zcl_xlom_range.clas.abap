"! Has child classes ZCL_XLOM_COLUMNS and ZCL_XLOM_ROWS.
CLASS zcl_xlom_range DEFINITION
  PUBLIC
  CREATE PROTECTED
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.
    INTERFACES zif_xlom__va.
    INTERFACES zif_xlom__va_array.

    DATA application TYPE REF TO zcl_xlom_application READ-ONLY.
    DATA parent      TYPE REF TO zcl_xlom_worksheet   READ-ONLY.

    "! Address (RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo)
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.range.address
    "!
    "! @parameter row_absolute | True to return the row part of the reference as an absolute reference.
    "! @parameter column_absolute | True to return the column part of the reference as an absolute reference.
    "! @parameter reference_style | In A1 or R1C1 format.
    "! @parameter external | True to return an external reference. False to return a local reference.
    "! @parameter relative_to | If RowAbsolute and ColumnAbsolute are False, and ReferenceStyle is xlR1C1, you must include a starting point for the relative reference. This argument is a Range object that defines the starting point.
    "!                          NOTE: Testing with Excel VBA 7.1 shows that an explicit starting point is not mandatory. There appears to be a default reference of $A$1.
    "! @parameter result | Returns the address of the range, e.g. "A1", "$A$1", etc.
    METHODS address
      IMPORTING row_absolute    TYPE abap_bool                    DEFAULT abap_true
                column_absolute TYPE abap_bool                    DEFAULT abap_true
                reference_style TYPE zcl_xlom=>ty_reference_style DEFAULT zcl_xlom=>c_reference_style-a1
                external        TYPE abap_bool                    DEFAULT abap_false
                relative_to     TYPE REF TO zcl_xlom_range        OPTIONAL
      RETURNING VALUE(result)   TYPE string.

    METHODS calculate.

    "! Use either both row and column, or item alone.
    "! @parameter row    | Start from 1
    "! @parameter column | Start from 1.
    "! @parameter item   | Item number from 1, 16385 is the same as row = 2 column = 1.
    METHODS cells
      IMPORTING !row          TYPE i    OPTIONAL
                !column       TYPE i    OPTIONAL
                item          TYPE int8 OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    METHODS columns
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    METHODS count
      RETURNING VALUE(result) TYPE i.

    "! Called by the Worksheet.Range property.
    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    CLASS-METHODS create
      IMPORTING cell1         TYPE REF TO zcl_xlom_range
                cell2         TYPE REF TO zcl_xlom_range OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    CLASS-METHODS create_from_address_or_name
      IMPORTING address       TYPE clike
                relative_to   TYPE REF TO zcl_xlom_worksheet
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range
      RAISING   zcx_xlom__va.

    "! Range with integer row and column coordinates
    "! @parameter row    | Start from 1
    "! @parameter column | Start from 1
    CLASS-METHODS create_from_row_column
      IMPORTING worksheet     TYPE REF TO zcl_xlom_worksheet
                !row          TYPE i
                !column       TYPE i
                row_size      TYPE i DEFAULT 1
                column_size   TYPE i DEFAULT 1
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    CLASS-METHODS create_from_expr_range
      IMPORTING expr_range    TYPE REF TO zcl_xlom__ex_el_range
                relative_to   TYPE REF TO zcl_xlom_worksheet
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range
      RAISING   zcx_xlom__va.

    METHODS formula2
      RETURNING VALUE(result) TYPE REF TO zif_xlom__ex.

    "! Offset (RowOffset, ColumnOffset)
    "! https://learn.microsoft.com/fr-fr/office/vba/api/excel.range.offset
    "! @parameter row_offset    | Start from 0
    "! @parameter column_offset | Start from 0
    METHODS offset
      IMPORTING row_offset    TYPE i
                column_offset TYPE i
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    "! Resize (RowSize, ColumnSize)
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.range.resize
    "! @parameter row_size    | Start from 1
    "! @parameter column_size | Start from 1
    METHODS resize
      IMPORTING row_size      TYPE i
                column_size   TYPE i
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    METHODS row
      RETURNING VALUE(result) TYPE i.

    "! Rows ([item])
    "! - Rows() all rows
    "! - Rows(1) or Rows.Item(1) first row of the range
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.range.rows
    METHODS rows
      IMPORTING !index        TYPE i OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    METHODS set_formula2
      IMPORTING !value TYPE string
      RAISING   zcx_xlom__ex_ut_parser.

    METHODS set_value
      IMPORTING !value TYPE REF TO zif_xlom__va.

    METHODS value
      RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

  PROTECTED SECTION.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ts_formula_buffer_line,
        formula TYPE string,
        object  TYPE REF TO zif_xlom__ex,
      END OF ts_formula_buffer_line.
    TYPES tt_formula_buffer        TYPE HASHED TABLE OF ts_formula_buffer_line WITH UNIQUE KEY formula.
    "! By default, the "count" method counts the number of cells.
    "! It's possible to make it count only the columns or rows
    "! when the range is created by the methods "columns" or "rows".
    TYPES ty_column_row_collection TYPE i.
    TYPES:
      BEGIN OF ts_range_buffer_line,
        worksheet             TYPE REF TO zcl_xlom_worksheet,
        address               TYPE zcl_xlom=>ts_range_address,
        column_row_collection TYPE ty_column_row_collection,
        object                TYPE REF TO zcl_xlom_range,
      END OF ts_range_buffer_line.
    TYPES tt_range_buffer TYPE HASHED TABLE OF ts_range_buffer_line WITH UNIQUE KEY worksheet address column_row_collection.
    TYPES:
      BEGIN OF ts_range_name_or_coords,
        range_name TYPE string,
        column     TYPE i,
        row        TYPE i,
      END OF ts_range_name_or_coords.

    CONSTANTS:
      "! By default, the "count" method counts the number of cells.
      "! It's possible to make it count only the columns or rows
      "! when the range is created by the methods "columns" or "rows".
      BEGIN OF c_column_row_collection,
        none    TYPE ty_column_row_collection VALUE 1,
        columns TYPE ty_column_row_collection VALUE 2,
        rows    TYPE ty_column_row_collection VALUE 3,
      END OF c_column_row_collection.

    CLASS-DATA _formula_buffer TYPE tt_formula_buffer.
    CLASS-DATA _range_buffer   TYPE tt_range_buffer.

    DATA _address TYPE zcl_xlom=>ts_range_address.

    CLASS-METHODS convert_column_a_xfd_to_number
      IMPORTING roman_letters TYPE csequence
      RETURNING VALUE(result) TYPE i.

    CLASS-METHODS convert_column_number_to_a_xfd
      IMPORTING !number       TYPE i
      RETURNING VALUE(result) TYPE string.

    "! @parameter column_row_collection | By default, the "count" method counts the number of cells.
    "!                                    It's possible to make it count only the columns or rows
    "!                                    when the range is created by the methods "columns" or "rows".
    CLASS-METHODS create_from_top_left_bottom_ri
      IMPORTING worksheet             TYPE REF TO zcl_xlom_worksheet
                top_left              TYPE zcl_xlom=>ts_range_address-top_left
                bottom_right          TYPE zcl_xlom=>ts_range_address-bottom_right
                column_row_collection TYPE ty_column_row_collection DEFAULT c_column_row_collection-none
      RETURNING VALUE(result)         TYPE REF TO zcl_xlom_range.

    CLASS-METHODS decode_range_address
      IMPORTING address       TYPE string
      RETURNING VALUE(result) TYPE zif_xlom__va_array=>ts_address.

    CLASS-METHODS decode_range_address_a1
      IMPORTING address       TYPE string
      RETURNING VALUE(result) TYPE zif_xlom__va_array=>ts_address.

    CLASS-METHODS decode_range_coords
      IMPORTING words         TYPE string_table
                !from         TYPE i
                !to           TYPE i
      RETURNING VALUE(result) TYPE zif_xlom__va_array=>ts_address_one_cell.

    CLASS-METHODS decode_range_name_or_coords
      IMPORTING range_name_or_coords TYPE string
      RETURNING VALUE(result)        TYPE zcl_xlom=>ts_range_address_one_cell.

    METHODS _offset_resize
      IMPORTING row_offset    TYPE i
                column_offset TYPE i
                row_size      TYPE i
                column_size   TYPE i
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_range.

    CLASS-METHODS optimize_array_if_range
      IMPORTING array         TYPE REF TO zif_xlom__va_array
      RETURNING VALUE(result) TYPE zcl_xlom=>ts_range_address.
ENDCLASS.


CLASS zcl_xlom_range IMPLEMENTATION.
  METHOD address.
    IF reference_style <> zcl_xlom=>c_reference_style-a1.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    IF external = abap_true.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    IF relative_to IS BOUND.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.

    IF     _address-top_left-column     = 1
       AND _address-bottom_right-column = zcl_xlom_worksheet=>max_columns.
      " Whole rows (e.g. "$1:$1")
      result = |${ _address-top_left-row }:${ _address-bottom_right-row }|.
    ELSEIF     _address-top_left-row     = 1
           AND _address-bottom_right-row = zcl_xlom_worksheet=>max_rows.
      " Whole columns (e.g. "$A:$A")
      result = |${ zcl_xlom_range=>convert_column_number_to_a_xfd( _address-top_left-column )
                }:${ zcl_xlom_range=>convert_column_number_to_a_xfd( _address-bottom_right-column ) }|.
    ELSE.
      " one cell (e.g. "$A$1") or several cells (e.g. "$A$1:$A$2")
      result = |${ zcl_xlom_range=>convert_column_number_to_a_xfd( _address-top_left-column )
               }${ _address-top_left-row
               }{ COND #( WHEN _address-bottom_right <> _address-top_left
                          THEN |:${ zcl_xlom_range=>convert_column_number_to_a_xfd( _address-bottom_right-column )
                                }${ _address-bottom_right-row }| ) }|.
    ENDIF.
  ENDMETHOD.

  METHOD calculate.
    IF _address-top_left <> _address-bottom_right.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    DATA(cell) = REF #( parent->_array->_cells[ column = _address-top_left-column
                                                row    = _address-top_left-row ] OPTIONAL ).
    IF cell IS NOT BOUND.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    DATA(context) = zcl_xlom__ex_ut_eval_context=>create(
                        worksheet       = parent
                        containing_cell = VALUE #( row    = _address-top_left-row
                                                   column = _address-top_left-column ) ).
    zcl_xlom__ex_ut_eval=>evaluate_array_operands( expression = cell->formula
                                                   context    = context ).
*    cell->formula->evaluate( context ).
  ENDMETHOD.

  METHOD cells.
    IF item IS NOT INITIAL.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    " This will change Z20:
    " Range("Z20:AA25").Cells(1, 1) = "C"
    result = zcl_xlom_range=>create_from_row_column( worksheet = parent
                                                     row       = _address-top_left-row + row - 1
                                                     column    = _address-top_left-column + column - 1 ).
  ENDMETHOD.

  METHOD columns.
    result = zcl_xlom_range=>create_from_top_left_bottom_ri( worksheet             = parent
                                                             top_left              = _address-top_left
                                                             bottom_right          = _address-bottom_right
                                                             column_row_collection = c_column_row_collection-columns ).
  ENDMETHOD.

  METHOD convert_column_a_xfd_to_number.
    DATA(offset) = 0.
    WHILE offset < strlen( roman_letters ).
      FIND roman_letters+offset(1) IN sy-abcde MATCH OFFSET DATA(offset_a_to_z).
      IF sy-subrc <> 0.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      ENDIF.
      result = ( result * 26 ) + offset_a_to_z + 1.
      IF result > zcl_xlom_worksheet=>max_columns.
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      ENDIF.
      offset = offset + 1.
    ENDWHILE.
  ENDMETHOD.

  METHOD convert_column_number_to_a_xfd.
    IF number NOT BETWEEN 1 AND zcl_xlom_worksheet=>max_columns.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    DATA(work_number) = number.
    DO.
      DATA(lv_mod) = ( work_number - 1 ) MOD 26.
      DATA(lv_div) = ( work_number - 1 ) DIV 26.
      work_number = lv_div.
      result = sy-abcde+lv_mod(1) && result.
      IF work_number <= 0.
        EXIT.
      ENDIF.
    ENDDO.
  ENDMETHOD.

  METHOD count.
    result = zif_xlom__va_array~column_count * zif_xlom__va_array~row_count.
  ENDMETHOD.

  METHOD create.
    IF cell1 IS NOT BOUND.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    IF cell2 IS NOT BOUND.
      result = cell1.
      RETURN.
    ENDIF.
    IF cell1->parent <> cell2->parent.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    " This will set "g" from Z20 to AD24:
    " Range(Range("Z20:AA21"), Range("AC23:AD24")).Value = "g"
    " This will set "H" from Z20 to AD25:
    " Range(Range("Z20:AA25"), Range("AC23:AD24")).Value = "H"
    " This will set "H" from Z20 to AD25:
    " Range(Range("AA25:Z20"), Range("AD24:AC23")).Value = "H"
    DATA(structured_address) = VALUE zcl_xlom=>ts_range_address(
        top_left     = VALUE #( column = nmin( val1 = nmin( val1 = cell1->_address-top_left-column
                                                            val2 = cell2->_address-top_left-column )
                                               val2 = nmin( val1 = cell1->_address-bottom_right-column
                                                            val2 = cell2->_address-bottom_right-column ) )
                                row    = nmin( val1 = nmin( val1 = cell1->_address-top_left-row
                                                            val2 = cell2->_address-top_left-row )
                                               val2 = nmin( val1 = cell1->_address-bottom_right-row
                                                            val2 = cell2->_address-bottom_right-row ) ) )
        bottom_right = VALUE #( column = nmax( val1 = nmax( val1 = cell1->_address-top_left-column
                                                            val2 = cell2->_address-top_left-column )
                                               val2 = nmax( val1 = cell1->_address-bottom_right-column
                                                            val2 = cell2->_address-bottom_right-column ) )
                                row    = nmax( val1 = nmax( val1 = cell1->_address-top_left-row
                                                            val2 = cell2->_address-top_left-row )
                                               val2 = nmax( val1 = cell1->_address-bottom_right-row
                                                            val2 = cell2->_address-bottom_right-row ) ) ) ).
    result = create_from_top_left_bottom_ri( worksheet    = cell1->parent
                                             top_left     = structured_address-top_left
                                             bottom_right = structured_address-bottom_right ).
  ENDMETHOD.

  METHOD create_from_address_or_name.
    DATA(structured_address) = decode_range_address( address ).
    result = create_from_top_left_bottom_ri(
        worksheet    = COND #( WHEN structured_address-worksheet_name IS INITIAL
                               THEN relative_to
                               ELSE relative_to->parent->worksheets->item( structured_address-worksheet_name ) )
        top_left     = VALUE #( column = structured_address-top_left-column
                                row    = structured_address-top_left-row )
        bottom_right = VALUE #( column = structured_address-bottom_right-column
                                row    = structured_address-bottom_right-row ) ).
  ENDMETHOD.

  METHOD create_from_expr_range.
    result = create_from_address_or_name( address     = expr_range->_address_or_name
                                          relative_to = relative_to ).
  ENDMETHOD.

  METHOD create_from_row_column.
    result = create_from_top_left_bottom_ri( worksheet    = worksheet
                                             top_left     = VALUE #( column = column
                                                                     row    = row )
                                             bottom_right = VALUE #( column = column + column_size - 1
                                                                     row    = row + row_size - 1 ) ).
  ENDMETHOD.

  METHOD create_from_top_left_bottom_ri.
    DATA range TYPE REF TO zcl_xlom_range.

    " If row = 0, it means the range is a whole column (rows from 1 to 1048576).
    " If column = 0, it means the range is a whole row (columns from 1 to 16384).
    DATA(address) = VALUE zcl_xlom=>ts_range_address(
                              top_left     = VALUE #( row    = COND #( WHEN top_left-row > 0
                                                                       THEN top_left-row
                                                                       ELSE 1 )
                                                      column = COND #( WHEN top_left-column > 0
                                                                       THEN top_left-column
                                                                       ELSE 1 ) )
                              bottom_right = VALUE #( row    = COND #( WHEN bottom_right-row > 0
                                                                       THEN bottom_right-row
                                                                       ELSE zcl_xlom_worksheet=>max_rows )
                                                      column = COND #( WHEN bottom_right-column > 0
                                                                       THEN bottom_right-column
                                                                       ELSE zcl_xlom_worksheet=>max_columns ) ) ).
    DATA(range_buffer_line) = REF #( _range_buffer[ worksheet             = worksheet
                                                    address               = address
                                                    column_row_collection = column_row_collection ] OPTIONAL ).
    IF range_buffer_line IS NOT BOUND.
      CASE column_row_collection.
        WHEN c_column_row_collection-columns.
          range = NEW zcl_xlom_columns( ).
        WHEN c_column_row_collection-rows.
          range = NEW zcl_xlom_rows( ).
        WHEN c_column_row_collection-none.
          range = NEW zcl_xlom_range( ).
        WHEN OTHERS.
          RAISE EXCEPTION TYPE zcx_xlom_unexpected.
      ENDCASE.
      range->zif_xlom__va~type               = zif_xlom__va=>c_type-range.
      range->application                     = worksheet->application.
      range->parent                          = worksheet.
      range->_address                        = address.
      range->zif_xlom__va_array~row_count    = address-bottom_right-row - address-top_left-row + 1.
      range->zif_xlom__va_array~column_count = address-bottom_right-column - address-top_left-column + 1.
      INSERT VALUE #( worksheet             = worksheet
                      address               = address
                      column_row_collection = column_row_collection
                      object                = range )
             INTO TABLE _range_buffer
             REFERENCE INTO range_buffer_line.
    ENDIF.
    result = range_buffer_line->object.
  ENDMETHOD.

  METHOD decode_range_address.
    " The range address should always be in A1 reference style.
    result = zcl_xlom_range=>decode_range_address_a1( address ).
    IF result-top_left IS INITIAL.
      " address is an invalid range address so it's probably referring to an existing name.
      " Maybe it's a table name
      " Maybe it's a range name (local or global scope, i.e. worksheet or workbook)
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
  ENDMETHOD.

  METHOD decode_range_address_a1.
    " Special characters are ":", "$", "!", "'", "[" and "]".
    " When "'" is found, all characters till the next "'" form
    "   one word or two words if there are "[" and "]".
    " When "[" is found, all characters till the next "]" form
    "   the workbook name.
    " Subsequent non-special characters form a word.
    "
    " Examples:
    "   In the current worksheet:
    "     NB: they are case-insensitive
    "     A1 (relative column and row)             word
    "     $A1 (absolute column, relative column)   $ word
    "     A$1                                      word $ word
    "     $A$1                                     $ word $ word
    "     A1:A2                                    word : word
    "     $A$A                                     $ word $ word
    "     A:A                                      word : word
    "     1:1                                      word : word
    "     NAME                                     word
    "   Other worksheet:
    "     Sheet1!A1                                word ! word
    "     'Sheet 1'!A1                             word ! word
    "     [1]Sheet1!$A$3                           [word] word ! $ word $ word   (XLSX internal notation for workbooks)
    "   Other workbook:
    "     '[C:\workbook.xlsx]'!NAME                [word] ! word                 (workbook absolute path / name in the global scope)
    "     '[workbook.xlsx]Sheet 1'!$A$1            [word] word ! word            (workbook relative path)
    "     [1]!NAME                                 [word] ! word                 (XLSX internal notation for workbooks)
    TYPES ty_state TYPE i.

    CONSTANTS:
      BEGIN OF c_state,
        normal                        TYPE ty_state VALUE 1,
        within_single_quotes          TYPE ty_state VALUE 2,
        within_single_quotes_brackets TYPE ty_state VALUE 3,
        within_brackets               TYPE ty_state VALUE 4,
      END OF c_state.
    DATA colon_position TYPE i.

    DATA(words) = VALUE string_table( ).
    INSERT INITIAL LINE INTO TABLE words REFERENCE INTO DATA(current_word).
    DATA(state) = c_state-normal.
    DATA(offset) = 0.
    WHILE offset < strlen( address ).
      DATA(character) = substring( val = address
                                   off = offset
                                   len = 1 ).

      DATA(start_a_new_word) = abap_false.
      DATA(store_dedicated_word) = abap_false.
      CASE state.
        WHEN c_state-normal.
          CASE character.
            WHEN ''''.
              state = c_state-within_single_quotes.
            WHEN '$'.
              store_dedicated_word = abap_true.
            WHEN '!'.
              store_dedicated_word = abap_true.
            WHEN ':'.
              store_dedicated_word = abap_true.
            WHEN '['.
              DATA(square_bracket_position) = 1.
              state = c_state-within_brackets.
            WHEN OTHERS.
              current_word->* = current_word->* && character.
          ENDCASE.
        WHEN c_state-within_single_quotes.
          CASE character.
            WHEN ''''.
              start_a_new_word = abap_true.
              state = c_state-normal.
            WHEN '['.
              square_bracket_position = 1.
              state = c_state-within_single_quotes_brackets.
            WHEN OTHERS.
              current_word->* = current_word->* && character.
          ENDCASE.
        WHEN c_state-within_single_quotes_brackets.
          CASE character.
            WHEN ']'.
              start_a_new_word = abap_true.
              state = c_state-within_single_quotes.
            WHEN OTHERS.
              current_word->* = current_word->* && character.
          ENDCASE.
        WHEN c_state-within_brackets.
          CASE character.
            WHEN ']'.
              start_a_new_word = abap_true.
              state = c_state-normal.
            WHEN OTHERS.
              current_word->* = current_word->* && character.
          ENDCASE.
      ENDCASE.
      IF    start_a_new_word     = abap_true
         OR store_dedicated_word = abap_true.
        IF current_word->* IS NOT INITIAL.
          INSERT INITIAL LINE INTO TABLE words REFERENCE INTO current_word.
        ENDIF.
        CASE character.
          WHEN '!'.
            DATA(exclamation_mark_position) = lines( words ).
          WHEN ':'.
            colon_position = lines( words ).
        ENDCASE.
        IF store_dedicated_word = abap_true.
          current_word->* = character.
          INSERT INITIAL LINE INTO TABLE words REFERENCE INTO current_word.
        ENDIF.
        start_a_new_word = abap_false.
        store_dedicated_word = abap_false.
      ENDIF.
      offset = offset + 1.
    ENDWHILE.

    IF square_bracket_position = 1.
      result-workbook_name = words[ 1 ].
    ENDIF.

    IF    exclamation_mark_position = 3
       OR (     exclamation_mark_position = 2
            AND square_bracket_position   = 0 ).
      result-worksheet_name = words[ exclamation_mark_position - 1 ].
    ENDIF.

    IF colon_position = 0.
      IF exclamation_mark_position + 1 = lines( words ).
        " A1 or NAME
        result-top_left = decode_range_coords( words = words
                                               from  = lines( words )
                                               to    = lines( words ) ).
        IF result-top_left IS INITIAL.
          " NAME
          result-range_name = words[ lines( words ) ].
        ELSE.
          result-bottom_right = result-top_left.
        ENDIF.
      ELSE.
        " $A$1, A$1, $A1
        result-top_left     = decode_range_coords( words = words
                                                   from  = exclamation_mark_position + 1
                                                   to    = lines( words ) ).
        result-bottom_right = result-top_left.
      ENDIF.
    ELSE.
      " A1:A2, $A$1:$B$2, A1:$B$2, etc.
      result-top_left     = decode_range_coords( words = words
                                                 from  = exclamation_mark_position + 1
                                                 to    = colon_position - 1 ).
      result-bottom_right = decode_range_coords( words = words
                                                 from  = colon_position + 1
                                                 to    = lines( words ) ).
    ENDIF.
  ENDMETHOD.

  METHOD decode_range_coords.
    " Remove $ if any
    DATA(coords_without_dollar) = REDUCE #( INIT t = ``
                       FOR <word> IN words FROM from TO to
                       WHERE ( table_line <> '$' )
                       NEXT t = t && <word> ).

    DATA(coords) = decode_range_name_or_coords( coords_without_dollar ).

    IF coords IS INITIAL.
      RETURN.
    ENDIF.

    result = VALUE #( column = coords-column
                      row    = coords-row ).

    IF words[ from ] = '$'.
      result-column_fixed = abap_true.
      IF     from + 2          <= lines( words )
         AND words[ from + 2 ]  = '$'.
        result-row_fixed = abap_true.
      ENDIF.
    ELSEIF     from              < lines( words )
           AND words[ from + 1 ] = '$'.
      result-row_fixed = abap_true.
    ENDIF.
  ENDMETHOD.

  METHOD decode_range_name_or_coords.
    DATA(offset) = 0.
    WHILE     offset < strlen( range_name_or_coords )
          AND range_name_or_coords+offset(1) NA '123456789'.
      offset = offset + 1.
    ENDWHILE.
    IF     offset <= 3
       AND range_name_or_coords(offset) CO 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
       AND (    offset < 3
             OR (     offset                   = 3
                  AND range_name_or_coords(3) <= 'XFD' ) ).
      IF     range_name_or_coords+offset           CO '1234567890'
         AND CONV i( range_name_or_coords+offset ) <= zcl_xlom_worksheet=>max_rows.
        result-column = convert_column_a_xfd_to_number( substring( val = range_name_or_coords
                                                                   len = offset ) ).
        result-row    = range_name_or_coords+offset.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD formula2.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD offset.
    result = _offset_resize( row_offset    = row_offset
                             column_offset = column_offset
                             row_size      = _address-bottom_right-row - _address-top_left-row + 1
                             column_size   = _address-bottom_right-column - _address-top_left-column + 1 ).
  ENDMETHOD.

  METHOD optimize_array_if_range.
*    DATA(row_count) = 0.
*    DATA(column_count) = 0.
    IF array->zif_xlom__va~type = array->zif_xlom__va~c_type-range.
      DATA(range) = CAST zcl_xlom_range( array ).
      result = range->application->_intersect_2_basis(
                   arg1 = VALUE #( top_left-column     = range->_address-top_left-column
                                   top_left-row        = range->_address-top_left-row
                                   bottom_right-column = range->_address-bottom_right-column
                                   bottom_right-row    = range->_address-bottom_right-row )
                   arg2 = range->parent->_array->used_range ).
    ELSE.
      result = VALUE #( top_left-column     = 1
                        top_left-row        = 1
                        bottom_right-column = array->column_count
                        bottom_right-row    = array->row_count ).
    ENDIF.
  ENDMETHOD.

  METHOD resize.
    result = _offset_resize( row_offset    = 0
                             column_offset = 0
                             row_size      = row_size
                             column_size   = column_size ).
  ENDMETHOD.

  METHOD row.
    result = _address-top_left-row.
  ENDMETHOD.

  METHOD rows.
    DATA(bottom_right) = COND zcl_xlom=>ts_range_address_one_cell( WHEN index = 0 THEN
                                                                     _address-bottom_right
                                                                   WHEN index >= 1
                                                                    AND index <= ( _address-bottom_right-row - _address-top_left-row + 1 ) THEN
                                                                     VALUE #(
                                                                         column = _address-bottom_right-column
                                                                         row    = _address-top_left-row + index - 1 )
                                                                   ELSE
                                                                     THROW zcx_xlom_todo( ) ).
    result = zcl_xlom_range=>create_from_top_left_bottom_ri( worksheet             = parent
                                                             top_left              = _address-top_left
                                                             bottom_right          = bottom_right
                                                             column_row_collection = c_column_row_collection-rows ).
  ENDMETHOD.

  METHOD set_formula2.
    DATA(formula_buffer_line) = REF #( _formula_buffer[ formula = value ] OPTIONAL ).
    IF formula_buffer_line IS NOT BOUND.
      DATA(lexer) = zcl_xlom__ex_ut_lexer=>create( ).
      DATA(lexer_tokens) = lexer->lexe( value ).
      DATA(parser) = zcl_xlom__ex_ut_parser=>create( ).
      INSERT VALUE #( formula = value
                      object  = parser->parse( lexer_tokens ) )
             INTO TABLE _formula_buffer
             REFERENCE INTO formula_buffer_line.
    ENDIF.
    DATA(formula_expression) = formula_buffer_line->object.

    IF application->calculation = zcl_xlom=>c_calculation-automatic.
          data(cell_value) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = formula_expression
                                 context    = zcl_xlom__ex_ut_eval_context=>create(
                                                                 worksheet       = parent
                                                                 containing_cell = VALUE #(
                                                                     row    = _address-top_left-row
                                                                     column = _address-top_left-column ) ) ).
      parent->_array->zif_xlom__va_array~set_cell_value( row        = _address-top_left-row
                                                         column     = _address-top_left-column
                                                         formula    = formula_expression
                                                         calculated = abap_true
                                                         value      = cell_value ).
*                                                         value      = formula_expression->evaluate(
*                                                             context = zcl_xlom__ex_ut_eval_context=>create(
*                                                                 worksheet       = parent
*                                                                 containing_cell = VALUE #(
*                                                                     row    = _address-top_left-row
*                                                                     column = _address-top_left-column ) ) ) ).
    ELSE.
      parent->_array->zif_xlom__va_array~set_cell_value( row        = _address-top_left-row
                                                         column     = _address-top_left-column
                                                         value      = zcl_xlom__va_number=>get( 0 )
                                                         formula    = formula_expression
                                                         calculated = abap_false ).
    ENDIF.
  ENDMETHOD.

  METHOD set_value.
    parent->_array->zif_xlom__va_array~set_cell_value( row    = _address-top_left-row
                                                       column = _address-top_left-column
                                                       value  = value ).
  ENDMETHOD.

  METHOD value.
    IF _address-top_left = _address-bottom_right.
      result = parent->_array->zif_xlom__va_array~get_cell_value( column = _address-top_left-column
                                                                  row    = _address-top_left-row ).
    ELSE.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
*      IF     ZIF_xlom_result_array~row_count    > 1
*         AND ZIF_xlom_result_array~column_count > 1.
*        result = ZCL_xlom__va_itab_2=>create( VALUE #( ( row = 1 column  )  ) ).
*      ELSEIF ZIF_xlom_result_array~row_count > 1.
*        result = ZCL_xlom__va_itab_1=>create( ).
*      ELSE.
*        result = ZCL_xlom__va_itab_1=>create( ).
*      ENDIF.
*      result = parent->_array->ZIF_xlom_result_array~get_array_value( top_left     = _address-top_left
*                                                                      bottom_right = _address-bottom_right ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__va_array~get_array_value.
    result = _offset_resize( row_offset    = top_left-row - 1
                             column_offset = top_left-column - 1
                             row_size      = bottom_right-row - top_left-row + 1
                             column_size   = bottom_right-column - top_left-column + 1 ).
  ENDMETHOD.

  METHOD zif_xlom__va_array~get_cell_value.
    " if the current range starts from row 2 column 2 and the requested cell is row 2 column 2
    " then get the worksheet cell from row 3 column 3.
    result = parent->_array->zif_xlom__va_array~get_cell_value( column = _address-top_left-column + column - 1
                                                                row    = _address-top_left-row + row - 1 ).
  ENDMETHOD.

  METHOD zif_xlom__va_array~set_array_value.
    parent->_array->zif_xlom__va_array~set_array_value( rows = rows ).
  ENDMETHOD.

  METHOD zif_xlom__va_array~set_cell_value.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~get_value.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_array.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_boolean.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_equal.
    IF input_result->type = zif_xlom__va=>c_type-range.
      DATA(input_range) = CAST zcl_xlom_range( input_result ).
      IF me->_address = input_range->_address.
        result = abap_true.
      ELSE.
        result = abap_false.
      ENDIF.
    ELSE.
      result = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__va~is_error.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_number.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_string.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD _offset_resize.
    result = create_from_row_column( worksheet   = parent
                                     row         = _address-top_left-row + row_offset
                                     column      = _address-top_left-column + column_offset
                                     row_size    = row_size
                                     column_size = column_size ).
  ENDMETHOD.
ENDCLASS.
