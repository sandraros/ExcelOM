"! MATCH(lookup_value, lookup_array, [match_type])
"! https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a
CLASS zcl_xlom__ex_fu_match DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.
    INTERFACES zif_xlom__ex.

    "! MATCH(lookup_value, lookup_array, [match_type])
    "! https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a
    "! MATCH returns the position of the matched value within lookup_array, not the value itself. For example, MATCH("b",{"a","b","c"},0) returns 2, which is the relative position of "b" within the array {"a","b","c"}.
    "! MATCH does not distinguish between uppercase and lowercase letters when matching text values.
    "! If MATCH is unsuccessful in finding a match, it returns the #N/A error value.
    "! If match_type is 0 and lookup_value is a text string, you can use the wildcard characters — the question mark (?) and asterisk (*) — in the lookup_value argument.
    "! A question mark matches any single character; an asterisk matches any sequence of characters.
    "! If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    "! @parameter lookup_value | Required. The value that you want to match in lookup_array.
    "!                           For example, when you look up someone's number in a telephone book,
    "!                           you are using the person's name as the lookup value, but the telephone number is the value you want.
    "!                           The lookup_value argument can be a value (number, text, or logical value)
    "!                           or a cell reference to a number, text, or logical value.
    "! @parameter lookup_array | Required. The range of cells being searched.
    "! @parameter match_type   | Optional. The number -1, 0, or 1. The match_type argument specifies how Excel
    "!                           matches lookup_value with values in lookup_array. The default value for this argument is 1.
    "!                           <ul>
    "!                           <li>1 or omitted: MATCH finds the largest value that is less than or equal to lookup_value.
    "!                                             The values in the lookup_array argument must be placed in ascending order, for example:
    "!                                             ...-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE.</li>
    "!                           <li>0: MATCH finds the first value that is exactly equal to lookup_value.
    "!                                  The values in the lookup_array argument can be in any order.</li>
    "!                           <li>-1: MATCH finds the smallest value that is greater than or equal tolookup_value. The values in the lookup_array argument
    "!                                   must be placed in descending order, for example: TRUE, FALSE, Z-A, ...2, 1, 0, -1, -2, ..., and so on.</li>
    "!                           </ul>
    "! @parameter result       | MATCH returns the position of the matched value within lookup_array, not the value itself.
    "!                           For example, MATCH("b",{"a","b","c"},0) returns 2, which is the relative position of "b" within the array {"a","b","c"}.
    CLASS-METHODS create
      IMPORTING lookup_value  TYPE REF TO zif_xlom__ex
                lookup_array  TYPE REF TO zif_xlom__ex OPTIONAL
                match_type    TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_match.

  PRIVATE SECTION.
    DATA lookup_value TYPE REF TO zif_xlom__ex.
    DATA lookup_array TYPE REF TO zif_xlom__ex.
    DATA match_type   TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_match IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_match( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-match.
    result->lookup_value      = lookup_value.
    result->lookup_array      = lookup_array.
    result->match_type        = match_type.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
        expression = me
        context    = context
        operands   = VALUE #( ( name = 'LOOKUP_VALUE' object = lookup_value )
                              ( name = 'LOOKUP_ARRAY' object = lookup_array not_part_of_result_array = abap_true )
                              ( name = 'MATCH_TYPE  ' object = match_type ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    DATA temp_result TYPE REF TO zif_xlom__va.

    TRY.
        DATA(lookup_value_result) = arguments[ name = 'LOOKUP_VALUE' ]-object.
        DATA(lookup_array_result) = zcl_xlom__va=>to_array( arguments[ name = 'LOOKUP_ARRAY' ]-object ).
        DATA(match_type_result) = COND i( LET result_num_chars = arguments[ name = 'MATCH_TYPE' ]-object IN
                                          WHEN result_num_chars IS BOUND
                                          THEN COND #( WHEN result_num_chars->type = result_num_chars->c_type-empty
                                                       THEN 1
                                                       ELSE zcl_xlom__va=>to_number( result_num_chars )->get_integer( ) ) ).
        IF match_type_result <> 0.
          RAISE EXCEPTION TYPE zcx_xlom_todo.
        ENDIF.
        IF     lookup_array_result->row_count    > 1
           AND lookup_array_result->column_count > 1.
          " MATCH cannot lookup a two-dimension array, it can search either one row or one column.
          temp_result = zcl_xlom__va_error=>na_not_applicable.
        ELSE.
          DATA(optimized_lookup_array) = zcl_xlom_range=>optimize_array_if_range( lookup_array_result ).
          IF optimized_lookup_array IS NOT INITIAL.
            DATA(row_number) = optimized_lookup_array-top_left-row.
            WHILE row_number <= optimized_lookup_array-bottom_right-row.
              DATA(column_number) = optimized_lookup_array-top_left-column.
              WHILE column_number <= optimized_lookup_array-bottom_right-column.
                DATA(cell_value) = lookup_array_result->get_cell_value( column = column_number
                                                                        row    = row_number ).
                DATA(equal) = abap_false.
                DATA(lookup_value_result_2) = SWITCH #( lookup_value_result->type
                                                        WHEN zif_xlom__va=>c_type-array
                                                          OR zif_xlom__va=>c_type-range
                                                        THEN CAST zif_xlom__va( CAST zif_xlom__va_array( lookup_value_result )->get_cell_value(
                                                                                    column = 1
                                                                                    row    = 1 ) )
                                                        ELSE lookup_value_result ).
                IF    lookup_value_result_2->type = zif_xlom__va=>c_type-string
                   OR cell_value->type            = zif_xlom__va=>c_type-string.
                  equal = xsdbool( zcl_xlom__va=>to_string( lookup_value_result_2 )->get_string( )
                                   = zcl_xlom__va=>to_string( cell_value )->get_string( ) ).
                ELSEIF    lookup_value_result_2->type = zif_xlom__va=>c_type-number
                       OR cell_value->type            = zif_xlom__va=>c_type-number.
                  equal = xsdbool( zcl_xlom__va=>to_number( lookup_value_result_2 )->get_number( )
                                   = zcl_xlom__va=>to_number( cell_value )->get_number( ) ).
                ELSEIF    lookup_value_result_2->type = zif_xlom__va=>c_type-boolean
                       OR cell_value->type            = zif_xlom__va=>c_type-boolean.
                  equal = xsdbool( zcl_xlom__va=>to_boolean( lookup_value_result_2 )->boolean_value
                                   = zcl_xlom__va=>to_boolean( cell_value )->boolean_value ).
                ELSEIF     lookup_value_result_2->type = zif_xlom__va=>c_type-empty
                       AND cell_value->type            = zif_xlom__va=>c_type-empty.
                  equal = abap_true.
                ELSE.
                  RAISE EXCEPTION TYPE zcx_xlom_todo.
                ENDIF.
                IF equal = abap_true.
                  IF lookup_array_result->row_count > 1.
                    temp_result = zcl_xlom__va_number=>create( EXACT #( row_number ) ).
                  ELSE.
                    temp_result = zcl_xlom__va_number=>create( EXACT #( column_number ) ).
                  ENDIF.
                  " Dummy code to exit the two loops
                  row_number = lookup_array_result->row_count.
                  column_number = lookup_array_result->column_count.
                ENDIF.
                column_number = column_number + 1.
              ENDWHILE.
              row_number = row_number + 1.
            ENDWHILE.
          ENDIF.
          IF temp_result IS NOT BOUND.
            " no match found
            temp_result = zcl_xlom__va_error=>na_not_applicable.
          ENDIF.
        ENDIF.
        result = zif_xlom__ex~set_result( temp_result ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
