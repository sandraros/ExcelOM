*"* use this source file for your ABAP unit test classes

CLASS ltc_parser DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS array                          FOR TESTING RAISING cx_static_check.
    METHODS array_two_rows                 FOR TESTING RAISING cx_static_check.
    METHODS function_argument_minus_unary  FOR TESTING RAISING cx_static_check.
    METHODS function_function              FOR TESTING RAISING cx_static_check.
    METHODS function_optional_argument     FOR TESTING RAISING cx_static_check.
    METHODS if                             FOR TESTING RAISING cx_static_check.
    METHODS number                         FOR TESTING RAISING cx_static_check.
    METHODS one_plus_one                   FOR TESTING RAISING cx_static_check.
    METHODS operator_function              FOR TESTING RAISING cx_static_check.
    METHODS operator_function_operator     FOR TESTING RAISING cx_static_check.
    METHODS parentheses_arithmetic         FOR TESTING RAISING cx_static_check.
    METHODS parentheses_arithmetic_complex FOR TESTING RAISING cx_static_check.
    METHODS priority                       FOR TESTING RAISING cx_static_check.


    TYPES tt_token       TYPE zcl_xlom__ex_ut_lexer=>tt_token.
    TYPES ts_result_lexe TYPE zcl_xlom__ex_ut_lexer=>ts_result_lexe.

    CONSTANTS c_type LIKE zcl_xlom__ex_ut_lexer=>c_type VALUE zcl_xlom__ex_ut_lexer=>c_type.

    METHODS assert_equals
      IMPORTING act           TYPE REF TO zif_xlom__ex
                exp           TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

    METHODS lexe
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE tt_token. " ts_result_lexe.

    METHODS parse
      IMPORTING !tokens       TYPE zcl_xlom__ex_ut_lexer=>tt_token
      RETURNING VALUE(result) TYPE REF TO zif_xlom__ex
      RAISING   zcx_xlom__ex_ut_parser.
ENDCLASS.


CLASS ltc_parser IMPLEMENTATION.
  METHOD array.
    assert_equals(
        act = parse( tokens = VALUE #( ( value = `{` type = c_type-curly_bracket_open )
                                       ( value = `1` type = c_type-number )
                                       ( value = `,` type = c_type-comma )
                                       ( value = `2` type = c_type-number )
                                       ( value = `}` type = c_type-curly_bracket_close ) ) )
        exp = zcl_xlom__ex_el_array=>create(
                  rows = VALUE #( ( columns_of_row = VALUE #( ( zcl_xlom__ex_el_number=>create( 1 ) )
                                                              ( zcl_xlom__ex_el_number=>create( 2 ) ) ) ) ) ) ).
  ENDMETHOD.

  METHOD array_two_rows.
    assert_equals(
        act = parse( tokens = VALUE #( ( value = `{` type = c_type-curly_bracket_open )
                                       ( value = `1` type = c_type-number )
                                       ( value = `,` type = c_type-comma )
                                       ( value = `2` type = c_type-number )
                                       ( value = `;` type = c_type-semicolon )
                                       ( value = `3` type = c_type-number )
                                       ( value = `,` type = c_type-comma )
                                       ( value = `4` type = c_type-number )
                                       ( value = `}` type = c_type-curly_bracket_close ) ) )
        exp = zcl_xlom__ex_el_array=>create(
                  rows = VALUE #( ( columns_of_row = VALUE #( ( zcl_xlom__ex_el_number=>create( 1 ) )
                                                              ( zcl_xlom__ex_el_number=>create( 2 ) ) ) )
                                  ( columns_of_row = VALUE #( ( zcl_xlom__ex_el_number=>create( 3 ) )
                                                              ( zcl_xlom__ex_el_number=>create( 4 ) ) ) ) ) ) ).
  ENDMETHOD.

  METHOD assert_equals.
    cl_abap_unit_assert=>assert_true( zcl_xlom__ex_ut=>are_equal( expression_1 = exp
                                                                  expression_2 = act ) ).
  ENDMETHOD.

  METHOD function_argument_minus_unary.
    DATA(act) = parse( tokens = VALUE #( ( value = `OFFSET` type = c_type-function_name )
                                         ( value = `B2`     type = c_type-text_literal )
                                         ( value = `,`      type = c_type-comma )
                                         ( value = `-`      type = c_type-operator )
                                         ( value = `1`      type = c_type-number )
                                         ( value = `,`      type = c_type-comma )
                                         ( value = `-`      type = c_type-operator )
                                         ( value = `1`      type = c_type-number )
                                         ( value = `)`      type = c_type-parenthesis_close ) ) ).
    assert_equals(
        act = act
        exp = zcl_xlom__ex_fu_offset=>create(
                  reference = zcl_xlom__ex_el_string=>create( text = 'B2' )
                  rows      = zcl_xlom__ex_op_minus_unary=>create( operand = zcl_xlom__ex_el_number=>create( 1 ) )
                  cols      = zcl_xlom__ex_op_minus_unary=>create( operand = zcl_xlom__ex_el_number=>create( 1 ) ) ) ).
  ENDMETHOD.

  METHOD function_function.
    DATA(act) = parse( tokens = VALUE #( ( value = `LEN`   type = c_type-function_name )
                                         ( value = `RIGHT` type = c_type-function_name )
                                         ( value = `text`  type = c_type-text_literal )
                                         ( value = `,`     type = c_type-comma )
                                         ( value = `2`     type = c_type-number )
                                         ( value = `)`     type = c_type-parenthesis_close )
                                         ( value = `)`     type = c_type-parenthesis_close ) ) ).
    assert_equals( act = act
                   exp = zcl_xlom__ex_fu_len=>create( text = zcl_xlom__ex_fu_right=>create(
                                                                 text      = zcl_xlom__ex_el_string=>create( 'text' )
                                                                 num_chars = zcl_xlom__ex_el_number=>create( 2 ) ) ) ).
  ENDMETHOD.

  METHOD function_optional_argument.
    DATA(act) = parse( tokens = VALUE #( ( value = `RIGHT` type = c_type-function_name )
                                         ( value = `text`  type = c_type-text_literal )
                                         ( value = `,`     type = c_type-comma )
                                         ( value = `)`     type = c_type-parenthesis_close ) ) ).
    assert_equals( act = act
                   exp = zcl_xlom__ex_fu_right=>create( text      = zcl_xlom__ex_el_string=>create( 'text' )
                                                        num_chars = zcl_xlom__ex_el_empty_arg=>create( ) ) ).
  ENDMETHOD.

  METHOD if.
    assert_equals( act = parse( tokens = VALUE #( ( value = `IF` type = c_type-function_name )
                                                  ( value = `1`  type = c_type-number )
                                                  ( value = `=`  type = c_type-operator )
                                                  ( value = `1`  type = c_type-number )
                                                  ( value = `,`  type = ',' )
                                                  ( value = `0`  type = c_type-number )
                                                  ( value = `,`  type = ',' )
                                                  ( value = `1`  type = c_type-number )
                                                  ( value = `)`  type = ')' ) ) )
                   exp = zcl_xlom__ex_fu_if=>create(
                             condition     = zcl_xlom__ex_op_equal=>create(
                                                 left_operand  = zcl_xlom__ex_el_number=>create( 1 )
                                                 right_operand = zcl_xlom__ex_el_number=>create( 1 ) )
                             expr_if_true  = zcl_xlom__ex_el_number=>create( 0 )
                             expr_if_false = zcl_xlom__ex_el_number=>create( 1 ) ) ).
  ENDMETHOD.

  METHOD lexe.
    DATA(lexer) = zcl_xlom__ex_ut_lexer=>create( ).
    result = lexer->lexe( text ).
  ENDMETHOD.

  METHOD number.
    assert_equals( act = parse( tokens = VALUE #( ( value = `25` type = c_type-number ) ) )
                   exp = zcl_xlom__ex_el_number=>create( 25 ) ).

    assert_equals( act = parse( tokens = VALUE #( ( value = `-` type = c_type-operator )
                                                  ( value = `1` type = c_type-number ) ) )
                   exp = zcl_xlom__ex_op_minus_unary=>create( zcl_xlom__ex_el_number=>create( 1 ) ) ).
  ENDMETHOD.

  METHOD one_plus_one.
    assert_equals( act = parse( tokens = VALUE #( ( value = `1`  type = c_type-number )
                                                  ( value = `+`  type = c_type-operator )
                                                  ( value = `1`  type = c_type-number ) ) )
                   exp = zcl_xlom__ex_op_plus=>create( left_operand  = zcl_xlom__ex_el_number=>create( 1 )
                                                       right_operand = zcl_xlom__ex_el_number=>create( 1 ) ) ).
  ENDMETHOD.

  METHOD operator_function.
    DATA(act) = parse( tokens = VALUE #( ( value = `1`    type = c_type-number )
                                         ( value = `+`    type = c_type-operator )
                                         ( value = `LEN`  type = c_type-function_name )
                                         ( value = `text` type = c_type-text_literal )
                                         ( value = `)`    type = c_type-parenthesis_close ) ) ).
    assert_equals( act = act
                   exp = zcl_xlom__ex_op_plus=>create(
                             left_operand  = zcl_xlom__ex_el_number=>create( 1 )
                             right_operand = zcl_xlom__ex_fu_len=>create(
                                                 text = zcl_xlom__ex_el_string=>create( 'text' ) ) ) ).
  ENDMETHOD.

  METHOD operator_function_operator.
    DATA(act) = parse( tokens = VALUE #( ( value = `1`    type = c_type-number )
                                         ( value = `+`    type = c_type-operator )
                                         ( value = `LEN`  type = c_type-function_name )
                                         ( value = `text` type = c_type-text_literal )
                                         ( value = `)`    type = c_type-parenthesis_close )
                                         ( value = `+`    type = c_type-operator )
                                         ( value = `1`    type = c_type-number ) ) ).
    assert_equals( act = act
                   exp = zcl_xlom__ex_op_plus=>create(
                             left_operand  = zcl_xlom__ex_el_number=>create( 1 )
                             right_operand = zcl_xlom__ex_op_plus=>create(
                                                 left_operand  = zcl_xlom__ex_fu_len=>create(
                                                                     text = zcl_xlom__ex_el_string=>create( 'text' ) )
                                                 right_operand = zcl_xlom__ex_el_number=>create( 1 ) ) ) ).
  ENDMETHOD.

  METHOD parentheses_arithmetic.
    " lexe( '2*(1+3)' )
    DATA(act) = parse( VALUE #( ( value = `2`  type = c_type-number )
                                ( value = `*`  type = c_type-operator )
                                ( value = `(`  type = c_type-parenthesis_open )
                                ( value = `1`  type = c_type-number )
                                ( value = `+`  type = c_type-operator )
                                ( value = `3`  type = c_type-number )
                                ( value = `)`  type = c_type-parenthesis_close ) ) ).
    DATA(exp) = zcl_xlom__ex_op_mult=>create( left_operand  = zcl_xlom__ex_el_number=>create( 2 )
                                              right_operand = zcl_xlom__ex_op_plus=>create(
                                                  left_operand  = zcl_xlom__ex_el_number=>create( 1 )
                                                  right_operand = zcl_xlom__ex_el_number=>create( 3 ) ) ).
    assert_equals( act = act
                   exp = exp ).
  ENDMETHOD.

  METHOD parentheses_arithmetic_complex.
    " lexe( '2*(1+3*(5+1))' )
    DATA(act) = parse( tokens = VALUE #( ( value = `2`  type = c_type-number )
                                         ( value = `*`  type = c_type-operator )
                                         ( value = `(`  type = c_type-parenthesis_open )
                                         ( value = `1`  type = c_type-number )
                                         ( value = `+`  type = c_type-operator )
                                         ( value = `3`  type = c_type-number )
                                         ( value = `*`  type = c_type-operator )
                                         ( value = `(`  type = c_type-parenthesis_open )
                                         ( value = `5`  type = c_type-number )
                                         ( value = `+`  type = c_type-operator )
                                         ( value = `1`  type = c_type-number )
                                         ( value = `)`  type = c_type-parenthesis_close )
                                         ( value = `)`  type = c_type-parenthesis_close ) ) ).
    DATA(exp) = zcl_xlom__ex_op_mult=>create(
                    left_operand  = zcl_xlom__ex_el_number=>create( 2 )
                    right_operand = zcl_xlom__ex_op_plus=>create(
                        left_operand  = zcl_xlom__ex_el_number=>create( 1 )
                        right_operand = zcl_xlom__ex_op_mult=>create(
                                            left_operand  = zcl_xlom__ex_el_number=>create( 3 )
                                            right_operand = zcl_xlom__ex_op_plus=>create(
                                                left_operand  = zcl_xlom__ex_el_number=>create( 5 )
                                                right_operand = zcl_xlom__ex_el_number=>create( 1 ) ) ) ) ).
    assert_equals( act = act
                   exp = exp ).
  ENDMETHOD.

  METHOD parse.
    result = zcl_xlom__ex_ut_parser=>create( )->parse( tokens ).
  ENDMETHOD.

  METHOD priority.
    " lexe( '1+2*3' )
    DATA(act) = parse( VALUE #( ( value = `1`  type = c_type-number )
                                ( value = `+`  type = c_type-operator )
                                ( value = `2`  type = c_type-number )
                                ( value = `*`  type = c_type-operator )
                                ( value = `3`  type = c_type-number ) ) ).
    DATA(exp) = zcl_xlom__ex_op_plus=>create( left_operand  = zcl_xlom__ex_el_number=>create( 1 )
                                              right_operand = zcl_xlom__ex_op_mult=>create(
                                                  left_operand  = zcl_xlom__ex_el_number=>create( 2 )
                                                  right_operand = zcl_xlom__ex_el_number=>create( 3 ) ) ).
    assert_equals( act = act
                   exp = exp ).
  ENDMETHOD.
ENDCLASS.
