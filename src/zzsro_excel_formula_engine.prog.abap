REPORT zzsro_excel_formula_engine.

CLASS lcl_excelom DEFINITION DEFERRED.
CLASS LCL_EXCELOM_FORMULA2 DEFINITION DEFERRED.
CLASS LCL_EXCELOM_RANGE DEFINITION DEFERRED.
CLASS LCL_EXCELOM_RANGE_VALUE DEFINITION DEFERRED.
CLASS lcl_excelom_tools DEFINITION DEFERRED.
CLASS lcl_excelom_workbook DEFINITION DEFERRED.
CLASS lcl_excelom_workbooks DEFINITION DEFERRED.
CLASS lcl_excelom_worksheet DEFINITION DEFERRED.
CLASS lcl_excelom_worksheets DEFINITION DEFERRED.


CLASS lcx_to_do DEFINITION INHERITING FROM cx_no_check.
ENDCLASS.


CLASS lcx_unexpected DEFINITION INHERITING FROM cx_no_check.
ENDCLASS.


CLASS lcx_parser DEFINITION INHERITING FROM cx_static_check.
  PUBLIC SECTION.
    METHODS constructor
      IMPORTING !text     TYPE csequence OPTIONAL
                msgv1     TYPE csequence OPTIONAL
                msgv2     TYPE csequence OPTIONAL
                msgv3     TYPE csequence OPTIONAL
                msgv4     TYPE csequence OPTIONAL
                textid    LIKE textid    OPTIONAL
                !previous LIKE previous  OPTIONAL.

    METHODS get_text     REDEFINITION.
    METHODS get_longtext REDEFINITION.

  PRIVATE SECTION.
    DATA text  TYPE string.
    DATA msgv1 TYPE string.
    DATA msgv2 TYPE string.
    DATA msgv3 TYPE string.
    DATA msgv4 TYPE string.
ENDCLASS.


CLASS lcx_parser IMPLEMENTATION.
  METHOD constructor ##ADT_SUPPRESS_GENERATION.
    super->constructor( previous = previous
                        textid   = textid ).
    me->text  = text.
    me->msgv1 = msgv1.
    me->msgv2 = msgv2.
    me->msgv3 = msgv3.
    me->msgv4 = msgv4.
  ENDMETHOD.

  METHOD get_longtext.
    IF text IS NOT INITIAL.
      result = get_text( ).
    ELSE.
      result = super->get_longtext( ).
    ENDIF.
  ENDMETHOD.

  METHOD get_text.
    IF text IS NOT INITIAL.
      result = text.
      REPLACE '&1' IN result WITH msgv1.
      REPLACE '&2' IN result WITH msgv2.
      REPLACE '&3' IN result WITH msgv3.
      REPLACE '&4' IN result WITH msgv4.
    ELSE.
      result = super->get_text( ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.


INTERFACE lif_expression.
  METHODS evaluate RETURNING VALUE(result) TYPE i.
ENDINTERFACE.


CLASS lcl_expressions DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    METHODS append IMPORTING expression TYPE REF TO lif_expression.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_expressions.
ENDCLASS.


CLASS lcl_function_call DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_expression.

    CLASS-METHODS create
      IMPORTING !name         TYPE csequence
                arguments     TYPE REF TO lcl_expressions
      RETURNING VALUE(result) TYPE REF TO lcl_function_call.
ENDCLASS.


CLASS lcl_sub_expression DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_expression.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_sub_expression.
ENDCLASS.


CLASS lcl_table_expression DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_expression.

    CLASS-METHODS create
      IMPORTING table_name            TYPE csequence
                row_column_specifiers TYPE string_table
      RETURNING VALUE(result)         TYPE REF TO lcl_table_expression.
ENDCLASS.


CLASS lcl_array DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_expression.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_array.
ENDCLASS.


CLASS lcl_numeric_literal DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_expression.

    CLASS-METHODS create
      IMPORTING !number       TYPE f
      RETURNING VALUE(result) TYPE REF TO lcl_numeric_literal.

  PRIVATE SECTION.
    DATA number TYPE f.
ENDCLASS.


INTERFACE lif_operator.
  METHODS get_priority RETURNING VALUE(result) TYPE i.
  "! 1 : preceding operand only (%)
  "! 2 : before and after operand only (+ - * / ^ &)
  "! 3 : succeding operand only (unary + and -)
  "!
  "! @parameter result | /
  METHODS get_operand_positions RETURNING VALUE(result) TYPE i.

  METHODS set_operands IMPORTING !preceding TYPE REF TO lif_expression OPTIONAL
                                 succeding  TYPE REF TO lif_expression OPTIONAL.

  " Priorities:
  "  1  : (colon)  (single space) , (comma) Reference operators
  "  2  – Negation (as in –1)
  "  3  % Percent
  "  4  ^ Exponentiation
  "  5  * and / Multiplication and division
  "  6  + and – Addition and subtraction
  "  7  & Connects two strings of text (concatenation)
  "  8  = < > <= >= <> Comparison
ENDINTERFACE.


CLASS lcl_plus DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_expression.
    INTERFACES lif_operator.

    CLASS-METHODS create
      IMPORTING left_operand type ref to lif_expression
      right_operand type ref to lif_expression
      RETURNING VALUE(result) TYPE REF TO lcl_plus.

PRIVATE SECTION.
DATA left_operand type ref to lif_expression.
      DATA right_operand type ref to lif_expression.
ENDCLASS.


CLASS lcl_string_literal DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_expression.

    CLASS-METHODS create
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE REF TO lcl_string_literal.
ENDCLASS.


CLASS lcl_numeric_literal IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_numeric_literal( ).
    result->number = number.
  ENDMETHOD.

  METHOD lif_expression~evaluate.
    result = number.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_plus IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_plus( ).
    result->left_operand = left_operand.
    result->right_operand = right_operand.
  ENDMETHOD.

  METHOD lif_expression~evaluate.
    result = left_operand->evaluate( ) + right_operand->evaluate( ).
  ENDMETHOD.

  METHOD lif_operator~get_priority.
    result = 6.
  ENDMETHOD.

  METHOD lif_operator~get_operand_positions.
  ENDMETHOD.

  METHOD lif_operator~set_operands.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_string_literal IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.

  METHOD lif_expression~evaluate.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_sub_expression IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.

  METHOD lif_expression~evaluate.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_table_expression IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.

  METHOD lif_expression~evaluate.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_array IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.

  METHOD lif_expression~evaluate.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_function_call IMPLEMENTATION.
  METHOD create.
    " TODO: parameter NAME is never used (ABAP cleaner)
    " TODO: parameter ARGUMENTS is never used (ABAP cleaner)

    result = NEW lcl_function_call( ).
  ENDMETHOD.

  METHOD lif_expression~evaluate.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_expressions IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_expressions( ).
  ENDMETHOD.

  METHOD append.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_lexer DEFINITION
    FINAL
    CREATE PRIVATE.
  PUBLIC SECTION.
    TYPES:
      BEGIN OF ts_token,
        value TYPE string,
        type  TYPE c LENGTH 1,
      END OF ts_token.
    TYPES tT_token TYPE STANDARD TABLE OF ts_token WITH EMPTY KEY.

    CLASS-METHODS create
      RETURNING
        VALUE(result) TYPE REF TO lcl_lexer.
    METHODS lexe IMPORTING !text         TYPE csequence
                 RETURNING VALUE(result) TYPE TT_token.
PRIVATE SECTION.
    METHODS complete_with_non_matches
      IMPORTING i_string  TYPE string
      CHANGING  c_matches TYPE match_result_tab.
ENDCLASS.

CLASS lcl_lexer IMPLEMENTATION.
  METHOD complete_with_non_matches.
    DATA(last_offset) = 0.
    LOOP AT c_matches ASSIGNING FIELD-SYMBOL(<match>).
      IF <match>-offset > last_offset.
        INSERT VALUE match_result( offset = last_offset
                                   length = <match>-offset - last_offset ) INTO c_matches INDEX sy-tabix.
      ENDIF.
      last_offset = <match>-offset + <match>-length.
    ENDLOOP.
    IF strlen( i_string ) > last_offset.
      APPEND VALUE match_result( offset = last_offset
                                 length = strlen( i_string ) - last_offset ) TO c_matches.
    ENDIF.
  ENDMETHOD.

  METHOD create.
    result = NEW lcl_lexer( ).
  ENDMETHOD.
  METHOD lexe.
    " Note about `[ ` and ` ]` (https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e):
    "   > Use the space character to improve readability in a structured reference
    "   > You can use space characters to improve the readability of a structured reference.
    "   > For example: =DeptSales[ [Sales Person]:[Region] ] or =DeptSales[[#Headers], [#Data], [% Commission]]"
    "   > It’s recommended to use one space:
    "   >   - After the first left bracket ([)
    "   >   - Preceding the last right bracket (]).
    "   >   - After a comma.
    "
    " Between `[` and `]`, the escape character is `'` e.g. `['[value']]` for the column header `[value]`.
    "
    " Note: -- is not an operator, it's a chain of the unary "-" operator (there could be even 3 or more subsequent unary operators); + can also be a unary operator,
    "       hence the formula +--++-1 is a valid formula which simply means -1. https://stackoverflow.com/questions/3286197/what-does-do-in-excel-formulas
    FIND ALL OCCURRENCES OF REGEX '(?:\(|\{|\[(?:''[\[\]]|[^\[\]])+\]|\[ ?|\)|\}| ?\]|, ?|;| |:|<>|<=|>=|<|>|=|\+|-|\*|/|\^|&|%|"(?:""|[^"])*")' IN text RESULTS DATA(matches).
    complete_with_non_matches( EXPORTING i_string  = text
                               CHANGING  c_matches = matches ).
    DATA(token) = VALUE ts_token( ).
    LOOP AT matches REFERENCE INTO DATA(match).
      DATA(token_value) = substring( val = text
                                     off = match->offset
                                     len = match->length ).
      " is comma a separator or a union operator?
      " https://techcommunity.microsoft.com/t5/excel/does-the-union-operator-exist/m-p/2590110
      " With argument-list functions, there is no union. Example: A1 contains 1, both =SUM(A1,A1) and =SUM((A1,A1)) return 2.
      " With no-argument-list functions, there is a union. Example: =LARGE((A1,B1),2) (=LARGE(A1,B1,2) is invalid, too many arguments)
      CASE token_value.
        WHEN '('
          OR '[]'
          OR '['
          OR `[ `
          OR '{'
          OR ')'
          OR '}'
          OR ']'
          OR ` ]`
          OR ',' " separator or union operator?
          OR `, `
          OR ';'.
          token = VALUE #( value = condense( token_value )
                           type  = condense( token_value ) ).
        WHEN ` `
*          OR ','
          OR ':' " =B1:A1:B2:B3:A1:B2:B2:B3:B2 is same as =A1:B3
          OR '<>'
          OR '<='
          OR '>='
          OR '<'
          OR '>'
          OR '='
          OR '+'
          OR '-'
          OR '*'
          OR '/'  " 10/2 = 5
          OR '^'  " 10^2 = 100
          OR '&'  " "A"&"B" = "AB"
          OR '%'. " 10% = 0.1
          token = VALUE #( value = token_value
                           type  = 'O' ).
        WHEN OTHERS.
          IF substring( val = token_value
                        len = 1 ) = '"'.
            " text literal
            token = VALUE #( value = token_value
                             type  = '"' ).
          ELSEIF substring( val = token_value
                            len = 1 ) = '['.
            " table argument
            token = VALUE #( value = token_value
                             type  = '[' ).
          ELSE.
            " function name, --, cell reference, table name, name of named range, constant (TRUE, FALSE, number)
            token = VALUE #( value = token_value
                             type  = 'W' ).
          ENDIF.
      ENDCASE.
      APPEND token TO result.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.


CLASS lcl_parser DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS create
*      IMPORTING formula       TYPE csequence
      RETURNING VALUE(result) TYPE REF TO lcl_parser.

    METHODS parse IMPORTING lexer_tokens  TYPE lcl_lexer=>TT_token
                  RETURNING VALUE(result) TYPE REF TO lif_expression
      RAISING lcx_parser.

  PRIVATE SECTION.
*    DATA formula        TYPE string.
    DATA formula_offset TYPE i.
    DATA current_token_index TYPE sytabix.
    DATA tokens              TYPE lcl_lexer=>tt_token.

    METHODS get_token
      RETURNING VALUE(result) TYPE string.

    METHODS parse_expression         RETURNING VALUE(result) TYPE REF TO lif_expression.

    METHODS parse_function_arguments RETURNING VALUE(result) TYPE REF TO lcl_expressions.

    METHODS parse_tokens_up_to IMPORTING stop_at_token TYPE csequence
                               RETURNING VALUE(result) TYPE string_table.

    METHODS skip_spaces.

ENDCLASS.


CLASS lcl_parser IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_parser( ).
*    result->formula = formula.
  ENDMETHOD.

  METHOD get_token.
  ENDMETHOD.

  METHOD parse.
    current_token_index = 1.
    tokens = lexer_tokens.
    result = parse_expression( ).
  ENDMETHOD.

  METHOD parse_expression.
    TYPES:
      BEGIN OF ts_expression_part,
        s_token      TYPE REF TO lcl_lexer=>ts_token,
        o_expression TYPE REF TO lif_expression,
      END OF ts_expression_part.
    TYPES tt_expression_part TYPE STANDARD TABLE OF ts_expression_part WITH EMPTY KEY.
    TYPES to_expression      TYPE REF TO lif_expression.

    DATA(table_expressions) = VALUE tt_expression_part( ).
    " TODO: variable is assigned but only used in commented-out code (ABAP cleaner)
    DATA(start_token_index) = current_token_index.
    WHILE current_token_index <= lines( tokens ).
      DATA(token) = REF #( tokens[ current_token_index ] OPTIONAL ).
      DATA(expression) = VALUE to_expression( ).
      CASE token->type.
        WHEN '('.
          " either a sub-expression like (1+1)
          " or a union of ranges like (A1:A2,B1:B2) (which is equivalent to A1:B2)
          " or an intersection of ranges like (A1:B2 B2:C3) (which is equivalent to B2)
*            data sub_expressions type STANDARD TABLE OF ref to lif_expression with EMPTY KEY.
*            while current_token_index < lines( tokens ).
*            current_token_index = current_token_index + 1.
*        if tokens[ current_token_index ]-type = ')'.
*          current_token_index = current_token_index + 1.
*          exit.
*        endif.
*            data(sub_expression) = parse_expression( ).
*            append sub_expression to sub_expressions.
*            ENDWHILE.
*          append lcl_sub_expression=>create( ) to table_expressions.
        WHEN '['.
          " table1[
          " should not happen because it's processed with the previous word
          RAISE EXCEPTION TYPE lcx_unexpected.
        WHEN '{'.
          RAISE EXCEPTION TYPE lcx_to_do.
        WHEN ')'.
          RAISE EXCEPTION TYPE lcx_to_do.
        WHEN ']'.
          RAISE EXCEPTION TYPE lcx_to_do.
        WHEN '}'.
          RAISE EXCEPTION TYPE lcx_to_do.
*          result = lcl_array=>create( ).
        WHEN ` `.
          RAISE EXCEPTION TYPE lcx_to_do.
        WHEN ':'.
          RAISE EXCEPTION TYPE lcx_to_do.
        WHEN ','.
          RAISE EXCEPTION TYPE lcx_to_do.
          " end of expression
          EXIT.
        WHEN ';'.
          RAISE EXCEPTION TYPE lcx_to_do.
          " end of expression
          EXIT.
        WHEN 'O'.
          " operator
          expression = VALUE #( ).
          current_token_index = current_token_index + 1.
*          CASE token->value.
*            WHEN '+'.
*              IF current_token_index = start_token_index OR tokens[ current_token_index - 1 ]-type <> 'W'.
*                " append lcl_unary_plus=>create( ) to table_expressions.
*              ELSE.
*                APPEND value #( lcl_plus=>create( ) TO table_expressions.
*              ENDIF.
*          ENDCASE.
        WHEN 'W'.
          " word
          IF current_token_index < lines( tokens ) AND tokens[ current_token_index + 1 ]-type = '('.
            " The word is a function name
            current_token_index = current_token_index + 1.
            DATA(arguments) = parse_function_arguments( ).
            expression = lcl_function_call=>create( name      = token->value
                                                    arguments = arguments ).
            " result = function.
          ELSEIF current_token_index < lines( tokens ) AND tokens[ current_token_index + 1 ]-type = '['.
            " The word is a table name e.g. Table1[] (means Table1[#Data])
            DATA(row_column_specifiers) = VALUE string_table( ).
            LOOP AT tokens REFERENCE INTO DATA(token_2)
                 FROM current_token_index.
              current_token_index = current_token_index + 1.
              CASE token_2->type.
                WHEN '['.
                  IF token_2->value <> '['.
                    APPEND token_2->value TO row_column_specifiers.
                  ENDIF.
                WHEN ']'.
                  EXIT.
              ENDCASE.
            ENDLOOP.
            expression = lcl_table_expression=>create( table_name            = token->value
                                                       row_column_specifiers = row_column_specifiers ).
          ELSEIF substring( val = token->value
                            len = 1 ) CA '0123456789'.
            " The word is a number.
            " NB: the number 0.5 is represented in the formulas with the leading 0 (i.e. "0.5")
            current_token_index = current_token_index + 1.
            expression = lcl_numeric_literal=>create( EXACT #( token->value ) ).
          ELSE.
            " The word is a cell reference, name of named range, constant (TRUE, FALSE)
            " TODO
          ENDIF.
        WHEN '"'.
          " Remove double quotes e.g. "say ""hello""" -> say "hello"
          current_token_index = current_token_index + 1.
          expression = lcl_string_literal=>create( replace( val   = token->value
                                                            regex = '^"|(")"|"$'
                                                            with  = '$1'
                                                            occ   = 0 ) ).
        WHEN OTHERS.
          RAISE EXCEPTION TYPE lcx_unexpected.
      ENDCASE.
      " Here, EXPRESSION maybe unbound like e.g. concerning operators.
      APPEND VALUE #( s_token      = token
                      o_expression = expression )
             TO table_expressions.
    ENDWHILE.
    IF lines( table_expressions ) = 1.
      result = table_expressions[ 1 ]-o_expression.
    ELSE.
      " operator precedence
      " Get operator priorities
      TYPES:
        BEGIN OF ts_operator,
          name     TYPE string,
          "! To distinguish unary from binary operators + and -
          unary    TYPE abap_bool,
          priority TYPE i,
          "! % is the only postfix operator e.g. 10% (=0.1)
          postfix  TYPE abap_bool,
          desc     TYPE string,
        END OF ts_operator.
      TYPES tt_operator TYPE SORTED TABLE OF ts_operator WITH UNIQUE KEY name unary.
      TYPES:
        BEGIN OF ts_work,
          position          TYPE sytabix,
          operator          TYPE string,
          operator_expression TYPE REF TO lif_expression,
          preceding_operand TYPE ts_expression_part,
          succeding_operand TYPE ts_expression_part ,
          priority          TYPE i,
        END OF ts_work.
      TYPES tt_work TYPE STANDARD TABLE OF ts_work WITH EMPTY KEY.
      " Priorities:
      "  1  : (colon)  (single space) , (comma) Reference operators
      "  2  – Negation (as in –1) + (as in +1)
      "  3  % Percent (as in =50%)
      "  4  ^ Exponentiation
      "  5  * and / Multiplication and division
      "  6  + and – Addition and subtraction
      "  7  & Connects two strings of text (concatenation)
      "  8  = < > <= >= <> Comparison
      DATA(operators) = VALUE tt_operator( ( name = ':'              priority = 1  desc = 'range A1:A2 or A1:A2:A2' )
                                           ( name = ` `              priority = 1  desc = 'intersection A1 A2' )
                                           ( name = ','              priority = 1  desc = 'union A1,A2' )
                                           ( name = '-'  unary = 'X' priority = 2  desc = '-1' )
                                           ( name = '+'  unary = 'X' priority = 2  desc = '+1' )
                                           ( name = '%'  unary = 'X' priority = 3  desc = 'percent' postfix = 'X' )
                                           ( name = '^'              priority = 4  desc = 'exponent 2^8' )
                                           ( name = '*'              priority = 5  desc = '2*2' )
                                           ( name = '/'              priority = 5  desc = '2/2' )
                                           ( name = '+'              priority = 6  desc = '2+2' )
                                           ( name = '-'              priority = 6  desc = '2-2' )
                                           ( name = '&'              priority = 7  desc = 'concatenate "A"&"B"' )
                                           ( name = '='              priority = 8  desc = 'A1=1' )
                                           ( name = '<'              priority = 8  desc = 'A1<1' )
                                           ( name = '>'              priority = 8  desc = 'A1>1' )
                                           ( name = '<='             priority = 8  desc = 'A1<=1' )
                                           ( name = '>='             priority = 8  desc = 'A1>=1' )
                                           ( name = '<>'             priority = 8  desc = 'A1<>1' ) ).
      DATA(work_table) = VALUE tt_work( ).
      LOOP AT table_expressions REFERENCE INTO DATA(expression_part)
           WHERE s_token->type = 'O'.
*        TRY.
*            DATA(operator) = CAST lif_operator( expression_part->o_expression ).
*          CATCH cx_sy_move_cast_error.
*            CONTINUE.
*        ENDTRY.
        INSERT VALUE #( position            = sy-tabix
                        operator            = expression_part->s_token->value
                        operator_expression = VALUE #( )
                        preceding_operand   = COND #( WHEN sy-tabix >= 2
                                                      THEN VALUE #( table_expressions[ sy-tabix - 1 ] OPTIONAL ) )
                        succeding_operand   = COND #( WHEN sy-tabix < lines( table_expressions )
                                                      THEN VALUE #( table_expressions[ sy-tabix + 1 ] OPTIONAL ) )
                        priority            = operators[ name = expression_part->s_token->value ]-priority )
               INTO TABLE work_table.
      ENDLOOP.
      SORT work_table BY priority DESCENDING
                         position.
      " Determine operator expressions
      LOOP AT work_table REFERENCE INTO DATA(work_line)
           GROUP BY work_line->priority REFERENCE INTO DATA(group_by_priority).
        DATA(buffer_expression) = value to_expression( ).
        LOOP AT GROUP group_by_priority REFERENCE INTO work_line.
          if buffer_expression is not bound.
            buffer_expression = work_line->preceding_operand-o_expression.
          endif.
          CASE work_line->operator.
            WHEN '+'.
              work_line->operator_expression = lcl_plus=>create(
                                                   left_operand  = buffer_expression
                                                   right_operand = work_line->succeding_operand-o_expression ).
          ENDCASE.
          buffer_expression = work_line->operator_expression.
        ENDLOOP.
      ENDLOOP.
      result = buffer_expression.
    ENDIF.
  ENDMETHOD.

  METHOD parse_function_arguments.
    result = lcl_expressions=>create( ).
    DO.
      DATA(expression) = parse_expression( ).
      result->append( expression ).
      CASE tokens[ current_token_index ]-type.
        WHEN ','.
          current_token_index = current_token_index + 1.
        WHEN ')'.
          current_token_index = current_token_index + 1.
          EXIT.
        WHEN OTHERS.
          RAISE EXCEPTION TYPE lcx_unexpected.
      ENDCASE.
    ENDDO.
  ENDMETHOD.

  METHOD parse_tokens_up_to.
  ENDMETHOD.

  METHOD skip_spaces.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_tools DEFINITION FINAL.

  PUBLIC SECTION.
    CLASS-METHODS type
      IMPORTING any_data_object TYPE any
      RETURNING VALUE(result)   TYPE abap_typekind.

ENDCLASS.


CLASS lcl_excelom_tools IMPLEMENTATION.
  METHOD type.
    DESCRIBE FIELD any_data_object TYPE result.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_formula2 DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS create RETURNING value(result) TYPE REF TO lcl_excelom_formula2.
    METHODS set importing value type string.
ENDCLASS.


CLASS lcl_excelom_formula2 IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_formula2( ).
  ENDMETHOD.
  METHOD set.

  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_range_value DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range_value.

    METHODS set_double importing value type f.
    METHODS set_string importing value type string.
ENDCLASS.


CLASS lcl_excelom_range_value IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_range_value( ).
  ENDMETHOD.
  METHOD set_double.

  ENDMETHOD.

  METHOD set_string.

  ENDMETHOD.

ENDCLASS.


CLASS lcl_excelom_range DEFINITION.
  PUBLIC SECTION.
    "!
    "! @parameter cell1 | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2 | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    CLASS-METHODS create
      IMPORTING cell1         TYPE any
                cell2         TYPE any OPTIONAL
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.

    METHODS value    RETURNING VALUE(result) TYPE REF TO lcl_excelom_range_value.
    METHODS formula2 RETURNING VALUE(result) TYPE REF TO lcl_excelom_formula2.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ty_address_one_cell,
        column TYPE i,
        row    TYPE i,
      END OF ty_address_one_cell.
    TYPES:
      BEGIN OF ty_address,
        top_left     TYPE ty_address_one_cell,
        bottom_right TYPE ty_address_one_cell,
      END OF ty_address.

    CLASS-METHODS decode_range_address
      IMPORTING address       TYPE string
      RETURNING VALUE(result) TYPE ty_address.

    DATA _formula2 TYPE REF TO lcl_excelom_formula2.
    DATA _address  TYPE ty_address.
ENDCLASS.


CLASS lcl_excelom_range IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_range( ).
      if cell2 is initial.
      result->_address = decode_range_address( cell1 ).
    else.
  endif.
  ENDMETHOD.

  METHOD decode_range_address.

  ENDMETHOD.

  METHOD formula2.
    IF _formula2 IS NOT BOUND.
      _formula2 = lcl_excelom_formula2=>create( ).
    ELSE.
      result = _formula2.
    ENDIF.
  ENDMETHOD.

  METHOD value.
  ENDMETHOD.

ENDCLASS.


CLASS lcl_excelom_worksheet DEFINITION.
  PUBLIC SECTION.
    TYPES ty_name TYPE string.

    CLASS-METHODS create RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

    "!
    "! @parameter cell1 | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2 | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    METHODS range
      IMPORTING cell1         TYPE any
                cell2         TYPE any optional
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.
    METHODS calculate.
  PRIVATE SECTION.
    DATA formulas TYPE STANDARD TABLE OF REF TO lif_expression WITH EMPTY KEY.
ENDCLASS.


CLASS lcl_excelom_worksheet IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_worksheet( ).
  ENDMETHOD.
  METHOD range.
    result = lcl_excelom_range=>create( cell1 = cell1 cell2 = cell2 ).
  ENDMETHOD.
  METHOD calculate.
    LOOP AT formulas INTO DATA(formula).
      formula->evaluate( ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_worksheets DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheets.

    METHODS add
      IMPORTING !name         TYPE lcl_excelom_worksheet=>ty_name
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

    METHODS count
      RETURNING VALUE(result) TYPE i.

    "!
    "! @parameter index | Required    Variant The name or index number of the object.
    "! @parameter result | .
    METHODS item
      IMPORTING index         TYPE simple
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

  PRIVATE SECTION.
      types:
      begin OF ty_worksheet,
        name   TYPE lcl_excelom_worksheet=>ty_name,
        object TYPE REF TO lcl_excelom_worksheet,
      END OF ty_worksheet.
    TYPES ty_worksheets TYPE SORTED TABLE OF ty_worksheet WITH UNIQUE KEY name.

    DATA worksheets TYPE ty_worksheets.
ENDCLASS.


CLASS lcl_excelom_worksheets IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_worksheets( ).
  ENDMETHOD.

  METHOD add.
    DATA worksheet TYPE ty_worksheet.

    worksheet-name   = name.
    worksheet-object = lcl_excelom_worksheet=>create( ).
    INSERT worksheet INTO TABLE worksheets.
    result = worksheet-object.
  ENDMETHOD.

  METHOD count.
    result = lines( worksheets ).
  ENDMETHOD.

  METHOD item.
    CASE lcl_excelom_tools=>type( index ).
      WHEN cl_abap_typedescr=>typekind_string
        OR cl_abap_typedescr=>typekind_char.
        result = worksheets[ name = index ]-object.
      WHEN cl_abap_typedescr=>typekind_int.
        result = worksheets[ index ]-object.
      WHEN OTHERS.
        " TODO
    ENDCASE.
  ENDMETHOD.

ENDCLASS.


CLASS lcl_excelom_workbook DEFINITION.
  PUBLIC SECTION.
    TYPES ty_name TYPE string.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbook.

    METHODS worksheets RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheets.

  PRIVATE SECTION.
    DATA _worksheets TYPE REF TO lcl_excelom_worksheets.
ENDCLASS.


CLASS lcl_excelom_workbook IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_workbook( ).
    result->_worksheets = lcl_excelom_worksheets=>create( ).
    result->_worksheets->add( name = 'Sheet1' ).
  ENDMETHOD.

  METHOD worksheets.
    result = _worksheets.
  ENDMETHOD.

ENDCLASS.


CLASS lcl_excelom_workbooks DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbooks.

    METHODS add
      IMPORTING !name         TYPE lcl_excelom_workbook=>ty_name
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbook.

    METHODS count
      RETURNING VALUE(result) TYPE i.

    "!
    "! @parameter index | Required    Variant The name or index number of the object.
    "! @parameter result | .
    METHODS item
      IMPORTING index         TYPE simple
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbook.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ty_workbook,
        name   TYPE lcl_excelom_workbook=>ty_name,
        object TYPE REF TO lcl_excelom_workbook,
      END OF ty_workbook.
    TYPES ty_workbooks TYPE SORTED TABLE OF ty_workbook WITH UNIQUE KEY name.

    DATA workbooks TYPE ty_workbooks.
ENDCLASS.


CLASS lcl_excelom_workbooks IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_workbooks( ).
  ENDMETHOD.

  METHOD add.
    DATA workbook TYPE ty_workbook.

    workbook-name   = name.
    workbook-object = lcl_excelom_workbook=>create( ).
    INSERT workbook INTO TABLE workbooks.
    result = workbook-object.
  ENDMETHOD.

  METHOD count.
    result = lines( workbooks ).
  ENDMETHOD.

  METHOD item.
    CASE lcl_excelom_tools=>type( index ).
      WHEN cl_abap_typedescr=>typekind_string.
        result = workbooks[ name = index ]-object.
      WHEN cl_abap_typedescr=>typekind_int.
        result = workbooks[ index ]-object.
      WHEN OTHERS.
        " TODO
    ENDCASE.
  ENDMETHOD.

ENDCLASS.


CLASS lcl_excelom DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS create RETURNING VALUE(result) TYPE REF TO lcl_excelom.
    METHODS workbooks RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbooks.
    METHODS calculate.

  PRIVATE SECTION.
    DATA _workbooks TYPE REF TO lcl_excelom_workbooks.
ENDCLASS.


CLASS lcl_excelom IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom( ).
    result->_workbooks = lcl_excelom_workbooks=>create( ).
  ENDMETHOD.

  METHOD calculate.
    DATA(workbook_number) = 1.
    WHILE workbook_number <= _workbooks->count( ).
      DATA(workbook) = _workbooks->item( workbook_number ).

      DATA(worksheet_number) = 1.
      WHILE worksheet_number <= workbook->worksheets( )->count( ).
        DATA(worksheet) = workbook->worksheets( )->item( worksheet_number ).
        worksheet->calculate( ).
      ENDWHILE.
    ENDWHILE.
  ENDMETHOD.

  METHOD workbooks.
    RESULT = _workbooks.
  ENDMETHOD.

ENDCLASS.


CLASS ltc_parser DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS test  FOR TESTING RAISING cx_static_check.
    METHODS test2 FOR TESTING RAISING cx_static_check.
    METHODS test3 FOR TESTING RAISING cx_static_check.
    METHODS test4 FOR TESTING RAISING cx_static_check.
    METHODS test31 FOR TESTING RAISING cx_static_check.

  types tt_token TYPE lcl_lexer=>tt_token.

    METHODS lexe
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE lcl_lexer=>tt_token.

    METHODS parse IMPORTING lexer_tokens  TYPE lcl_lexer=>TT_token
                  RETURNING VALUE(result) TYPE REF TO lif_expression.

    METHODS evaluate IMPORTING expression    TYPE REF TO lif_expression
                     RETURNING VALUE(result) TYPE i.

    METHODS get_texts_from_matches IMPORTING i_string      TYPE string
                                             i_matches     TYPE match_result_tab
                                   RETURNING VALUE(result) TYPE string_table.

ENDCLASS.


CLASS ltc_parser IMPLEMENTATION.
  METHOD evaluate.
    result = expression->evaluate( ).
  ENDMETHOD.

  METHOD get_texts_from_matches.
    LOOP AT i_matches REFERENCE INTO DATA(match).
      APPEND substring( val = i_string
                        off = match->offset
                        len = match->length ) TO result.
    ENDLOOP.
  ENDMETHOD.

  METHOD lexe.
    data(lexer) = lcl_lexer=>create( ).
    result = lexer->lexe( text ).
  ENDMETHOD.

  METHOD parse.
    data(parser) = lcl_parser=>create( ).
    result = parser->parse( lexer_tokens ).
  ENDMETHOD.

  METHOD test.
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `IF` type = 'W' )
                                                              ( value = `(`  type = '(' )
                                                              ( value = `1`  type = 'W' )
                                                              ( value = `=`  type = 'O' )
                                                              ( value = `1`  type = 'W' )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `0`  type = 'W' )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `1`  type = 'W' )
                                                              ( value = `)`  type = ')' ) )
                                        act = lexe( 'IF(1=1,0,1)' ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `Sheet1!$A$1` type = 'W' ) )
                                        act = lexe( 'Sheet1!$A$1' ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `"IF(1=1,0,1)"` type = '"' ) )
                                        act = lexe( '"IF(1=1,0,1)"' ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `"IF(A1=""X"",0,1)"` type = '"' ) )
                                        act = lexe( '"IF(A1=""X"",0,1)"' ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `(` type = '(' )
                                                              ( value = `a` type = 'W' )
                                                              ( LINES OF VALUE
                                                                tt_token( FOR i = 1 WHILE i <= 5000
                                                                          ( value = `,` type = ',' )
                                                                          ( value = `a` type = 'W' ) ) )
                                                              ( value = `)` type = ')' ) )
                                        act = lexe( |(a{ repeat( val = ',a'
                                                                 occ = 5000 ) })| ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `Table1` type = 'W' )
                                                              ( value = `[`     type = `[` )
                                                              ( value = `]`     type = `]` )  )
                                        act = lexe( 'Table1[]' ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `Table1`    type = 'W' )
                                                              ( value = `[Column1]` type = `[` ) )
                                        act = lexe( 'Table1[Column1]' ) ).
  ENDMETHOD.

  METHOD test2.
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) )
                                        act = lexe( `DeptSales[[#Headers],[#Data],[% Commission]]` ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) )
                                        act = lexe( `DeptSales[ [#Headers],[#Data],[% Commission] ]` ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) )
                                        act = lexe( `DeptSales[[#Headers], [#Data], [% Commission]]` ) ).
    cl_abap_unit_assert=>assert_equals( exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) )
                                        act = lexe( `DeptSales[ [#Headers], [#Data], [% Commission] ]` ) ).
  ENDMETHOD.

  METHOD test3.
    cl_abap_unit_assert=>assert_equals( exp = 2
                                        act = evaluate( parse( lexe( `1+1` ) ) ) ).
  ENDMETHOD.

  METHOD test31.
    DATA(app) = lcl_excelom=>create( ).
    DATA(workbook) = app->workbooks( )->add( 'name' ).
    DATA(worksheet) = workbook->worksheets( )->item( 'Sheet1' ).
    worksheet->range( 'A1' )->value( )->set_double( 10 ).
    DATA(range) = worksheet->range( 'A2' ).
    range->formula2( )->set( '=A1+1' ).
    app->calculate( ).
    cl_abap_unit_assert=>assert_equals( exp = 11
                                        act = range->value( ) ).
  ENDMETHOD.

  METHOD test4.
*    data(a) = parse( lexe(
*`IFERROR(IF(C2<>"",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assig` &&
*`ned Attorney, or Sales Team",B2<>"Jimmy Edwards",B2<>"Kathleen McCarthy"),B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assigned Attorney, or Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(VL` &&
*`OOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(C2<>"",VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1` &&
*`!$A:$B,2,FALSE),"INTAKE TEAM")))))), VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE),"")` ) ).
  ENDMETHOD.
ENDCLASS.
