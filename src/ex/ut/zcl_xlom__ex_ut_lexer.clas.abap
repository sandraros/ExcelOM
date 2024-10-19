class ZCL_XLOM__EX_UT_LEXER definition
  public
  final
  create private .

public section.

  types TY_TOKEN_TYPE type STRING .
  types:
    BEGIN OF ts_token,
        value TYPE string,
        type  TYPE ty_token_type,
      END OF ts_token .
  types:
    tt_token TYPE STANDARD TABLE OF ts_token WITH EMPTY KEY .
  types:
    BEGIN OF ts_parenthesis_group,
        from_token          TYPE i,
        to_token            TYPE i,
        level               TYPE i,
        last_subgroup_token TYPE i,
      END OF ts_parenthesis_group .
  types:
    tt_parenthesis_group TYPE STANDARD TABLE OF ts_parenthesis_group WITH EMPTY KEY .
  types:
    BEGIN OF ts_result_lexe,
        tokens TYPE tt_token,
      END OF ts_result_lexe .

  constants:
    BEGIN OF c_type,
        comma                      TYPE ty_token_type VALUE ',',
        comma_space                TYPE ty_token_type VALUE `, `,
        curly_bracket_close        TYPE ty_token_type VALUE '}',
        curly_bracket_open         TYPE ty_token_type VALUE '{',
        "! In RIGHT("hello",) the second argument is empty, interpreted as 0,
        "! which is different from RIGHT("hello"), where the second argument
        "! is interpreted as being 1.
        empty_argument             TYPE ty_token_type VALUE '∅',
        "! #N/A!, etc.
        error_name                 TYPE ty_token_type VALUE '#',
        "! LEN(...), etc.
        function_name              TYPE ty_token_type VALUE 'F',
        number                     TYPE ty_token_type VALUE 'N',
        operator                   TYPE ty_token_type VALUE 'O',
        parenthesis_close          TYPE ty_token_type VALUE ')',
        parenthesis_open           TYPE ty_token_type VALUE '(',
        semicolon                  TYPE ty_token_type VALUE ';',
        square_bracket_close       TYPE ty_token_type VALUE ']',
        square_bracket_space_close TYPE ty_token_type VALUE ' ]',
        square_bracket_open        TYPE ty_token_type VALUE '[',
        square_brackets_open_close TYPE ty_token_type VALUE '[]',
        symbol_name                TYPE ty_token_type VALUE 'W',
        table_name                 TYPE ty_token_type VALUE 'T',
        text_literal               TYPE ty_token_type VALUE '"',
      END OF c_type .

  class-methods CREATE
    returning
      value(RESULT) type ref to ZCL_XLOM__EX_UT_LEXER .
  methods LEXE
    importing
      !TEXT type CSEQUENCE
    returning
      value(RESULT) type TT_TOKEN .
  PRIVATE SECTION.
    "! Insert the parts of the text in "FIND ... IN text ..." for which there was no match.
    METHODS complete_with_non_matches
      IMPORTING i_string  TYPE string
      CHANGING  c_matches TYPE match_result_tab.
ENDCLASS.



CLASS ZCL_XLOM__EX_UT_LEXER IMPLEMENTATION.


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
    result = NEW zcl_xlom__ex_ut_lexer( ).
  ENDMETHOD.


  METHOD lexe.
    TYPES ty_ref_to_parenthesis_group TYPE REF TO ts_parenthesis_group.

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
    FIND ALL OCCURRENCES OF REGEX '(?:'
         & '\('
         & '|\{'
         & '|\[ '             " opening bracket after table name
         & '|\['              " table column name, each character can be:
         & '(?:''.'       "   either one single quote (escape) with next character
         & '|[^\[\]]'       "   or any other character except [ and ]
         & ')+'
         & '\]'
         & '|\['              " opening bracket after table name
         & '|\)'
         & '|\}'
         & '| ?\]'
         & '|, ?'
         & '|;'
         & '|:'
         & '|<>'
         & '|<='
         & '|>='
         & '|<'
         & '|>'
         & '|='
         & '|\+'
         & '|-'
         & '|\*'
         & '|/'
         & '|\^'
         & '|&'
         & '|%'
         & '|"(?:""|[^"])*"'  " string literal
         & '|#[A-Z0-9/!?]+'      " error name (#DIV/0!, #N/A, #VALUE!, #GETTING_DATA!, #NAME?, etc.)
         & ')'
         IN text RESULTS DATA(matches).

    complete_with_non_matches( EXPORTING i_string  = text
                               CHANGING  c_matches = matches ).

    DATA(token_values) = VALUE string_table( ).
    LOOP AT matches REFERENCE INTO DATA(match).
      INSERT substring( val = text
                        off = match->offset
                        len = match->length )
             INTO TABLE token_values.
    ENDLOOP.

    " TODO: variable is assigned but never used (ABAP cleaner)
    DATA(current_parenthesis_group) = VALUE ty_ref_to_parenthesis_group( ).
    " TODO: variable is assigned but never used (ABAP cleaner)
    DATA(parenthesis_group) = VALUE ts_parenthesis_group( ).
    " TODO: variable is assigned but never used (ABAP cleaner)
    DATA(parenthesis_level) = 0.
    " TODO: variable is assigned but never used (ABAP cleaner)
    DATA(table_specification) = abap_false.
    DATA(token) = VALUE ts_token( ).
    DATA(token_number) = 1.
    LOOP AT token_values REFERENCE INTO DATA(token_value).
      " is comma a separator or a union operator?
      " https://techcommunity.microsoft.com/t5/excel/does-the-union-operator-exist/m-p/2590110
      " With argument-list functions, there is no union. Example: A1 contains 1, both =SUM(A1,A1) and =SUM((A1,A1)) return 2.
      " With no-argument-list functions, there is a union. Example: =LARGE((A1,B1),2) (=LARGE(A1,B1,2) is invalid, too many arguments)
      CASE token_value->*.
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
          token = VALUE #( value = condense( token_value->* )
                           type  = condense( token_value->* ) ).
        WHEN ` `
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
          token = VALUE #( value = token_value->*
                           type  = 'O' ).
        WHEN OTHERS.
          DATA(first_character) = substring( val = token_value->*
                                             len = 1 ).
          IF first_character = '"'.
            " text literal
            token = VALUE #( value = replace( val  = substring( val = token_value->*
                                                                off = 1
                                                                len = strlen( token_value->* ) - 2 )
                                              sub  = '""'
                                              with = '"'
                                              occ  = 0 )
                             type  = c_type-text_literal ).
          ELSEIF first_character = '['.
            " table argument
            token = VALUE #( value = token_value->*
                             type  = c_type-square_bracket_open ).
          ELSEIF first_character = '#'.
            " error name
            token = VALUE #( value = token_value->*
                             type  = c_type-error_name ).
          ELSEIF first_character CA '0123456789.-+'.
            " number
            token = VALUE #( value = token_value->*
                             type  = c_type-number ).
          ELSE.
            " function name, --, cell reference, table name, name of named range, constant (TRUE, FALSE)
            TYPES ty_ref_to_string TYPE REF TO string.
            DATA(next_token_value) = COND ty_ref_to_string( WHEN token_number < lines( token_values )
                                                            THEN REF #( token_values[ token_number + 1 ] ) ).
            DATA(token_type) = c_type-symbol_name.
            IF next_token_value IS BOUND.
              DATA(next_token_first_character) = substring( val = next_token_value->*
                                                            len = 1 ).
              CASE next_token_first_character.
                WHEN '('.
                  token_type = c_type-function_name.
                  DELETE token_values INDEX token_number + 1.
                WHEN '['.
                  token_type = c_type-table_name.
                  CASE next_token_value->*.
                    WHEN `[` OR `[ `.
                      DELETE token_values INDEX token_number + 1.
                    WHEN OTHERS.
                      IF strlen( next_token_value->* ) > 2.
                        " Excel formula "table[column]"; 1 token "[column]" becomes 2 tokens "[column]" and "]".
                        INSERT ']' INTO token_values INDEX token_number + 2.
                      ENDIF.
                  ENDCASE.
              ENDCASE.
            ENDIF.
            token = VALUE #( value = token_value->*
                             type  = token_type ).
          ENDIF.
      ENDCASE.

      INSERT token INTO TABLE result.
      token_number = token_number + 1.
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
