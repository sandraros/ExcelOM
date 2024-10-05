REPORT zzsro_excel_formula_engine.

" Excel saves formulas with R1C1-style addresses in A1 style.
"
" EXCELOM will convert R1C1-style addresses into A1 style during the parsing
" of the formula, which means that the parsing needs to know in which cell
" the formula is stored to resolve addresses like R[+1]C[+1].
" The method "calculate" doesn't need to know in which cell the formula is stored
" because the addresses are stored internally in A1 style but needs to know in which
" worksheet the formula runs ("B2" might refer to any worksheet). To simplify,
" the parsing will store internally the exact cell referred
"
" & : not a function but is an alias of CONCATENATE
" ADDRESS
" AND : logical boolean function
" CELL
" CHOOSE : conditional function, the first argument is an integer, if it's 1 the result will be the second argument, if it's 2 the result will be the third argument, etc.
" COLUMN
" CONCATENATE
" COUNTIF
" FILTER
" FIND
" FLOOR.MATH
" IF
" IFERROR
" IFS
" INDEX
" INDIRECT
" LEFT
" LEN
" MATCH
" MID
" MOD
" OFFSET
" RIGHT
" ROW
" T
" VLOOKUP
"
"  formula with operation on arrays        result
"
"  -> if the left operand is one line high, the line is replicated till max lines of the right operand.
"  -> if the left operand is one column wide, the column is replicated till max columns of the right operand.
"  -> if the right operand is one line high, the line is replicated till max lines of the left operand.
"  -> if the right operand is one column wide, the column is replicated till max columns of the left operand.
"
"  -> if the left operand has less lines than the right operand, additional lines are added with #N/A.
"  -> if the left operand has less columns than the right operand, additional columns are added with #N/A.
"  -> if the right operand has less lines than the left operand, additional lines are added with #N/A.
"  -> if the right operand has less columns than the left operand, additional columns are added with #N/A.
"
"  -> target array size = max lines of both operands + max columns of both operands.
"  -> each target cell of the target array is calculated like this:
"     T(1,1) = L(1,1) op R(1,1)
"     T(2,1) = L(2,1) op R(2,1)
"     etc.
"     If the left cell or right cell is #N/A, the target cell is also #N/A.
"
"  Examples where one of the two operands has 1 cell, 1 line or 1 column
"
"  a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
"
"  a | b | c   op   k                      a op k | b op k | c op k
"  d | e | f                               d op k | e op k | f op k
"  g | h | i                               g op k | h op k | i op k
"
"  a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
"  d | e | f                               d op k | e op l | f op m | #N/A
"  g | h | i                               g op k | h op l | i op m | #N/A
"
"  a | b | c   op   k                      a op k | b op k | c op k
"  d | e | f        l                      d op l | e op l | f op l
"  g | h | i        m                      g op m | h op m | i op m
"                   n                      #N/A   | #N/A   | #N/A
"
"  a | b | c   op   k                      a op k | b op k | c op k
"  d | e | f        l                      d op l | e op l | f op l
"  g | h | i                               #N/A   | #N/A   | #N/A
"
"  a | b | c   op   k                      a op k | b op k | c op k
"                   l                      a op l | b op l | c op l
"                   m                      a op m | b op m | c op m
"
"  Both operands have more than 1 line and more than 1 column
"
"  a | b | c   op   k | n                  a op k | b op n | #N/A
"  d | e | f        l | o                  d op l | e op o | #N/A
"  g | h | i                               #N/A   | #N/A   | #N/A
"
"  a | b | c   op   k | n                  a op k | b op n | #N/A
"  d | e | f        l | o                  d op l | e op o | #N/A
"                   m | p                  #N/A   | #N/A   | #N/A
"
"  a | b       op   k | n | q              a op k | b op n | #N/A
"  d | e            l | o | r              d op l | e op o | #N/A
"  g | h                                   #N/A   | #N/A   | #N/A

CLASS lcl_excelom_application DEFINITION DEFERRED.
CLASS lcl_excelom_error_value DEFINITION DEFERRED.
CLASS lcl_excelom_evaluation_context DEFINITION DEFERRED.
CLASS lcl_excelom_expr_array DEFINITION DEFERRED.
CLASS lcl_excelom_expr_expressions DEFINITION DEFERRED.
CLASS lcl_excelom_expr_function_call DEFINITION DEFERRED.
CLASS lcl_excelom_expr_number DEFINITION DEFERRED.
CLASS lcl_excelom_exprh_operator DEFINITION DEFERRED.
CLASS lcl_excelom_exprh_parser DEFINITION DEFERRED.
CLASS lcl_excelom_expr_plus DEFINITION DEFERRED.
CLASS lcl_excelom_expr_string DEFINITION DEFERRED.
CLASS lcl_excelom_expr_sub_expr DEFINITION DEFERRED.
CLASS lcl_excelom_expr_table DEFINITION DEFERRED.
*CLASS lcl_excelom_formula2 DEFINITION DEFERRED.
CLASS lcl_excelom_range DEFINITION DEFERRED.
CLASS lcl_excelom_range_value DEFINITION DEFERRED.
CLASS lcl_excelom_result_array DEFINITION DEFERRED.
CLASS lcl_excelom_result_error DEFINITION DEFERRED.
CLASS lcl_excelom_result_number DEFINITION DEFERRED.
CLASS lcl_excelom_result_string DEFINITION DEFERRED.
CLASS lcl_excelom_workbook DEFINITION DEFERRED.
CLASS lcl_excelom_workbooks DEFINITION DEFERRED.
CLASS lcl_excelom_worksheet DEFINITION DEFERRED.
CLASS lcl_excelom_worksheets DEFINITION DEFERRED.
CLASS lcx_excelom_expr_parser DEFINITION DEFERRED.
CLASS lcx_excelom_to_do DEFINITION DEFERRED.
CLASS lcx_excelom_unexpected DEFINITION DEFERRED.
INTERFACE lif_excelom_all_friends DEFERRED.
INTERFACE lif_excelom_expr DEFERRED.
INTERFACE lif_excelom_result DEFERRED.


CLASS lcx_excelom_to_do DEFINITION INHERITING FROM cx_no_check.
ENDCLASS.


CLASS lcx_excelom_unexpected DEFINITION INHERITING FROM cx_no_check.
ENDCLASS.


CLASS lcx_excelom_expr_parser DEFINITION INHERITING FROM cx_static_check.
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


CLASS lcx_excelom_expr_parser IMPLEMENTATION.
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


INTERFACE lif_excelom_all_friends.
ENDINTERFACE.


INTERFACE lif_excelom_expr.
  TYPES ty_expression_type TYPE i.

  CONSTANTS:
    BEGIN OF c_type,
      array          TYPE ty_expression_type VALUE 1,
      number         TYPE ty_expression_type VALUE 2,
      operation_mult TYPE ty_expression_type VALUE 3,
      operation_plus TYPE ty_expression_type VALUE 4,
      text_literal   TYPE ty_expression_type VALUE 5,
    END OF c_type.

  DATA type TYPE ty_expression_type READ-ONLY.

  METHODS is_equal
  IMPORTING expression TYPE REF TO lif_excelom_expr
    RETURNING VALUE(result) type abap_bool.

  METHODS evaluate
  importing context type ref to lcl_excelom_evaluation_context
  RETURNING VALUE(result) TYPE REF TO lif_excelom_result.
ENDINTERFACE.


INTERFACE lif_excelom_result.
  TYPES ty_type TYPE i.

  CONSTANTS:
    BEGIN OF c_type,
      number TYPE ty_type VALUE 1,
      string TYPE ty_type VALUE 2,
      array  TYPE ty_type VALUE 3,
      error  TYPE ty_type VALUE 4,
    END OF c_type.

  DATA type         TYPE ty_type READ-ONLY.
  DATA row_count    TYPE i       READ-ONLY.
  DATA column_count TYPE i       READ-ONLY.

  METHODS get_cell_value
    IMPORTING column_offset TYPE i
              row_offset    TYPE i
    RETURNING VALUE(result) TYPE REF TO lif_excelom_result.

  METHODS is_array
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_boolean
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_error
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_number
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_string
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS set_cell_value
    IMPORTING column_offset TYPE i
              row_offset    TYPE i
              !value        TYPE REF TO lif_excelom_result.
ENDINTERFACE.


CLASS lcl_excelom_error_value DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS get_singleton
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_error_value.

  PRIVATE SECTION.
    CLASS-DATA singleton TYPE REF TO lcl_excelom_error_value.
ENDCLASS.


CLASS lcl_excelom_exprh DEFINITION.
  PUBLIC SECTION.
    TYPES:
      BEGIN OF ts_evaluate_array_operands,
        result        TYPE REF TO lif_excelom_result,
        left_operand  TYPE REF TO lif_excelom_result,
        right_operand TYPE REF TO lif_excelom_result,
      END OF ts_evaluate_array_operands.
    CLASS-METHODS evaluate_array_operands
      IMPORTING expression    TYPE REF TO lif_excelom_expr
                context type REF TO lcl_excelom_evaluation_context
                left_operand  TYPE REF TO lif_excelom_expr
                right_operand TYPE REF TO lif_excelom_expr
      RETURNING VALUE(result) TYPE ts_evaluate_array_operands.
ENDCLASS.


CLASS lcl_excelom_exprh_lexer DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    types TY_token_type type string.

    TYPES:
      BEGIN OF ts_token,
        value TYPE string,
        type  TYPE TY_token_type,
      END OF ts_token.
    TYPES tt_token TYPE STANDARD TABLE OF ts_token WITH EMPTY KEY.

    TYPES:
      BEGIN OF ts_parenthesis_group,
        from_token TYPE i,
        to_token   TYPE i,
        level      TYPE i,
        last_subgroup_token type i,
      END OF ts_parenthesis_group.
    TYPES tt_parenthesis_group TYPE STANDARD TABLE OF ts_parenthesis_group WITH EMPTY KEY.

    TYPES:
      BEGIN OF ts_result_lexe,
        tokens             TYPE tt_token,
*        parenthesis_groups TYPE tt_parenthesis_group,
      END OF ts_result_lexe.

    CONSTANTS:
      BEGIN OF c_type,
        comma                      TYPE ty_token_type VALUE ',',
        comma_space                TYPE ty_token_type VALUE `, `,
        curly_bracket_close        TYPE ty_token_type VALUE '}',
        curly_bracket_open         TYPE ty_token_type VALUE '{',
        function_name              TYPE ty_token_type VALUE 'F',
        number                     TYPE ty_token_type VALUE 'N',
        operator                   TYPE ty_token_type VALUE 'O',
        parenthesis_close          TYPE ty_token_type VALUE ')',
        parenthesis_open           TYPE ty_token_type VALUE '(',
        semicolon                  TYPE ty_token_type VALUE ',',
        square_bracket_close       TYPE ty_token_type VALUE ']',
        square_bracket_space_close TYPE ty_token_type VALUE ' ]',
        square_bracket_open        TYPE ty_token_type VALUE '[',
        square_brackets_open_close TYPE ty_token_type VALUE '[]',
        symbol_name                TYPE ty_token_type VALUE 'W',
        table_name                 TYPE ty_token_type VALUE 'T',
        text_literal               TYPE ty_token_type VALUE '"',
      END OF c_type.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_exprh_lexer.

    METHODS lexe IMPORTING !text         TYPE csequence
                 RETURNING VALUE(result) TYPE tt_token. "ts_result_lexe.

  PRIVATE SECTION.
    "! Insert the parts of the text in "FIND ... IN text ..." for which there was no match.
    METHODS complete_with_non_matches
      IMPORTING i_string  TYPE string
      CHANGING  c_matches TYPE match_result_tab.
ENDCLASS.


CLASS lcl_excelom_exprh_group DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    TYPES:
      BEGIN OF ts_item,
        token      TYPE REF TO lcl_excelom_exprh_lexer=>ts_token,
        group      TYPE REF TO lcl_excelom_exprh_group,
        operator   TYPE REF TO lcl_excelom_exprh_operator,
        priority   TYPE i,
        expression TYPE REF TO lif_excelom_expr,
      END OF ts_item.
    TYPES tt_item TYPE STANDARD TABLE OF ts_item.

    DATA type       TYPE lcl_excelom_exprh_lexer=>ts_token-type READ-ONLY.
    DATA operator   TYPE REF TO lcl_excelom_exprh_operator      READ-ONLY.
    DATA expression TYPE REF TO lif_excelom_expr     READ-ONLY.
    DATA items      TYPE tt_item                                READ-ONLY.

    METHODS append
      IMPORTING item TYPE ts_item.

    CLASS-METHODS create
        IMPORTING type TYPE lcl_excelom_exprh_lexer=>ts_token-type
        RETURNING VALUE(result) TYPE REF TO lcl_excelom_exprh_group.

    METHODS delete
      IMPORTING !index TYPE i.

    METHODS insert
      IMPORTING item   TYPE ts_item
                !index TYPE i.

    METHODS set_expression
      IMPORTING expression TYPE REF TO lif_excelom_expr.

    METHODS set_item_expression
      IMPORTING !index     TYPE sytabix
                expression TYPE REF TO lif_excelom_expr.

    METHODS set_item_group
      IMPORTING !index TYPE sytabix
                !group TYPE REF TO lcl_excelom_exprh_group.

    METHODS set_item_operator
      IMPORTING !index   TYPE sytabix
                operator TYPE REF TO lcl_excelom_exprh_operator.

    METHODS set_operator
      IMPORTING operator TYPE REF TO lcl_excelom_exprh_operator.
    METHODS set_item_priority
      IMPORTING
        index    TYPE syst-tabix
        priority TYPE i.

  PRIVATE SECTION.
ENDCLASS.


CLASS lcl_excelom_exprh_operator DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
*    INTERFACES lif_excelom_expr_expression.
*    INTERFACES lif_excelom_expr_operator.

    TYPES tt_operand_position_offset TYPE STANDARD TABLE OF i WITH EMPTY KEY.
    TYPES tt_expression              TYPE STANDARD TABLE OF REF TO lif_excelom_expr WITH EMPTY KEY.

    CLASS-DATA multiply TYPE REF TO lcl_excelom_exprh_operator READ-ONLY.
    CLASS-DATA plus TYPE REF TO lcl_excelom_exprh_operator READ-ONLY.

    CLASS-METHODS class_constructor.

    CLASS-METHODS create
      IMPORTING !name                    TYPE string
                operand_position_offsets TYPE tt_operand_position_offset
                !priority                TYPE i
      RETURNING VALUE(result)            TYPE REF TO lcl_excelom_exprh_operator.

    METHODS create_expression
      IMPORTING operands      TYPE tt_expression
      RETURNING VALUE(result) TYPE REF TO lif_excelom_expr.

    CLASS-METHODS get
      IMPORTING operator TYPE string
                unary    TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_exprh_operator.

    "! <ul>
    "! <li>1 : Reference operators ":" (colon), " " (single space), "," (comma)</li>
    "! <li>2 : – (as in –1) and + (as in +1)</li>
    "! <li>3 : % (as in =50%)</li>
    "! <li>4 : ^ Exponentiation (as in 2^8)</li>
    "! <li>5 : * and / Multiplication and division                    </li>
    "! <li>6 : + and – Addition and subtraction                       </li>
    "! <li>7 : & Connects two strings of text (concatenation)         </li>
    "! <li>8 : = < > <= >= <> Comparison</li>
    "! </ul>
    "!
    "! @parameter result | .
    METHODS get_priority
      RETURNING VALUE(result) TYPE i.

    "! 1 : predecessor operand only (% e.g. 10%)
    "! 2 : before and after operand only (+ - * / ^ & e.g. 1+1)
    "! 3 : successor operand only (unary + and - e.g. +5)
    "!
    "! @parameter result | .
    METHODS get_operand_position_offsets
      RETURNING VALUE(result) TYPE tt_operand_position_offset.

*  METHODS set_operands IMPORTING predecessor TYPE REF TO lif_excelom_expr_expression OPTIONAL
*                                 successor   TYPE REF TO lif_excelom_expr_expression OPTIONAL.

  PRIVATE SECTION.

    TYPES:
      "! operator precedence
      "! Get operator priorities
      BEGIN OF ts_operator,
        name              TYPE string,
        "! +1 for unary operators (e.g. -1)
        "! -1 and +1 for binary operators (e.g. 1*2)
        "! -1 for postfix operators (e.g. 10%)
        operand_position_offsets TYPE tt_operand_position_offset,
        "! To distinguish unary from binary operators + and -
        unary             TYPE abap_bool,
        "! <ul>
        "! <li>1 : Reference operators ":" (colon), " " (single space), "," (comma)</li>
        "! <li>2 : – (as in –1) and + (as in +1)</li>
        "! <li>3 : % (as in =50%)</li>
        "! <li>4 : ^ Exponentiation (as in 2^8)</li>
        "! <li>5 : * and / Multiplication and division                    </li>
        "! <li>6 : + and – Addition and subtraction                       </li>
        "! <li>7 : & Connects two strings of text (concatenation)         </li>
        "! <li>8 : = < > <= >= <> Comparison</li>
        "! </ul>
        priority          TYPE i,
*        "! % is the only postfix operator e.g. 10% (=0.1)
*        postfix           TYPE abap_bool,
        desc              TYPE string,
        handler           TYPE REF TO lcl_excelom_exprh_operator,
      END OF ts_operator.
    TYPES tt_operator TYPE SORTED TABLE OF ts_operator WITH UNIQUE KEY name unary.

    CLASS-DATA operators TYPE lcl_excelom_exprh_operator=>tt_operator.

    DATA name                     TYPE string.
    "! +1 for unary operators (e.g. -1)
    "! -1 and +1 for binary operators (e.g. 1*2)
    "! -1 for postfix operators (e.g. 10%)
    DATA operand_position_offsets TYPE tt_operand_position_offset.
    "! <ul>
    "! <li>1 : Reference operators ":" (colon), " " (single space), "," (comma)</li>
    "! <li>2 : – (as in –1) and + (as in +1)</li>
    "! <li>3 : % (as in =50%)</li>
    "! <li>4 : ^ Exponentiation (as in 2^8)</li>
    "! <li>5 : * and / Multiplication and division                    </li>
    "! <li>6 : + and – Addition and subtraction                       </li>
    "! <li>7 : & Connects two strings of text (concatenation)         </li>
    "! <li>8 : = < > <= >= <> Comparison</li>
    "! </ul>
    DATA priority                 TYPE i.
ENDCLASS.


CLASS lcl_excelom_exprh_parser DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS create
*      IMPORTING formula_cell  TYPE REF TO lcl_excelom_range
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_exprh_parser.

    METHODS parse
      IMPORTING !tokens            TYPE lcl_excelom_exprh_lexer=>tt_token
*                parenthesis_groups TYPE lcl_excelom_exprh_lexer=>tt_parenthesis_group
      RETURNING VALUE(result)      TYPE REF TO lif_excelom_expr
      RAISING   lcx_excelom_expr_parser.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ts_parsed_group,
        from_token TYPE i,
        to_token   TYPE i,
        expression TYPE REF TO lif_excelom_expr,
      END OF ts_parsed_group.
    TYPES tt_parsed_group TYPE STANDARD TABLE OF ts_parsed_group WITH EMPTY KEY.

*    DATA formula_cell        TYPE REF TO lcl_excelom_range.
    DATA formula_offset      TYPE i.
    DATA current_token_index TYPE sytabix.
    DATA tokens              TYPE lcl_excelom_exprh_lexer=>tt_token.
*    DATA parenthesis_groups  TYPE lcl_excelom_exprh_lexer=>tt_parenthesis_group.
    DATA parsed_groups       TYPE tt_parsed_group.
    DATA previous_token_type TYPE lcl_excelom_exprh_lexer=>ts_token-type.
    DATA: current_token TYPE REF TO lcl_excelom_exprh_lexer=>ts_token.



    METHODS parse_expression
      RETURNING VALUE(result) TYPE REF TO lif_excelom_expr
      RAISING   lcx_excelom_expr_parser.

    METHODS parse_function_arguments
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_expressions
      RAISING   lcx_excelom_expr_parser.

    METHODS parse_tokens_up_to
      IMPORTING stop_at_token TYPE csequence
      RETURNING VALUE(result) TYPE string_table.

    METHODS skip_spaces.





    METHODS parse_expression_group
      IMPORTING group type REF TO lcl_excelom_exprh_group.

    METHODS parse_expression_group_2
      IMPORTING group type REF TO lcl_excelom_exprh_group.

    METHODS parse_expression_group_3
      CHANGING !group TYPE REF TO lcl_excelom_exprh_group.

    METHODS parse_expression_group_4
      IMPORTING  group TYPE REF TO lcl_excelom_exprh_group.
    METHODS get_expression_from_symbol_nam
      IMPORTING
        token_value   TYPE string
      RETURNING
        value(result) TYPE REF TO lif_excelom_expr.
ENDCLASS.


CLASS lcl_excelom_evaluation_context DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.

    data containing_cell type ref to lcl_excelom_range READ-ONLY.

    CLASS-METHODS create
      IMPORTING containing_cell type ref to lcl_excelom_range
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_evaluation_context.

  PRIVATE SECTION.
ENDCLASS.


CLASS lcl_excelom_expr_array DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_array.
ENDCLASS.


CLASS lcl_excelom_expr_boolean DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING boolean_value TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_boolean.

  PRIVATE SECTION.
    DATA boolean_value TYPE abap_bool.
endclass.


CLASS lcl_excelom_expr_expressions DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    METHODS append IMPORTING expression TYPE REF TO lif_excelom_expr.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_expressions.
ENDCLASS.


CLASS lcl_excelom_expr_function_call DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING !name         TYPE csequence
                arguments     TYPE REF TO lcl_excelom_expr_expressions
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_function_call.
ENDCLASS.


CLASS lcl_excelom_expr_mult DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING left_operand  TYPE REF TO lif_excelom_expr
                right_operand TYPE REF TO lif_excelom_expr
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_mult.

  PRIVATE SECTION.
    DATA left_operand  TYPE REF TO lif_excelom_expr.
    DATA right_operand TYPE REF TO lif_excelom_expr.
ENDCLASS.


CLASS lcl_excelom_expr_number DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING !number       TYPE f
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_number.

  PRIVATE SECTION.
    DATA number TYPE f.
ENDCLASS.


*CLASS lcl_excelom_expr_operation DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_excelom_expr_expression.
**    INTERFACES lif_excelom_expr_operator.
*
*    CLASS-METHODS create
*      IMPORTING operator      TYPE REF TO lcl_excelom_expr_operator
*                operands  TYPE REF TO lif_excelom_expr_expression
*                right_operand TYPE REF TO lif_excelom_expr_expression
*      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_plus.
*
*  PRIVATE SECTION.
*    DATA left_operand  TYPE REF TO lif_excelom_expr_expression.
*    DATA right_operand TYPE REF TO lif_excelom_expr_expression.
*ENDCLASS.


CLASS lcl_excelom_expr_plus DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING left_operand  TYPE REF TO lif_excelom_expr
                right_operand TYPE REF TO lif_excelom_expr
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_plus.

  PRIVATE SECTION.
    DATA left_operand  TYPE REF TO lif_excelom_expr.
    DATA right_operand TYPE REF TO lif_excelom_expr.
ENDCLASS.


CLASS lcl_excelom_expr_range DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_all_friends.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING address_or_name TYPE string
      RETURNING VALUE(result)   TYPE REF TO lcl_excelom_expr_range.

  PRIVATE SECTION.
    DATA _address_or_name TYPE string.
*    DATA range TYPE REF TO lcl_excelom_range.
ENDCLASS.


CLASS lcl_excelom_expr_string DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_string.

  PRIVATE SECTION.
    DATA text TYPE string.
ENDCLASS.


CLASS lcl_excelom_expr_sub_expr DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_sub_expr.
ENDCLASS.


CLASS lcl_excelom_expr_table DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr.

    CLASS-METHODS create
      IMPORTING table_name            TYPE csequence
                row_column_specifiers TYPE string_table
      RETURNING VALUE(result)         TYPE REF TO lcl_excelom_expr_table.
ENDCLASS.


CLASS lcl_excelom_cell_format DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_cell_format.

  PRIVATE SECTION.
    CLASS-DATA singleton TYPE REF TO lcl_excelom_error_value.
ENDCLASS.


*CLASS lcl_excelom_formula2 DEFINITION FINAL
*  CREATE PRIVATE FRIENDS lif_excelom_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_excelom_all_friends.
*
*    METHODS calculate.
*
*    CLASS-METHODS create
*      IMPORTING !range        TYPE REF TO lcl_excelom_range
*      RETURNING VALUE(result) TYPE REF TO lcl_excelom_formula2.
*
*    METHODS set_value
*      IMPORTING !value TYPE string
*      RAISING   lcx_excelom_expr_parser.
*
*  PRIVATE SECTION.
*    DATA range       TYPE REF TO lcl_excelom_range.
*    DATA _expression TYPE REF TO lif_excelom_expr.
*ENDCLASS.


CLASS lcl_excelom_range DEFINITION FINAL
  CREATE PRIVATE
  FRIENDS lif_excelom_all_friends.

  PUBLIC SECTION.
    INTERFACES lif_excelom_all_friends.

*    DATA value TYPE REF TO lif_excelom_result READ-ONLY.
    DATA formula2 TYPE string READ-ONLY.
    DATA parent TYPE REF TO lcl_excelom_worksheet READ-ONLY.
    DATA application TYPE REF TO lcl_excelom_application READ-ONLY.

    METHODS calculate.

    "! Called by the Worksheet.Range property.
    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    CLASS-METHODS create
      IMPORTING worksheet     TYPE REF TO lcl_excelom_worksheet
                cell1         TYPE REF TO lcl_excelom_range
                cell2         TYPE REF TO lcl_excelom_range OPTIONAL
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.

    CLASS-METHODS create_from_address_or_name
      IMPORTING address       TYPE clike
                relative_to   TYPE REF TO lcl_excelom_worksheet
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.

    METHODS set_formula2
      IMPORTING value TYPE string
      RAISING
        lcx_excelom_expr_parser.

    METHODS set_value
      IMPORTING value TYPE REF TO lif_excelom_result.

    METHODS value
      RETURNING VALUE(result) TYPE REF TO lif_excelom_result.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ty_address_one_cell,
        column       TYPE i,
        column_fixed TYPE abap_bool,
        row          TYPE i,
        row_fixed    TYPE abap_bool,
      END OF ty_address_one_cell.
    TYPES:
      BEGIN OF ty_address,
        top_left     TYPE ty_address_one_cell,
        bottom_right TYPE ty_address_one_cell,
      END OF ty_address.

    CLASS-METHODS convert_column_a_xfd_to_number
      IMPORTING roman_letters TYPE string
      RETURNING VALUE(result) TYPE i.

    METHODS decode_range_address
      IMPORTING address       TYPE string
      RETURNING VALUE(result) TYPE ty_address.
    METHODS _set_value
      IMPORTING
        value TYPE REF TO lif_excelom_result.

    CLASS-METHODS decode_range_address_a1
      IMPORTING address       TYPE string
      RETURNING VALUE(result) TYPE ty_address.

    CLASS-METHODS decode_range_address_r1_c1
      IMPORTING address       TYPE string
      RETURNING VALUE(result) TYPE ty_address.

    DATA _formula_expression TYPE REF TO lif_excelom_expr.
    DATA _address            TYPE ty_address.
ENDCLASS.


CLASS lcl_excelom_range_value DEFINITION FINAL
  CREATE PRIVATE
  FRIENDS lcl_excelom_worksheet.

  PUBLIC SECTION.
    INTERFACES lif_excelom_all_friends.

    CLASS-METHODS create
      IMPORTING !range        TYPE REF TO lcl_excelom_range
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range_value.

    METHODS set_double IMPORTING !value TYPE f.

    METHODS set_string IMPORTING !value TYPE string.

  PRIVATE SECTION.
    TYPES ty_value_type TYPE i.

    CONSTANTS:
      BEGIN OF c_value_type,
        empty         TYPE ty_value_type VALUE 1,
        number        TYPE ty_value_type VALUE 2,
        text          TYPE ty_value_type VALUE 3,
        "! Cell containing the value TRUE or FALSE.
        boolean       TYPE ty_value_type VALUE 4,
        error         TYPE ty_value_type VALUE 5,
        compound_data TYPE ty_value_type VALUE 6,
      END OF c_value_type.

    CONSTANTS:
      BEGIN OF c_boolean,
        false TYPE f VALUE 0,
        true  TYPE f VALUE -1,
      END OF c_boolean.

    "! Range to which the value applies
    DATA range TYPE REF TO lcl_excelom_range.

    METHODS set IMPORTING !value TYPE any
                          !type  TYPE ty_value_type.
ENDCLASS.


CLASS lcl_excelom_result_array DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    CLASS-METHODS create_from_range
      IMPORTING range       TYPE ref to lcl_excelom_range
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_result_array.

    CLASS-METHODS create_initial
      IMPORTING number_of_rows    TYPE i
                number_of_columns TYPE i
      RETURNING VALUE(result)     TYPE REF TO lcl_excelom_result_array.

  PRIVATE SECTION.
    DATA number_of_rows    TYPE i.
    DATA number_of_columns TYPE i.
ENDCLASS.


CLASS lcl_excelom_result_boolean DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    CLASS-METHODS create
      IMPORTING boolean_value TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_result_boolean.

  PRIVATE SECTION.
    DATA boolean_value TYPE abap_bool.
ENDCLASS.


CLASS lcl_excelom_result_empty DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    CLASS-METHODS get_singleton
        RETURNING VALUE(result) type ref to lcl_excelom_result_empty.
  PRIVATE SECTION.
    CLASS-DATA singleton type ref to lcl_excelom_result_empty.
ENDCLASS.


"! https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/cell-error-values
"! You can insert a cell error value into a cell or test the value of a cell for an error value by
"! using the CVErr function. The cell error values can be one of the following xlCVError constants.
"! NB: many errors are missing, the list of the other errors can be found in xlCVError enumeration.
"! <ul>
"! <li>Constant . .Error number . .Cell error value</li>
"! <li>xlErrDiv0 . 2007 . . . . . .#DIV/0!         </li>
"! <li>xlErrNA . . 2042 . . . . . .#N/A            </li>
"! <li>xlErrName . 2029 . . . . . .#NAME?          </li>
"! <li>xlErrNull . 2000 . . . . . .#NULL!          </li>
"! <li>xlErrNum . .2036 . . . . . .#NUM!           </li>
"! <li>xlErrRef . .2023 . . . . . .#REF!           </li>
"! <li>xlErrValue .2015 . . . . . .#VALUE!         </li>
"! </ul>
"! VB example:
"! <ul>
"! <li>If IsError(ActiveCell.Value) Then            </li>
"! <li>. If ActiveCell.Value = CVErr(xlErrDiv0) Then</li>
"! <li>. End If                                     </li>
"! <li>End If                                       </li>
"! </ul>
"! NB:
"! <ul>
"! <li>CVErr(xlErrDiv0) is of type Variant/Error and Locals/Watches shows: Error 2007</li>
"! <li>There is no Error data type, only Variant can be used.                        </li>
"! </ul>
CLASS lcl_excelom_result_error DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    TYPES ty_error_number TYPE i.

    CLASS-DATA blocked                    TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA calc                       TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA connect                    TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA division_by_zero           TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA field                      TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA getting_data               TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA na_not_applicable          TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA name                       TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA null                       TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA num                        TYPE REF TO lcl_excelom_result_error READ-ONLY.
    "! TODO #PYTHON! internal error number is not 2222, what is it?
    CLASS-DATA python                     TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA ref                        TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA spill                      TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA unknown                    TYPE REF TO lcl_excelom_result_error READ-ONLY.
    CLASS-DATA value_cannot_be_calculated TYPE REF TO lcl_excelom_result_error READ-ONLY.

    CLASS-METHODS class_constructor.

    CLASS-METHODS get_by_error_number
      IMPORTING !type         TYPE ty_error_number
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_result_error.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ts_error,
        error_name            TYPE string,
        internal_error_number TYPE ty_error_number,
        formula_error_number  TYPE ty_error_number,
        handler               TYPE REF TO lcl_excelom_result_error,
      END OF ts_error.
    TYPES tt_error TYPE STANDARD TABLE OF ts_error WITH EMPTY KEY.

    CLASS-DATA errors TYPE tt_error.

    DATA error_name TYPE string.
    DATA description TYPE string.
    DATA internal_error_number TYPE ty_error_number.
    DATA formula_error_number  TYPE ty_error_number.

    CLASS-METHODS create
      IMPORTING error_name type string
                internal_error_number TYPE ty_error_number
                formula_error_number TYPE ty_error_number
                description type string optional
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_result_error.
ENDCLASS.


CLASS lcl_excelom_result_number DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    CLASS-METHODS create
      IMPORTING !number       TYPE f
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_result_number.

    METHODS get_number
      RETURNING VALUE(result) TYPE f.

  PRIVATE SECTION.
    DATA number TYPE f.
ENDCLASS.


CLASS lcl_excelom_result_string DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    CLASS-METHODS create
      IMPORTING !string       TYPE csequence
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_result_string.

  PRIVATE SECTION.
    DATA string TYPE string.
ENDCLASS.


CLASS lcl_excelom_worksheet DEFINITION FINAL
  CREATE PRIVATE
  FRIENDS lif_excelom_all_friends.

  PUBLIC SECTION.
    TYPES ty_name TYPE string.

    DATA application TYPE REF TO lcl_excelom_application READ-ONLY.
    DATA parent TYPE REF TO lcl_excelom_workbook READ-ONLY.

    "! Worksheet.Calculate method (Excel).
    "! Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.
    "! <p>expression.Calculate</p>
    "! expression A variable that represents a Worksheet object.
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.calculate(method)
    METHODS calculate.

    CLASS-METHODS create
     IMPORTING workbook TYPE REF TO lcl_excelom_workbook
    RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

    "! Worksheet.Range property. Returns a Range object that represents a cell or a range of cells.
    "! <p>expression.Range (Cell1, Cell2)</p>
    "! expression A variable that represents a Worksheet object.
    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    METHODS range_from_address
      IMPORTING cell1         TYPE string
                cell2         TYPE string OPTIONAL
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.

    "! Worksheet.Range property. Returns a Range object that represents a cell or a range of cells.
    "! <p>expression.Range (Cell1, Cell2)</p>
    "! expression A variable that represents a Worksheet object.
    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    METHODS range_from_two_ranges
      IMPORTING cell1         TYPE REF TO lcl_excelom_range
                cell2         TYPE REF TO lcl_excelom_range
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.

  PRIVATE SECTION.
    TYPES ty_cell_type TYPE i.
    TYPES ty_value_type TYPE i.

    CONSTANTS:
      "! Formula function TYPE. TYPE(0) gives 1. An empty cell is of type NUMBER but it's impossible
      "! to differentiate a zero number from an empty cell with TYPE (the formula function ISBLANK may be used for that).
      "! https://support.microsoft.com/en-us/office/type-function-45b4e688-4bc3-48b3-a105-ffa892995899
      BEGIN OF c_excel_type,
        number        TYPE ty_cell_type VALUE 1,
        text          TYPE ty_cell_type VALUE 2,
        "! Cell containing TRUE or FALSE, or value calculated by formula
        logical_value TYPE ty_cell_type VALUE 4,
        error_value   TYPE ty_cell_type VALUE 16,
        array         TYPE ty_cell_type VALUE 64,
        compound_data TYPE ty_cell_type VALUE 128,
      END OF c_excel_type.

    CONSTANTS:
      BEGIN OF c_value_type,
        empty         TYPE ty_value_type VALUE 1,
        number        TYPE ty_value_type VALUE 2,
        text          TYPE ty_value_type VALUE 3,
        "! Cell containing the value TRUE or FALSE.
        boolean       TYPE ty_value_type VALUE 4,
        error         TYPE ty_value_type VALUE 5,
        compound_data TYPE ty_value_type VALUE 6,
      END OF c_value_type.

    CONSTANTS:
      BEGIN OF c_boolean,
        false TYPE f VALUE 0,
        true  TYPE f VALUE -1,
      END OF c_boolean.

    TYPES:
      BEGIN OF ts_cell,
        column        TYPE i,
        row           TYPE i,
        "! Type of cell value, among empty, number, text, boolean, error, compound data. For NUMBER, BOOLEAN and ERROR, the value is defined by VALUE2-DOUBLE.
        "! For TEXT, the value is defined by VALUE2-STRING.
        value_type          TYPE lcl_excelom_range_value=>ty_value_type,
        "! In arrays, it's empty in all cells except the top left cell where the array formula resides.
        formula2      TYPE string,
        "! In all cells of an array, it contains the array formula.
        formula_array TYPE string,
        "! False if formula2 is empty (in arrays, it's False in all cells except the top left cell where the array formula resides).
        has_formula   TYPE abap_bool,
        "! <p>Value of the cell. If formula2 or formula_array is defined, it contains the value calculated by the formula, otherwise
        "! it contains the value entered manually.</p>
        "! <p>Can return the values with the type Variant/Empty, Variant/Error.</p>
        "! <p>"A Variant can also contain the special values Empty, Error, Nothing, and Null."
        "! (source: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/variant-data-type)</p>
        BEGIN OF value2,
          "! Number, Error, Boolean: <ul>
          "! <li>If TYPE = C_TYPE-BOOLEAN, the possible values are the constants C_BOOLEAN-TRUE (-1) and C_BOOLEAN-FALSE (0).</li>
          "! <li>If TYPE = C_TYPE-ERROR, the possible values are the constants C_ERROR-NA_NOT_APPLICABLE, etc.</li>
          "! </ul>
          double  TYPE f,
          string  TYPE string,
        END OF value2,
        format TYPE REF TO lcl_excelom_cell_format,
      END OF ts_cell.
    TYPES tt_cell TYPE HASHED TABLE OF ts_cell WITH UNIQUE KEY column row.
*    TYPES tt_formula TYPE STANDARD TABLE OF REF TO lif_excelom_formula2 WITH EMPTY KEY.

*    DATA formulas TYPE tt_formula.
    DATA _cells   TYPE tt_cell.
ENDCLASS.


CLASS lcl_excelom_worksheets DEFINITION FRIENDS lif_excelom_all_friends.
  PUBLIC SECTION.
    INTERFACES lif_excelom_all_friends.

    DATA application TYPE REF TO lcl_excelom_application READ-ONLY.
    DATA count       TYPE i                              READ-ONLY.
    DATA workbook    TYPE REF TO lcl_excelom_workbook    READ-ONLY.

    CLASS-METHODS create
      IMPORTING workbook      TYPE REF TO lcl_excelom_workbook
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheets.

    METHODS add
      IMPORTING !name         TYPE lcl_excelom_worksheet=>ty_name
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

    "!
    "! @parameter index  | Required    Variant The name or index number of the object.
    "! @parameter result | .
    METHODS item
      IMPORTING !index        TYPE simple
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ty_worksheet,
        name   TYPE lcl_excelom_worksheet=>ty_name,
        object TYPE REF TO lcl_excelom_worksheet,
      END OF ty_worksheet.
    TYPES ty_worksheets TYPE SORTED TABLE OF ty_worksheet WITH UNIQUE KEY name.

    DATA worksheets TYPE ty_worksheets.
ENDCLASS.


CLASS lcl_excelom_workbook DEFINITION.
  PUBLIC SECTION.
    TYPES ty_name TYPE string.

    DATA application TYPE REF TO lcl_excelom_application READ-ONLY.
    DATA worksheets TYPE REF TO lcl_excelom_worksheets READ-ONLY.

    CLASS-METHODS create
      IMPORTING !application  TYPE REF TO lcl_excelom_application
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbook.

  PRIVATE SECTION.
ENDCLASS.


CLASS lcl_excelom_application DEFINITION FRIENDS lif_excelom_all_friends.
  PUBLIC SECTION.
    INTERFACES lif_excelom_all_friends.

    TYPES ty_calculation TYPE i.
    TYPES ty_reference_style TYPE i.

    CONSTANTS:
      BEGIN OF c_calculation,
        automatic     TYPE ty_calculation VALUE -4105,
        manual        TYPE ty_calculation VALUE -4135,
        semiautomatic TYPE ty_calculation VALUE 2,
      END OF c_calculation.

    CONSTANTS:
      BEGIN OF c_reference_style,
        a1    TYPE ty_reference_style VALUE 1,
        r1_c1 TYPE ty_reference_style VALUE -4150,
      END OF c_reference_style.

    DATA calculation     TYPE ty_calculation               VALUE c_calculation-automatic READ-ONLY.
    DATA reference_style TYPE ty_reference_style           VALUE c_reference_style-a1    READ-ONLY.
    DATA workbooks       TYPE REF TO lcl_excelom_workbooks                               READ-ONLY.

    METHODS calculate.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_application.

  PRIVATE SECTION.

    CLASS-METHODS type
      IMPORTING any_data_object TYPE any
      RETURNING VALUE(result)   TYPE abap_typekind.
ENDCLASS.


CLASS lcl_excelom_workbooks DEFINITION FRIENDS lif_excelom_all_friends.
  PUBLIC SECTION.
    INTERFACES lif_excelom_all_friends.

    DATA application TYPE REF TO lcl_excelom_application READ-ONLY.
    DATA count       TYPE i                              READ-ONLY.

    CLASS-METHODS create
      IMPORTING application TYPE REF TO lcl_excelom_application
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbooks.

    METHODS add
      IMPORTING !name         TYPE lcl_excelom_workbook=>ty_name
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbook.

    "!
    "! @parameter index  | Required    Variant The name or index number of the object.
    "! @parameter result | .
    METHODS item
      IMPORTING !index        TYPE simple
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





CLASS lcl_excelom_cell_format IMPLEMENTATION.
  METHOD create.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_error_value IMPLEMENTATION.
  METHOD get_singleton.
    IF singleton IS NOT BOUND.
      singleton = NEW lcl_excelom_error_value( ).
    ENDIF.
    result = singleton.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_exprh IMPLEMENTATION.
  METHOD evaluate_array_operands.
    result-left_operand = left_operand->evaluate( context ).
    result-right_operand = right_operand->evaluate( context ).

    CHECK result-left_operand->type = lif_excelom_result=>c_type-array
        OR result-right_operand->type = lif_excelom_result=>c_type-array.

    DATA(max_row_count) = nmax( val1 = result-left_operand->row_count
                                val2 = result-right_operand->row_count ).
    DATA(max_column_count) = nmax( val1 = result-left_operand->column_count
                                   val2 = result-right_operand->column_count ).

    DATA(target_array) = lcl_excelom_result_array=>create_initial( number_of_rows    = max_row_count
                                                                   number_of_columns = max_column_count ).
    DATA(row_offset) = 0.
    DATA(column_offset) = 0.
    DO max_row_count TIMES.
      DO max_column_count TIMES.
        DATA(left_operand_result_one_cell) = result-left_operand->get_cell_value( column_offset = column_offset
                                                                                  row_offset    = row_offset ).
        DATA(right_operand_result_one_cell) = result-right_operand->get_cell_value( column_offset = column_offset
                                                                                    row_offset    = row_offset ).
*        TRY.
*        DATA(result) = expression2->evaluate( ).
*            left_operand  = lcl_excelom_expr_number=>create(
*                                number = CAST lcl_excelom_result_number( left_operand_result_one_cell )->get_number( ) )
*            right_operand = lcl_excelom_expr_number=>create(
*                                number = CAST lcl_excelom_result_number( right_operand_result_one_cell )->get_number( ) ) ).
*        catch cx_sy_move_cast_error ##NO_HANDLER.
*        endtry.

        DATA(target_array_result_one_cell) = cond #( when expression is bound
                                             then expression->evaluate( context )
                                             ELSE lcl_excelom_result_error=>na_not_applicable ).

        target_array->lif_excelom_result~set_cell_value(
            row_offset    = row_offset
            column_offset = column_offset
            value         = target_array_result_one_cell ).

        column_offset = column_offset + 1.
      ENDDO.

      row_offset = row_offset + 1.
    ENDDO.
    result-result = target_array.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_exprh_lexer IMPLEMENTATION.
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
    result = NEW lcl_excelom_exprh_lexer( ).
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
                                & ')'
            IN text RESULTS DATA(matches).

    complete_with_non_matches( EXPORTING i_string  = text
                               CHANGING  c_matches = matches ).

    DATA(token_values) = value string_table( ).
    LOOP AT matches REFERENCE INTO DATA(match).
      INSERT substring( val = text
                        off = match->offset
                        len = match->length )
             INTO TABLE token_values.
    ENDLOOP.

    TYPES ty_ref_to_parenthesis_group TYPE REF TO ts_parenthesis_group.
    DATA(current_parenthesis_group) = VALUE ty_ref_to_parenthesis_group( ).
    DATA(parenthesis_group) = VALUE ts_parenthesis_group( ).
    DATA(parenthesis_level) = 0.
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
          IF substring( val = token_value->*
                        len = 1 ) = '"'.
            " text literal
            token = VALUE #( value = token_value->*
                             type  = '"' ).
          ELSEIF substring( val = token_value->*
                            len = 1 ) = '['.
            " table argument
            token = VALUE #( value = token_value->*
                             type  = '[' ).
          ELSEIF substring( val = token_value->*
                            len = 1 ) CO '0123456789.-+'.
            " number
            token = VALUE #( value = token_value->*
                             type  = c_type-number ).
          ELSE.
            " function name, --, cell reference, table name, name of named range, constant (TRUE, FALSE)
            TYPES ty_ref_to_string TYPE REF TO string.
            DATA(next_token_value) = COND ty_ref_to_string( WHEN token_number < lines( token_values )
                                                            THEN REF #( token_values[
                                                                            token_number + 1 ] ) ).
            DATA(token_type) = c_type-symbol_name.
            IF next_token_value IS BOUND.
              DATA(next_token_first_character) = substring( val = next_token_value->*
                                                            len = 1 ).
              CASE next_token_first_character.
                WHEN '('.
                  token_type = c_type-function_name.
                WHEN '['.
                  token_type = c_type-table_name.
              ENDCASE.
            ENDIF.
            token = VALUE #( value = token_value->*
                             type  = token_type ).
          ENDIF.
      ENDCASE.

*      CASE token-type.
*        WHEN '('.
*          parenthesis_level = parenthesis_level + 1.
*          INSERT VALUE #( level      = parenthesis_level
*                          from_token = token_number )
*                 INTO TABLE result-parenthesis_groups
*                 REFERENCE INTO current_parenthesis_group.
*        WHEN ','.
*          IF table_specification = abap_false.
*            INSERT VALUE #( level      = parenthesis_level + 1
*                            from_token = cond #( when current_parenthesis_group->last_subgroup_token = 0
*                                                 then current_parenthesis_group->from_token + 1
*                                                 else current_parenthesis_group->last_subgroup_token + 2 )
*                            to_token   = token_number - 1 )
*                   INTO TABLE result-parenthesis_groups.
*            current_parenthesis_group->last_subgroup_token = token_number - 1.
*          ENDIF.
*        WHEN ')'.
*          IF current_parenthesis_group->last_subgroup_token <> 0.
*            INSERT VALUE #( level      = parenthesis_level + 1
*                            from_token = current_parenthesis_group->last_subgroup_token + 2
*                            to_token   = token_number - 1 )
*                   INTO TABLE result-parenthesis_groups.
*          ENDIF.
*          current_parenthesis_group->last_subgroup_token = token_number - 1.
*          current_parenthesis_group->to_token = token_number.
*          parenthesis_level = parenthesis_level - 1.
*          current_parenthesis_group = REF #( result-parenthesis_groups[ level = parenthesis_level ] OPTIONAL ).
*        WHEN '['.
*          table_specification = abap_true.
*        WHEN ']'.
*          table_specification = abap_false.
*      ENDCASE.

      INSERT token INTO TABLE result. "-tokens.
      token_number = token_number + 1.
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_exprh_group IMPLEMENTATION.
  METHOD append.
    INSERT item INTO TABLE items.
  ENDMETHOD.

  METHOD create.
    result = NEW lcl_excelom_exprh_group( ).
    result->type = type.
  ENDMETHOD.

  METHOD delete.
    DELETE items INDEX index.
  ENDMETHOD.

  METHOD insert.
    INSERT item INTO items index index.
  ENDMETHOD.

  METHOD set_expression.
    me->expression = expression.
  ENDMETHOD.

  METHOD set_item_expression.
    items[ index ]-expression = expression.
  ENDMETHOD.

  METHOD set_item_group.
    items[ index ]-group = group.
  ENDMETHOD.

  METHOD set_item_operator.
    items[ index ]-operator = operator.
  ENDMETHOD.

  METHOD set_item_priority.
    items[ index ]-priority = priority.
  ENDMETHOD.

  METHOD set_operator.
    me->operator = operator.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_exprh_operator IMPLEMENTATION.
  METHOD class_constructor.
    LOOP AT VALUE tt_operator(
        ( name = ':'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 1 desc = 'range A1:A2 or A1:A2:A2' )
        ( name = ` `                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 1 desc = 'intersection A1 A2' )
        ( name = ','                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 1 desc = 'union A1,A2' )
        ( name = '-' unary = abap_true operand_position_offsets = VALUE #( ( +1 ) )        priority = 2 desc = '-1' )
        ( name = '+' unary = abap_true operand_position_offsets = VALUE #( ( +1 ) )        priority = 2 desc = '+1' )
        ( name = '%'                   operand_position_offsets = VALUE #( ( -1 ) )        priority = 3 desc = 'percent e.g. 10%' )
        ( name = '^'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 4 desc = 'exponent 2^8' )
        ( name = '*'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 5 desc = '2*2' )
        ( name = '/'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 5 desc = '2/2' )
        ( name = '+'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 6 desc = '2+2' )
        ( name = '-'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 6 desc = '2-2' )
        ( name = '&'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 7 desc = 'concatenate "A"&"B"' )
        ( name = '='                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1=1' )
        ( name = '<'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1<1' )
        ( name = '>'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1>1' )
        ( name = '<='                  operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1<=1' )
        ( name = '>='                  operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1>=1' )
        ( name = '<>'                  operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1<>1' ) )
         REFERENCE INTO DATA(operator).
      create( name                     = operator->name
              operand_position_offsets = operator->operand_position_offsets
              priority                 = operator->priority ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create.
    result = NEW lcl_excelom_exprh_operator( ).
    result->name = name.
    result->operand_position_offsets = operand_position_offsets.
    result->priority = priority.
    INSERT VALUE #( name    = name
                    handler = result )
           INTO TABLE operators.
  ENDMETHOD.

  METHOD create_expression.
    CASE name.
      WHEN '+'.
        result = lcl_excelom_expr_plus=>create( left_operand  = operands[ 1 ]
                                                right_operand = operands[ 2 ] ).
      WHEN '*'.
        result = lcl_excelom_expr_mult=>create( left_operand  = operands[ 1 ]
                                                right_operand = operands[ 2 ] ).
      WHEN OTHERS.
        RAISE EXCEPTION TYPE lcx_excelom_to_do.
    ENDCASE.
  ENDMETHOD.

  METHOD get.
    result = VALUE #( operators[ name  = operator
                                 unary = unary ]-handler OPTIONAL ).
    IF result IS NOT BOUND.
      ASSERT 1 = 1. " Debug helper to set a break-point
    ENDIF.
  ENDMETHOD.

  METHOD get_operand_position_offsets.
    result = operand_position_offsets.
  ENDMETHOD.

  METHOD get_priority.
    result = priority.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_exprh_parser IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_exprh_parser( ).
*    result->formula_cell = formula_cell.
  ENDMETHOD.

  METHOD get_expression_from_symbol_nam.
    IF token_value cp 'true'.
      result = lcl_excelom_expr_boolean=>create( boolean_value = abap_true ).
    elseif token_value cp 'false'.
      result = lcl_excelom_expr_boolean=>create( abap_true ).
    else.
      result = lcl_excelom_expr_range=>create( token_value ).
    endif.
  ENDMETHOD.

  METHOD parse.
    current_token_index = 1.
    me->tokens             = tokens.
    result = parse_expression( ).
  ENDMETHOD.

  METHOD parse_expression.

    " Determine the groups for the parentheses.
    DATA(initial_group) = lcl_excelom_exprh_group=>create( type = '(' ).
    current_token_index = 0.
    previous_token_type = ' '.
    parse_expression_group( group = initial_group ).

    " Determine the groups for the operators.
    parse_expression_group_2( group = initial_group ).

    " Remove useless groups of one item.
    parse_expression_group_3( CHANGING group = initial_group ).

    " Determine the expressions for each group.
    parse_expression_group_4( group = initial_group ).

    result = initial_group->expression.
  ENDMETHOD.

  METHOD parse_expression_group.
    WHILE current_token_index < lines( tokens ).
      current_token_index = current_token_index + 1.
      current_token = REF #( tokens[ current_token_index ] ).
      DATA(ls_item) = VALUE lcl_excelom_exprh_group=>ts_item( token = current_token ).
      CASE current_token->type.
        WHEN '('.
          ls_item-group = lcl_excelom_exprh_group=>create( '(' ).
          previous_token_type = '('.
          parse_expression_group( group = ls_item-group ).
        WHEN ')'.
          previous_token_type = ')'.
          RETURN.
        WHEN 'O'.
          ls_item-operator = lcl_excelom_exprh_operator=>get( operator = current_token->value
                                                              unary    = SWITCH #( previous_token_type
                                                                                   WHEN '(' OR ' ' OR 'O'
                                                                                   THEN abap_true ) ).
      ENDCASE.
      group->append( ls_item ).
      previous_token_type = current_token->type.
    ENDWHILE.
  ENDMETHOD.

  METHOD parse_expression_group_2.
    TYPES to_expression TYPE REF TO lif_excelom_expr.
    TYPES:
      BEGIN OF ts_work,
        position   TYPE sytabix,
        token      TYPE REF TO lcl_excelom_exprh_lexer=>ts_token,
        expression TYPE REF TO lif_excelom_expr,
        operator   TYPE REF TO lcl_excelom_exprh_operator,
        priority   TYPE i,
      END OF ts_work.
    TYPES tt_work TYPE SORTED TABLE OF ts_work WITH NON-UNIQUE KEY position
                    WITH NON-UNIQUE SORTED KEY by_priority COMPONENTS priority position.
    TYPES tt_operand_positions TYPE STANDARD TABLE OF i WITH EMPTY KEY.
    DATA priorities TYPE SORTED TABLE OF i WITH UNIQUE KEY table_line.
    DATA item_index TYPE syst-tabix.

    LOOP AT group->items REFERENCE INTO DATA(item)
        WHERE group IS BOUND.
      parse_expression_group_2( group = item->group ).
    ENDLOOP.

    DATA(work_table) = VALUE tt_work( ).
    LOOP AT group->items REFERENCE INTO item
        WHERE token->type = lcl_excelom_exprh_lexer=>c_type-operator.
      item_index = sy-tabix.
      DATA(priority) = item->operator->get_priority( ).
      group->set_item_priority( index    = item_index
                                priority = priority ).
      INSERT priority INTO TABLE priorities.
    ENDLOOP.

    " Process operators with priority 1 first, then 2, etc.
    " The priority 0 corresponds to functions, tables, boolean values, numeric literals and text literals.
    LOOP AT priorities INTO priority.
      LOOP AT group->items REFERENCE INTO item
           WHERE     token       IS BOUND
                 AND token->type  = lcl_excelom_exprh_lexer=>c_type-operator
                 AND priority     = priority.

        item_index = sy-tabix.
        DATA(operand_position_offsets) = item->operator->get_operand_position_offsets( ).

        DATA(subgroup) = lcl_excelom_exprh_group=>create( type = item->token->type ).
        subgroup->set_operator( item->operator ).
        LOOP AT operand_position_offsets INTO DATA(operand_position_offset).
          subgroup->append( group->items[ item_index + operand_position_offset ] ).
        ENDLOOP.
        group->set_item_group( group = subgroup
                               index = item_index ).

        DATA(positions_of_operands_to_delet) = VALUE tt_operand_positions(
                                                         FOR <operand_position_offset> IN operand_position_offsets
                                                         ( item_index + <operand_position_offset> ) ).
        SORT positions_of_operands_to_delet BY table_line DESCENDING.
        LOOP AT positions_of_operands_to_delet INTO DATA(position).
          group->delete( index = position ).
        ENDLOOP.
      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.

  METHOD parse_expression_group_3.
    LOOP AT group->items REFERENCE INTO DATA(item)
         WHERE group IS BOUND.
      DATA(item_index) = sy-tabix.
      DATA(temp_group) = item->group.
      parse_expression_group_3( CHANGING group = temp_group ).
      IF temp_group <> item->group.
        group->set_item_group( index = item_index
                               group = temp_group ).
      ENDIF.
    ENDLOOP.
    IF     group->type = '('
       AND lines( group->items ) = 1
       AND group->items[ 1 ]-group IS BOUND.
      group = group->items[ 1 ]-group.
    ENDIF.
  ENDMETHOD.

  METHOD parse_expression_group_4.
    TYPES to_expression TYPE REF TO lif_excelom_expr.

    LOOP AT group->items REFERENCE INTO DATA(item).
      DATA(item_index) = sy-tabix.
      IF item->group IS BOUND.
        parse_expression_group_4( group = item->group ).
        group->set_item_expression( index      = item_index
                                    expression = item->group->expression ).
        group->set_item_operator( index    = item_index
                                  operator = item->group->operator ).
      ENDIF.
      IF item->operator IS NOT BOUND.
*        group->set_item_expression( index      = item_index
*                                    expression = item->operator->create_expression(
*                                                     operands = VALUE #( FOR <item> IN item->group->items
*                                                                         ( <item>-expression ) ) ) ).
*      ELSE.
        DATA(expression) = SWITCH to_expression( item->token->type
                                     WHEN lcl_excelom_exprh_lexer=>c_type-text_literal THEN
                                       lcl_excelom_expr_string=>create( item->token->value )
                                     WHEN lcl_excelom_exprh_lexer=>c_type-number THEN
                                       lcl_excelom_expr_number=>create( CONV #( item->token->value ) )
                                     WHEN lcl_excelom_exprh_lexer=>c_type-symbol_name THEN
                                       " range address, range name, boolean true/false
                                       get_expression_from_symbol_nam( item->token->value )
                                     ELSE THROW lcx_excelom_to_do( ) ).
        group->set_item_expression( index      = item_index
                                    expression = expression ).
      ENDIF.
    ENDLOOP.
    IF group->operator IS BOUND.
      group->set_expression( group->operator->create_expression( operands = VALUE #( FOR <item> IN group->items
                                                                                     ( <item>-expression ) ) ) ).
    ENDIF.
  ENDMETHOD.

  METHOD parse_function_arguments.
    result = lcl_excelom_expr_expressions=>create( ).
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
          RAISE EXCEPTION TYPE lcx_excelom_unexpected.
      ENDCASE.
    ENDDO.
  ENDMETHOD.

  METHOD parse_tokens_up_to.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD skip_spaces.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_evaluation_context IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_evaluation_context( ).
    result->containing_cell = containing_cell.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_array IMPLEMENTATION.
  METHOD create.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


class lcl_excelom_expr_boolean implementation.
  METHOD create.
    result = new lcl_excelom_expr_boolean( ).
    result->boolean_value = boolean_value.
  ENDMETHOD.

  method lif_excelom_expr~evaluate.
    result = lcl_excelom_result_boolean=>create( boolean_value ).
  endmethod.

  method lif_excelom_expr~is_equal.
    raise EXCEPTION type lcx_excelom_to_do.
  endmethod.
endclass.


CLASS lcl_excelom_expr_expressions IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_expressions( ).
  ENDMETHOD.

  METHOD append.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_function_call IMPLEMENTATION.
  METHOD create.
    " TODO: parameter NAME is never used (ABAP cleaner)
    " TODO: parameter ARGUMENTS is never used (ABAP cleaner)

    RAISE EXCEPTION TYPE lcx_excelom_to_do.
    result = NEW lcl_excelom_expr_function_call( ).
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_mult IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_mult( ).
    result->left_operand  = left_operand.
    result->right_operand = right_operand.
    result->lif_excelom_expr~type = lif_excelom_expr=>c_type-operation_mult.
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    IF     expression       IS BOUND
       AND expression->type  = lif_excelom_expr=>c_type-operation_mult
       AND left_operand->is_equal( CAST lcl_excelom_expr_mult( expression )->left_operand )
       AND right_operand->is_equal( CAST lcl_excelom_expr_mult( expression )->right_operand ).
      result = abap_true.
    ELSE.
      result = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    DATA(array_evaluation) = lcl_excelom_exprh=>evaluate_array_operands( expression    = me
                                                                         context       = context
                                                                         left_operand  = left_operand
                                                                         right_operand = right_operand ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      IF     array_evaluation-left_operand->type  = lif_excelom_result=>c_type-number
         AND array_evaluation-right_operand->type = lif_excelom_result=>c_type-number.
        result = lcl_excelom_result_number=>create(
                     CAST lcl_excelom_result_number( array_evaluation-left_operand )->get_number( )
                      * CAST lcl_excelom_result_number( array_evaluation-right_operand )->get_number( ) ).
      ELSE.
        RAISE EXCEPTION TYPE lcx_excelom_to_do.
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_number IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_number( ).
    result->number = number.
    result->lif_excelom_expr~type = lif_excelom_expr=>c_type-number.
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    IF     expression->type = lif_excelom_expr=>c_type-number
       AND number           = CAST lcl_excelom_expr_number( expression )->number.
      result = abap_true.
    ELSE.
      result = abap_false.
    endif.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    result = lcl_excelom_result_number=>create( number ).
  ENDMETHOD.
ENDCLASS.


*CLASS lcl_excelom_expr_operation IMPLEMENTATION.
*  METHOD create.
*  ENDMETHOD.
*
*  METHOD lif_excelom_expr_expression~evaluate.
*
*  ENDMETHOD.
*
*  METHOD lif_excelom_expr_operator~get_operand_position_offsets.
*    result = VALUE #( start = -1
*                      end   = +1 ).
*  ENDMETHOD.
*
*  METHOD lif_excelom_expr_operator~get_priority.
*
*  ENDMETHOD.
*
*  METHOD lif_excelom_expr_operator~set_operands.
*
*  ENDMETHOD.
*
*ENDCLASS.


CLASS lcl_excelom_expr_plus IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_plus( ).
    result->left_operand  = left_operand.
    result->right_operand = right_operand.
    result->lif_excelom_expr~type = lif_excelom_expr=>c_type-operation_plus.
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    IF     expression IS BOUND
       AND expression->type = lif_excelom_expr=>c_type-operation_plus
       AND left_operand->is_equal( CAST lcl_excelom_expr_plus( expression )->left_operand )
       AND right_operand->is_equal( CAST lcl_excelom_expr_plus( expression )->right_operand ).
      result = abap_true.
    ELSE.
      result = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    DATA(array_evaluation) = lcl_excelom_exprh=>evaluate_array_operands( expression    = me
                                                                         context       = context
                                                                         left_operand  = left_operand
                                                                         right_operand = right_operand ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      IF     array_evaluation-left_operand->type  = lif_excelom_result=>c_type-number
         AND array_evaluation-right_operand->type = lif_excelom_result=>c_type-number.
        result = lcl_excelom_result_number=>create(
                     CAST lcl_excelom_result_number( array_evaluation-left_operand )->get_number( )
                      + CAST lcl_excelom_result_number( array_evaluation-right_operand )->get_number( ) ).
      ELSE.
        RAISE EXCEPTION TYPE lcx_excelom_to_do.
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_range IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_range( ).
    result->_address_or_name = address_or_name.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    DATA(range) = lcl_excelom_range=>create_from_address_or_name( address     = _address_or_name
                                                                  relative_to = context->containing_cell->parent ).
    result = range->value( ).
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_string IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_string( ).
    result->text = text.
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_sub_expr IMPLEMENTATION.
  METHOD create.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_table IMPLEMENTATION.
  METHOD create.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~is_equal.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_expr~evaluate.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


*CLASS lcl_excelom_formula2 IMPLEMENTATION.
*  METHOD calculate.
*    _expression->evaluate( ).
*  ENDMETHOD.
*
*  METHOD create.
*    result = NEW lcl_excelom_formula2( ).
*    result->range = range.
*    DATA(worksheet) = range->parent.
*    INSERT result INTO TABLE worksheet->formulas.
*  ENDMETHOD.
*
*  METHOD set_value.
*    IF    value                  IS INITIAL
*       OR substring( val = value
*                     len = 1 )    = '='.
*      " Formula error if empty value or "=" at the beginning of the formula
*      RAISE EXCEPTION TYPE lcx_excelom_to_do.
*    ENDIF.
*    DATA(lexer) = lcl_excelom_exprh_lexer=>create( ).
*    DATA(lexer_tokens) = lexer->lexe( value ).
*    DATA(parser) = lcl_excelom_exprh_parser=>create( ).
*    _expression = parser->parse( lexer_tokens ).
*  ENDMETHOD.
*ENDCLASS.


CLASS lcl_excelom_range IMPLEMENTATION.
  METHOD calculate.
    data(context) = lcl_excelom_evaluation_context=>create( containing_cell = me ).
    _formula_expression->evaluate( context ).
  ENDMETHOD.

  METHOD convert_column_a_xfd_to_number.
    DATA(offset) = 0.
    WHILE offset < strlen( roman_letters ).
      FIND roman_letters+offset(1) IN sy-abcde MATCH OFFSET DATA(offset_a_to_z).
      result = ( result * 26 ) + offset_a_to_z + 1.
      IF result > 16384.
        raise EXCEPTION TYPE lcx_excelom_to_do.
      ENDIF.
      offset = offset + 1.
    ENDWHILE.
  ENDMETHOD.

  METHOD create.
    IF cell2 IS INITIAL.
      result = cell1.
    ELSE.
      result = NEW lcl_excelom_range( ).
      result->application = worksheet->application.
      result->_address = VALUE #( top_left     = VALUE #( column = nmin( val1 = cell1->_address-bottom_right-column
                                                                         val2 = cell2->_address-bottom_right-column )
                                                          row    = nmin( val1 = cell1->_address-bottom_right-column
                                                                         val2 = cell2->_address-bottom_right-column ) )
                                  bottom_right = VALUE #( column = nmax( val1 = cell1->_address-bottom_right-column
                                                                         val2 = cell2->_address-bottom_right-column )
                                                          row    = nmax( val1 = cell1->_address-bottom_right-column
                                                                         val2 = cell2->_address-bottom_right-column ) ) ).
    ENDIF.
  ENDMETHOD.

  METHOD create_from_address_or_name.
    result = NEW lcl_excelom_range( ).
    result->parent  = relative_to.
    result->application = relative_to->application.
    result->_address = result->decode_range_address( address ).
  ENDMETHOD.

  METHOD decode_range_address.
    IF application->reference_style = application->c_reference_style-a1.
      result = decode_range_address_a1( address ).
    ELSE.
      result = decode_range_address_r1_c1( address ).
    ENDIF.
    IF result-top_left IS INITIAL.
*     " address is an invalid range address so it's probably referring to an existing name.
*     " Find the name in the current worksheet
*     result-name = parent->parent->names[ worksheet = parent ].
*     " Find the name in the current workbook
*     result-name = parent->parent->names[ worksheet = parent ].
    ENDIF.
  ENDMETHOD.

  METHOD decode_range_address_a1.
    " In the current worksheet:
    "   A1 (relative column and row)
    "   $A1 (absolute column, relative column)
    "   A$1
    "   $A$1
    "   A1:A2
    "   $A$A
    "   A:A
    "   1:1
    "   NAME
    " Other worksheet:
    "   Sheet1!A1
    "   'Sheet 1'!A1
    "   [1]Sheet1!$A$3 (XLSX internal notation for workbooks)
    " Other workbook:
    "   '[C:\workbook.xlsx]'!NAME (workbook absolute path / name in the global scope)
    "   '[workbook.xlsx]Sheet 1'!$A$1 (workbook relative path)
    "   [1]!NAME (XLSX internal notation for workbooks)
    TYPES ty_state TYPE i.

    CONSTANTS:
      BEGIN OF c_state,
        initial                    TYPE ty_state VALUE 1,
        start_of_cell_address      TYPE ty_state VALUE 2,
        "! e.g. decoding A1 in A1:B2
        first_name_of_cell_address TYPE ty_state VALUE 3,
        "! e.g. decoding B2 in A1:B2
        second_name_of_address     TYPE ty_state VALUE 4,
      END OF c_state.

    DATA(end_of_address) = |\n|.
    DATA(address_is_a_name) = abap_false.
    DATA(roman_letters) = ``.
    DATA(numeric_digits) = ``.
    DATA(dollar) = abap_false.

    DATA(state) = c_state-initial.
    DATA(offset) = 0.
    WHILE offset <= strlen( address ).
      IF offset < strlen( address ).
        DATA(character) = substring( val = address
                                     off = offset
                                     len = 1 ).
      ELSE.
        character = end_of_address.
      ENDIF.
      CASE state.
        WHEN c_state-initial.
          CASE character.
            WHEN ''''.
              " other worksheet (in current or other workbook)
              RAISE EXCEPTION TYPE lcx_excelom_to_do.
            WHEN '['.
              " worksheet in other workbook
              RAISE EXCEPTION TYPE lcx_excelom_to_do.
            WHEN '$'.
              " $1:... or $A:... $A1:... or $A$1:...
              IF dollar = abap_true.
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              dollar = abap_true.
              state = c_state-first_name_of_cell_address.
            WHEN 'A' OR 'B' OR 'C' OR 'D' OR 'E' OR 'F' OR 'G' OR 'H' OR 'I' OR 'J' OR 'K' OR 'L' OR 'M'
              OR 'N' OR 'O' OR 'P' OR 'Q' OR 'R' OR 'S' OR 'T' OR 'U' OR 'V' OR 'W' OR 'X' OR 'Y' OR 'Z'.
              roman_letters = character.
              state = c_state-first_name_of_cell_address.
            WHEN '0' OR '1' OR '2' OR '3' OR '4' OR '5' OR '6' OR '7' OR '8' OR '9'.
              numeric_digits = character.
              state = c_state-first_name_of_cell_address.
            WHEN OTHERS.
              RAISE EXCEPTION TYPE lcx_excelom_to_do.
          ENDCASE.

        WHEN c_state-first_name_of_cell_address.
          CASE character.
            WHEN 'A' OR 'B' OR 'C' OR 'D' OR 'E' OR 'F' OR 'G' OR 'H' OR 'I' OR 'J' OR 'K' OR 'L' OR 'M'
              OR 'N' OR 'O' OR 'P' OR 'Q' OR 'R' OR 'S' OR 'T' OR 'U' OR 'V' OR 'W' OR 'X' OR 'Y' OR 'Z'.
              IF dollar = abap_true.
                result-top_left-column_fixed = abap_true.
                dollar = abap_false.
              ENDIF.
              IF numeric_digits IS NOT INITIAL.
                " 1A is invalid
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              roman_letters = roman_letters && character.

            WHEN '0' OR '1' OR '2' OR '3' OR '4' OR '5' OR '6' OR '7' OR '8' OR '9'.
              IF dollar = abap_true.
                result-top_left-row_fixed = abap_true.
                dollar = abap_false.
              ENDIF.
              numeric_digits = numeric_digits && character.

            WHEN '$'.
              IF dollar = abap_true.
                " $$ is invalid
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              dollar = abap_true.

            WHEN ':'
                OR end_of_address.
              IF     roman_letters  IS INITIAL
                 AND numeric_digits IS INITIAL.
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              IF     roman_letters  IS INITIAL
                 AND character = end_of_address.
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              IF     numeric_digits IS INITIAL
                 AND character = end_of_address.
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              IF     roman_letters  IS INITIAL
                 AND numeric_digits IS INITIAL.
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              IF roman_letters IS NOT INITIAL.
                IF     roman_letters NOT BETWEEN 'A' AND 'Z'
                   AND roman_letters NOT BETWEEN 'AA' AND 'ZZ'
                   AND roman_letters NOT BETWEEN 'AAA' AND 'XFD'.
                  address_is_a_name = abap_true.
                  EXIT.
                ENDIF.
                result-top_left-column = convert_column_a_xfd_to_number( roman_letters ).
              ENDIF.
              IF numeric_digits IS NOT INITIAL.
                IF strlen( numeric_digits ) >= 8.
                  address_is_a_name = abap_true.
                  EXIT.
                ENDIF.
                result-top_left-row = numeric_digits.
                IF result-top_left-row > 1048576.
                  address_is_a_name = abap_true.
                  EXIT.
                ENDIF.
              ENDIF.
              roman_letters = VALUE #( ).
              numeric_digits = VALUE #( ).
              state = c_state-second_name_of_address.

            WHEN OTHERS.
              address_is_a_name = abap_true.
              EXIT.
          ENDCASE.

        WHEN c_state-second_name_of_address.
          CASE character.
            WHEN 'A' OR 'B' OR 'C' OR 'D' OR 'E' OR 'F' OR 'G' OR 'H' OR 'I' OR 'J' OR 'K' OR 'L' OR 'M'
              OR 'N' OR 'O' OR 'P' OR 'Q' OR 'R' OR 'S' OR 'T' OR 'U' OR 'V' OR 'W' OR 'X' OR 'Y' OR 'Z'.
              IF numeric_digits IS NOT INITIAL.
                " 1A is invalid
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              roman_letters = roman_letters && character.

            WHEN '0' OR '1' OR '2' OR '3' OR '4' OR '5' OR '6' OR '7' OR '8' OR '9'.
              numeric_digits = numeric_digits && character.

            WHEN '$'.
              " A$1 or $1
              IF     roman_letters                    IS INITIAL
                 AND result-bottom_right-column_fixed  = abap_true.
                " $$A isn't valid
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ELSEIF     numeric_digits                IS INITIAL
                     AND result-bottom_right-row_fixed  = abap_true.
                " $$1 isn't valid
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ELSEIF     roman_letters  IS NOT INITIAL
                     AND numeric_digits IS NOT INITIAL.
                " A1$ isn't valid
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              elseif      roman_letters          IS NOT INITIAL
                      AND roman_letters NOT BETWEEN 'A' AND 'Z'
                      AND roman_letters NOT BETWEEN 'AA' AND 'ZZ'
                      AND roman_letters NOT BETWEEN 'AAA' AND 'XFD'.
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              result-bottom_right-row_fixed = abap_true.

            WHEN end_of_address.
              IF     roman_letters  IS INITIAL
                 AND numeric_digits IS INITIAL.
                RAISE EXCEPTION TYPE lcx_excelom_to_do.
              ENDIF.
              IF     result-bottom_right-column  = 0
                 AND roman_letters          IS NOT INITIAL.
                result-bottom_right-column = convert_column_a_xfd_to_number( roman_letters ).
              ENDIF.
              IF numeric_digits IS NOT INITIAL.
                result-bottom_right-row = numeric_digits.
              ENDIF.
              IF    numeric_digits           IS INITIAL
                 OR strlen( numeric_digits ) >= 8.
                address_is_a_name = abap_true.
                EXIT.
              ENDIF.
              result-bottom_right-row = numeric_digits.
              IF result-bottom_right-row > 1048576.
                address_is_a_name = abap_true.
                EXIT.
              ENDIF.

            WHEN OTHERS.
              " Possibly a name except that the address contains ":" so it's not a name
              RAISE EXCEPTION TYPE lcx_excelom_to_do.
          ENDCASE.

      ENDCASE.
      offset = offset + 1.
    ENDWHILE.

    IF address_is_a_name = abap_false.
      IF result-bottom_right IS INITIAL.
        result-bottom_right = result-top_left.
      ELSE.
        IF     result-top_left-column     IS INITIAL
           AND result-bottom_right-column IS NOT INITIAL.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        ELSEIF     result-top_left-column     IS NOT INITIAL
               AND result-bottom_right-column IS INITIAL.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        ELSEIF     result-top_left-row     IS INITIAL
               AND result-bottom_right-row IS NOT INITIAL.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        ELSEIF     result-top_left-row     IS NOT INITIAL
               AND result-bottom_right-row IS INITIAL.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        ENDIF.
      ENDIF.
      IF     result-top_left-row IS NOT INITIAL
             AND result-top_left-row  > result-bottom_right-row.
        RAISE EXCEPTION TYPE lcx_excelom_to_do.
      ELSEIF     result-top_left-column IS NOT INITIAL
             AND result-top_left-column  > result-bottom_right-column.
        RAISE EXCEPTION TYPE lcx_excelom_to_do.
      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD decode_range_address_r1_c1.

  ENDMETHOD.

  METHOD set_formula2.
    DATA(lexer) = lcl_excelom_exprh_lexer=>create( ).
    DATA(lexer_tokens) = lexer->lexe( value ).
    DATA(parser) = lcl_excelom_exprh_parser=>create( ).
    _formula_expression = parser->parse( lexer_tokens ).
    IF application->calculation = application->c_calculation-automatic.
      _set_value( value = _formula_expression->evaluate( context = lcl_excelom_evaluation_context=>create( containing_cell = me ) ) ).
    ENDIF.
  ENDMETHOD.

  METHOD _set_value.
    IF value IS NOT BOUND.
      RAISE EXCEPTION TYPE lcx_excelom_to_do.
    ENDIF.
    DATA(a) = REF #( parent->_cells[ row    = _address-top_left-row
                                     column = _address-top_left-column ] OPTIONAL ).
    IF a IS NOT BOUND.
      INSERT VALUE #( row    = _address-top_left-row
                      column = _address-top_left-column )
             INTO TABLE parent->_cells
             REFERENCE INTO a.
    ENDIF.
    CASE value->type.
      WHEN value->c_type-number.
        a->value_type = lcl_excelom_worksheet=>c_value_type-number.
        a->value2-double = CAST lcl_excelom_result_number( value )->get_number( ).
        a->value2-string = VALUE #( ).
      WHEN OTHERS.
        RAISE EXCEPTION TYPE lcx_excelom_to_do.
    ENDCASE.
  ENDMETHOD.

  METHOD set_value.
    formula2 = ''.
    _formula_expression = VALUE #( ).
    _set_value( value ).
  ENDMETHOD.

  METHOD value.
    IF     _address-top_left-column = _address-bottom_right-column
       AND _address-top_left-row    = _address-bottom_right-row.
      DATA(cell) = REF #( parent->_cells[ row    = _address-top_left-row
                                          column = _address-top_left-column ] OPTIONAL ).
      IF cell IS NOT BOUND.
        result = lcl_excelom_result_empty=>get_singleton( ).
      ELSE.
        CASE cell->value_type.
          WHEN parent->c_value_type-number.
            result = lcl_excelom_result_number=>create( number = cell->value2-double ).
          WHEN OTHERS.
            RAISE EXCEPTION TYPE lcx_excelom_to_do.
        ENDCASE.
      ENDIF.
    ELSE.
      result = lcl_excelom_result_array=>create_from_range( me ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_range_value IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_range_value( ).
    result->range = range.
  ENDMETHOD.

  METHOD set.
    DATA(row) = range->_address-bottom_right-row.
    WHILE row <= range->_address-bottom_right-row.
      DATA(column) = range->_address-bottom_right-column.
      WHILE column <= range->_address-bottom_right-column.
        DATA(cell) = REF #( range->parent->_cells[ column = column
                                                   row    = row ] OPTIONAL ).
        IF cell IS NOT BOUND.
          INSERT VALUE #( column = column
                          row    = row )
                 INTO TABLE range->parent->_cells
                 REFERENCE INTO cell.
        ENDIF.
        cell->value_type = type.
        CASE type.
          WHEN c_value_type-number.
            cell->value2-double = value.
          WHEN c_value_type-text.
            cell->value2-string = value.
          WHEN OTHERS.
            RAISE EXCEPTION TYPE lcx_excelom_to_do.
        ENDCASE.
        column = column + 1.
      ENDWHILE.
      row = row + 1.
    ENDWHILE.
  ENDMETHOD.

  METHOD set_double.
    set( value = value
         type  = c_value_type-number ).
  ENDMETHOD.

  METHOD set_string.
    set( value = value
         type  = c_value_type-text ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_result_array IMPLEMENTATION.
  METHOD create_from_range.
    result = NEW lcl_excelom_result_array( ).
    result->lif_excelom_result~type = lif_excelom_result=>c_type-array.
*    result->number_of_columns = range->rows( 3 )->count( ).
*    result->range = range.
  ENDMETHOD.

  METHOD create_initial.
    result = NEW lcl_excelom_result_array( ).
    result->number_of_rows    = number_of_rows.
    result->number_of_columns = number_of_columns.
  ENDMETHOD.

  METHOD lif_excelom_result~get_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
*    if range is bound.
*      range->_address-bottom_right-column = .
*      range->_parent->_
*    endif.
  ENDMETHOD.

  METHOD lif_excelom_result~is_array.
    result = abap_true.
  ENDMETHOD.

  METHOD lif_excelom_result~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_error.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_string.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~set_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_result_boolean IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_result_boolean( ).
    result->boolean_value = boolean_value.
  ENDMETHOD.

  METHOD lif_excelom_result~get_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~is_array.
    result = abap_true.
  ENDMETHOD.

  METHOD lif_excelom_result~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_error.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_string.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~set_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_result_empty IMPLEMENTATION.
  METHOD get_singleton.
    if singleton is not bound.
      singleton = new lcl_excelom_result_empty( ).
    endif.
  ENDMETHOD.

  METHOD lif_excelom_result~get_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~is_array.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~is_boolean.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~is_error.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~is_number.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~is_string.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~set_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_result_error IMPLEMENTATION.
  METHOD class_constructor.
    blocked                    = lcl_excelom_result_error=>create( error_name            = '#BLOCKED!     '
                                                                   internal_error_number = 2047
                                                                   formula_error_number  = 11 ).
    calc                       = lcl_excelom_result_error=>create( error_name            = '#CALC!        '
                                                                   internal_error_number = 2050
                                                                   formula_error_number  = 14 ).
    connect                    = lcl_excelom_result_error=>create( error_name            = '#CONNECT!     '
                                                                   internal_error_number = 2046
                                                                   formula_error_number  = 10 ).
    division_by_zero           = lcl_excelom_result_error=>create( error_name            = '#DIV/0!       '
                                                                   internal_error_number = 2007
                                                                   formula_error_number  = 2
                                                                   description           = 'Is produced by =1/0' ).
    field                      = lcl_excelom_result_error=>create( error_name            = '#FIELD!       '
                                                                   internal_error_number = 2049
                                                                   formula_error_number  = 13 ).
    getting_data               = lcl_excelom_result_error=>create( error_name            = '#GETTING_DATA!'
                                                                   internal_error_number = 2043
                                                                   formula_error_number  = 8 ).
    na_not_applicable          = lcl_excelom_result_error=>create( error_name            = '#N/A          '
                                                                   internal_error_number = 2042
                                                                   formula_error_number  = 7
                                                                   description           = 'Is produced by =ERROR.TYPE(1) or if C1 contains =A1:A2+B1:B3 -> C3=#N/A' ).
    name                       = lcl_excelom_result_error=>create( error_name            = '#NAME?        '
                                                                   internal_error_number = 2029
                                                                   formula_error_number  = 5 ).
    null                       = lcl_excelom_result_error=>create( error_name            = '#NULL!        '
                                                                   internal_error_number = 2000
                                                                   formula_error_number  = 1 ).
    num                        = lcl_excelom_result_error=>create( error_name            = '#NUM!         '
                                                                   internal_error_number = 2036
                                                                   formula_error_number  = 6
                                                                   description           = 'Is produced by =1E+240*1E+240' ).
    python                     = lcl_excelom_result_error=>create( error_name            = '#PYTHON!      '
                                                                   internal_error_number = 2222
                                                                   formula_error_number  = 19 ).
    ref                        = lcl_excelom_result_error=>create( error_name            = '#REF!         '
                                                                   internal_error_number = 2023
                                                                   formula_error_number  = 4 ).
    spill                      = lcl_excelom_result_error=>create( error_name            = '#SPILL!       '
                                                                   internal_error_number = 2045
                                                                   formula_error_number  = 9
                                                                   description           = 'Is produced by A1 containing ={1,2} and B1 containing a value -> A1=#SPILL!' ).
    unknown                    = lcl_excelom_result_error=>create( error_name            = '#UNKNOWN!     '
                                                                   internal_error_number = 2048
                                                                   formula_error_number  = 12 ).
    value_cannot_be_calculated = lcl_excelom_result_error=>create( error_name            = '#VALUE!       '
                                                                   internal_error_number = 2015
                                                                   formula_error_number  = 3
                                                                   description           = 'Is produced by =1+"a"' ).
  ENDMETHOD.

  METHOD create.
    result = NEW lcl_excelom_result_error( ).
    result->lif_excelom_result~type = lif_excelom_result=>c_type-error.
    result->error_name            = error_name.
    result->internal_error_number = internal_error_number.
    result->formula_error_number  = formula_error_number.
    result->description           = description.
    INSERT VALUE #( error_name            = error_name
                    internal_error_number = internal_error_number
                    formula_error_number  = formula_error_number
                    handler               = result )
           INTO TABLE errors.
  ENDMETHOD.

  METHOD get_by_error_number.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~is_array.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_error.
    result = abap_true.
  ENDMETHOD.

  METHOD lif_excelom_result~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_string.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~get_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~set_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_result_number IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_result_number( ).
    result->lif_excelom_result~type = lif_excelom_result=>c_type-number.
    result->number                  = number.
  ENDMETHOD.

  METHOD get_number.
    result = number.
  ENDMETHOD.

  METHOD lif_excelom_result~is_array.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_error.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_number.
    result = abap_true.
  ENDMETHOD.

  METHOD lif_excelom_result~is_string.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~get_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.

  METHOD lif_excelom_result~set_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_result_string IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_result_string( ).
    result->lif_excelom_result~type = lif_excelom_result=>c_type-string.
    result->string                  = string.
  ENDMETHOD.

  METHOD lif_excelom_result~get_cell_value.
    result = me.
  ENDMETHOD.

  METHOD lif_excelom_result~is_array.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_error.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD lif_excelom_result~is_string.
    result = abap_true.
  ENDMETHOD.

  METHOD lif_excelom_result~set_cell_value.
    RAISE EXCEPTION TYPE lcx_excelom_to_do.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_worksheet IMPLEMENTATION.
  METHOD calculate.
*    LOOP AT formulas INTO DATA(formula).
*      formula->calculate( ).
*    ENDLOOP.
  ENDMETHOD.

  METHOD create.
    result = NEW lcl_excelom_worksheet( ).
    result->parent = workbook.
    result->application = workbook->application.
  ENDMETHOD.

  METHOD range_from_address.
    DATA(range_1) = lcl_excelom_range=>create_from_address_or_name( address     = cell1
                                                                    relative_to = me ).
    IF cell2 IS INITIAL.
      result = range_1.
    ELSE.
      DATA(range_2) = lcl_excelom_range=>create_from_address_or_name( address     = cell2
                                                                      relative_to = me ).
      result = range_from_two_ranges( cell1 = range_1
                                      cell2 = range_2 ).
    ENDIF.
  ENDMETHOD.

  METHOD range_from_two_ranges.
    result = lcl_excelom_range=>create( worksheet = me
                                        cell1 = cell1
                                        cell2 = cell2 ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_worksheets IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_worksheets( ).
    result->application = workbook->application.
    result->workbook = workbook.
  ENDMETHOD.

  METHOD add.
    DATA worksheet TYPE ty_worksheet.

    worksheet-name   = name.
    worksheet-object = lcl_excelom_worksheet=>create( workbook ).
    INSERT worksheet INTO TABLE worksheets.
    count = count + 1.

    result = worksheet-object.
  ENDMETHOD.

  METHOD item.
    CASE lcl_excelom_application=>type( index ).
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


CLASS lcl_excelom_workbook IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_workbook( ).
    result->application = application.
    result->worksheets = lcl_excelom_worksheets=>create( workbook = result ).
    result->worksheets->add( name = 'Sheet1' ).
  ENDMETHOD.


ENDCLASS.


CLASS lcl_excelom_workbooks IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_workbooks( ).
    result->application = application.
  ENDMETHOD.

  METHOD add.
    DATA workbook TYPE ty_workbook.

    workbook-name   = name.
    workbook-object = lcl_excelom_workbook=>create( application ).
    INSERT workbook INTO TABLE workbooks.
    count = count + 1.

    result = workbook-object.
  ENDMETHOD.



  METHOD item.
    CASE lcl_excelom_application=>type( index ).
      WHEN cl_abap_typedescr=>typekind_string.
        result = workbooks[ name = index ]-object.
      WHEN cl_abap_typedescr=>typekind_int.
        result = workbooks[ index ]-object.
      WHEN OTHERS.
        " TODO
    ENDCASE.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_application IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_application( ).
    result->workbooks = lcl_excelom_workbooks=>create( result ).
  ENDMETHOD.

  METHOD calculate.
    DATA(workbook_number) = 1.
    WHILE workbook_number <= workbooks->count.
      DATA(workbook) = workbooks->item( workbook_number ).

      DATA(worksheet_number) = 1.
      WHILE worksheet_number <= workbook->worksheets->count.
        DATA(worksheet) = workbook->worksheets->item( worksheet_number ).
        worksheet->calculate( ).
      ENDWHILE.
    ENDWHILE.
  ENDMETHOD.

  METHOD type.
    DESCRIBE FIELD any_data_object TYPE result.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_evaluate DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS one_plus_one FOR TESTING RAISING cx_static_check.
    METHODS test31 FOR TESTING RAISING cx_static_check.

    TYPES tt_parenthesis_group TYPE lcl_excelom_exprh_lexer=>tt_parenthesis_group.
    TYPES tt_token             TYPE lcl_excelom_exprh_lexer=>tt_token.
    TYPES ts_result_lexe       TYPE lcl_excelom_exprh_lexer=>ts_result_lexe.

    CONSTANTS c_type LIKE lcl_excelom_exprh_lexer=>c_type VALUE lcl_excelom_exprh_lexer=>c_type.
    DATA: worksheet TYPE REF TO lcl_excelom_worksheet.



    METHODS lexe
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE tt_token. "ts_result_lexe.

    METHODS parse
      IMPORTING !tokens            TYPE lcl_excelom_exprh_lexer=>tt_token
      RETURNING VALUE(result)      TYPE REF TO lif_excelom_expr
      RAISING   lcx_excelom_expr_parser.

    METHODS setup.

ENDCLASS.


CLASS ltc_lexer DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS function FOR TESTING RAISING cx_static_check.
    METHODS number FOR TESTING RAISING cx_static_check.
    METHODS range FOR TESTING RAISING cx_static_check.
    METHODS text_literal FOR TESTING RAISING cx_static_check.
    METHODS text_literal_with_double_quote FOR TESTING RAISING cx_static_check.
    METHODS smart_table FOR TESTING RAISING cx_static_check.
    METHODS smart_table_all FOR TESTING RAISING cx_static_check.
    METHODS smart_table_column FOR TESTING RAISING cx_static_check.
    METHODS smart_table_no_space FOR TESTING RAISING cx_static_check.
    METHODS smart_table_space_separator FOR TESTING RAISING cx_static_check.
    METHODS smart_table_space_boundaries FOR TESTING RAISING cx_static_check.
    METHODS smart_table_space_all FOR TESTING RAISING cx_static_check.
    METHODS very_long FOR TESTING RAISING cx_static_check.
    METHODS arithmetic FOR TESTING RAISING cx_static_check.

    TYPES tt_parenthesis_group TYPE lcl_excelom_exprh_lexer=>tt_parenthesis_group.
    TYPES tt_token             TYPE lcl_excelom_exprh_lexer=>tt_token.
    TYPES ts_result_lexe       TYPE lcl_excelom_exprh_lexer=>ts_result_lexe.

    CONSTANTS c_type LIKE lcl_excelom_exprh_lexer=>c_type VALUE lcl_excelom_exprh_lexer=>c_type.

    METHODS lexe
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE tt_token."ts_result_lexe.

ENDCLASS.


CLASS ltc_parser DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.

    METHODS one_plus_one FOR TESTING RAISING cx_static_check.
    METHODS very_long  FOR TESTING RAISING cx_static_check.
*    METHODS test31 FOR TESTING RAISING cx_static_check.
    METHODS parentheses_arithmetic FOR TESTING RAISING cx_static_check.
    METHODS parentheses_arithmetic_complex FOR TESTING RAISING cx_static_check.

    TYPES tt_token       TYPE lcl_excelom_exprh_lexer=>tt_token.
    TYPES ts_result_lexe TYPE lcl_excelom_exprh_lexer=>ts_result_lexe.

    CONSTANTS c_type LIKE lcl_excelom_exprh_lexer=>c_type VALUE lcl_excelom_exprh_lexer=>c_type.

    METHODS assert_equals
      IMPORTING act            TYPE REF TO lif_excelom_expr
                exp            TYPE REF TO lif_excelom_expr
      RETURNING VALUE(result)  TYPE REF TO lif_excelom_result.



    METHODS lexe
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE tt_token. "ts_result_lexe.

    METHODS parse
      IMPORTING !tokens            TYPE lcl_excelom_exprh_lexer=>tt_token
      RETURNING VALUE(result)      TYPE REF TO lif_excelom_expr
      RAISING   lcx_excelom_expr_parser.

    METHODS get_texts_from_matches
      IMPORTING i_string      TYPE string
                i_matches     TYPE match_result_tab
      RETURNING VALUE(result) TYPE string_table.

ENDCLASS.


CLASS ltc_range DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PUBLIC SECTION.
    INTERFACES lif_excelom_all_friends.

  PRIVATE SECTION.
    METHODS convert_column_a_xfd_to_number FOR TESTING RAISING cx_static_check.
    METHODS decode_range_address_a1_invali FOR TESTING RAISING cx_static_check.
    METHODS decode_range_address_a1_valid  FOR TESTING RAISING cx_static_check.

    TYPES ty_address TYPE lcl_excelom_range=>ty_address.
ENDCLASS.


CLASS ltc_evaluate IMPLEMENTATION.
  METHOD lexe.
    DATA(lexer) = lcl_excelom_exprh_lexer=>create( ).
    result = lexer->lexe( text ).
  ENDMETHOD.

  METHOD one_plus_one.
    DATA(range) = worksheet->range_from_address( 'A1' ).
    range->set_formula2( value = `1+1` ).
    cl_abap_unit_assert=>assert_true( range->value( )->is_number( ) ).
    cl_abap_unit_assert=>assert_equals( act = CAST lcl_excelom_result_number( range->value( ) )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD parse.
    result = lcl_excelom_exprh_parser=>create( )->parse( tokens ).
  ENDMETHOD.

  METHOD setup.
    DATA(app) = lcl_excelom_application=>create( ).
    DATA(workbook) = app->workbooks->add( 'name' ).
    worksheet = workbook->worksheets->item( 'Sheet1' ).
  ENDMETHOD.

  METHOD test31.
    worksheet->range_from_address( 'A1' )->set_value( lcl_excelom_result_number=>create( 10 ) ).
    DATA(range) = worksheet->range_from_address( 'A2' ).
    range->set_formula2( 'A1+1' ).
    "range->calculate( ).
    cl_abap_unit_assert=>assert_equals( act = CAST lcl_excelom_result_number( range->value( ) )->get_number( )
                                        exp = 11 ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_lexer IMPLEMENTATION.
  METHOD arithmetic.
    cl_abap_unit_assert=>assert_equals( act = lexe( '2*(1+3*(5+1))' )
                                        exp = VALUE tt_token( ( value = `2`  type = c_type-number )
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
*                                    parenthesis_groups = VALUE #( ( from_token = 3 to_token = 13 level = 1 last_subgroup_token = 12 )
*                                                                  ( from_token = 8 to_token = 12 level = 2 last_subgroup_token = 11 ) ) ) ).
  ENDMETHOD.

  METHOD function.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'IF(1=1,0,1)' ) "-tokens
                                        exp = VALUE tt_token( ( value = `IF` type = c_type-function_name )
                                                              ( value = `(`  type = '(' )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `=`  type = c_type-operator )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `0`  type = c_type-number )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `1`  type = c_type-number )
                                                              ( value = `)`  type = ')' ) ) ).
  ENDMETHOD.

  METHOD lexe.
    DATA(lexer) = lcl_excelom_exprh_lexer=>create( ).
    result = lexer->lexe( text ).
  ENDMETHOD.

  METHOD number.
    cl_abap_unit_assert=>assert_equals( act = lexe( '25' )
                                        exp = VALUE tt_token( ( value = `25` type = c_type-number ) ) ).
  ENDMETHOD.

  METHOD range.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Sheet1!$A$1' )
                                        exp = VALUE tt_token( ( value = `Sheet1!$A$1` type = 'W' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( `'Sheet 1'!$A$1` )
                                        exp = VALUE tt_token( ( value = `'Sheet 1'!$A$1` type = 'W' ) ) ).
  ENDMETHOD.

  METHOD smart_table.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[]' )
                                        exp = VALUE tt_token( ( value = `Table1` type = c_type-table_name )
                                                              ( value = `[`      type = `[` )
                                                              ( value = `]`      type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_all.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[[#All]]' )
                                        exp = VALUE tt_token( ( value = `Table1` type = c_type-table_name )
                                                              ( value = `[`      type = `[` )
                                                              ( value = `[#All]` type = `[` )
                                                              ( value = `]`      type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_column.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[Column1]' )
                                        exp = VALUE tt_token( ( value = `Table1`    type = c_type-table_name )
                                                              ( value = `[Column1]` type = `[` ) ) ).
  ENDMETHOD.

  METHOD smart_table_no_space.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[[#Headers],[#Data],[% Commission]]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_space_all.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[ [#Headers], [#Data], [% Commission] ]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_space_boundaries.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[ [#Headers],[#Data],[% Commission] ]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD smart_table_space_separator.
    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[[#Headers], [#Data], [% Commission]]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD text_literal.
    cl_abap_unit_assert=>assert_equals( act = lexe( '"IF(1=1,0,1)"' )
                                        exp = VALUE tt_token( ( value = `"IF(1=1,0,1)"` type = '"' ) ) ).
  ENDMETHOD.

  METHOD text_literal_with_double_quote.
    cl_abap_unit_assert=>assert_equals( act = lexe( '"IF(A1=""X"",0,1)"' )
                                        exp = VALUE tt_token( ( value = `"IF(A1=""X"",0,1)"` type = '"' ) ) ).
  ENDMETHOD.

  METHOD very_long.
    cl_abap_unit_assert=>assert_equals( act = lexe( |(a{ repeat( val = ',a'
                                                                 occ = 5000 )
                                                    })| )
                                        exp = VALUE tt_token( ( value = `(` type = '(' )
                                                              ( value = `a` type = 'W' )
                                                              ( LINES OF VALUE
                                                                tt_token( FOR i = 1 WHILE i <= 5000
                                                                          ( value = `,` type = ',' )
                                                                          ( value = `a` type = 'W' ) ) )
                                                              ( value = `)` type = ')' ) ) ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_parser IMPLEMENTATION.
  METHOD assert_equals.
    cl_abap_unit_assert=>assert_true( xsdbool( exp->is_equal( act ) ) ).
  ENDMETHOD.

  METHOD get_texts_from_matches.
    LOOP AT i_matches REFERENCE INTO DATA(match).
      APPEND substring( val = i_string
                        off = match->offset
                        len = match->length ) TO result.
    ENDLOOP.
  ENDMETHOD.

  METHOD lexe.
    DATA(lexer) = lcl_excelom_exprh_lexer=>create( ).
    result = lexer->lexe( text ).
  ENDMETHOD.

  METHOD one_plus_one.
    assert_equals( act = parse( tokens = VALUE #( ( value = `1`  type = c_type-number )
                                                  ( value = `+`  type = c_type-operator )
                                                  ( value = `1`  type = c_type-number ) ) )
                   exp = lcl_excelom_expr_plus=>create( left_operand  = lcl_excelom_expr_number=>create( 1 )
                                                        right_operand = lcl_excelom_expr_number=>create( 1 ) ) ).
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
    DATA(exp) = lcl_excelom_expr_mult=>create( left_operand  = lcl_excelom_expr_number=>create( 2 )
                                               right_operand = lcl_excelom_expr_plus=>create(
                                                   left_operand  = lcl_excelom_expr_number=>create( 1 )
                                                   right_operand = lcl_excelom_expr_number=>create( 3 ) ) ).
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
    DATA(exp) = lcl_excelom_expr_mult=>create(
                    left_operand  = lcl_excelom_expr_number=>create( 2 )
                    right_operand = lcl_excelom_expr_plus=>create(
                        left_operand  = lcl_excelom_expr_number=>create( 1 )
                        right_operand = lcl_excelom_expr_mult=>create(
                                            left_operand  = lcl_excelom_expr_number=>create( 3 )
                                            right_operand = lcl_excelom_expr_plus=>create(
                                                left_operand  = lcl_excelom_expr_number=>create( 5 )
                                                right_operand = lcl_excelom_expr_number=>create( 1 ) ) ) ) ).
    assert_equals( act = act
                   exp = exp ).
  ENDMETHOD.

  METHOD parse.
    result = lcl_excelom_exprh_parser=>create( )->parse( tokens ).
  ENDMETHOD.

  METHOD very_long.
    cl_abap_unit_assert=>fail( msg   = 'TO DO'
                               level = if_aunit_constants=>tolerable ).
*    DATA(a) = parse(
*        lexe(
*            `IFERROR(IF(C2<>"",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assig` &&
*`ned Attorney, or Sales Team",B2<>"Jimmy Edwards",B2<>"Kathleen McCarthy"),B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assigned Attorney, or Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(VL` &&
*`OOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(C2<>"",VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1` &&
*            `!$A:$B,2,FALSE),"INTAKE TEAM")))))), VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE),"")` ) ).
  ENDMETHOD.
ENDCLASS.


CLASS ltc_range IMPLEMENTATION.
  METHOD convert_column_a_xfd_to_number.
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>convert_column_a_xfd_to_number( roman_letters = 'XFD' )
                                        exp = 16384 ).

    TRY.
        lcl_excelom_range=>convert_column_a_xfd_to_number( roman_letters = 'XFE' ).
        cl_abap_unit_assert=>fail( msg = 'Exception expected for XFE - Column does not exist' ).
      CATCH cx_root ##NO_HANDLER.
    ENDTRY.

    TRY.
        lcl_excelom_range=>convert_column_a_xfd_to_number( roman_letters = 'ZZZZ' ).
        cl_abap_unit_assert=>fail( msg = 'Exception expected for XFE - Column does not exist' ).
      CATCH cx_root ##NO_HANDLER.
    ENDTRY.
  ENDMETHOD.

  METHOD decode_range_address_a1_invali.
    LOOP AT VALUE string_table( ( `:` ) ( `` ) ( `$` ) ( `A` ) ( `A:` ) ( `$$A1` ) ( `A:A1` ) ( `B2:A1` ) ) INTO DATA(address).
      TRY.
          lcl_excelom_range=>decode_range_address_a1( address ).
          cl_abap_unit_assert=>fail( msg = |Exception expected for address "{ address }"| ).
        CATCH cx_root ##NO_HANDLER.
      ENDTRY.
    ENDLOOP.
  ENDMETHOD.

  METHOD decode_range_address_a1_valid.
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>decode_range_address_a1( 'A1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1
                                                                                        row    = 1 )
                                                                bottom_right = VALUE #( column = 1
                                                                                        row    = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>decode_range_address_a1( 'A$1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column    = 1
                                                                                        row       = 1
                                                                                        row_fixed = abap_true )
                                                                bottom_right = VALUE #( column    = 1
                                                                                        row       = 1
                                                                                        row_fixed = abap_true ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>decode_range_address_a1( '$A1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1 )
                                                                bottom_right = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>decode_range_address_a1( '$A$1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1
                                                                                        row_fixed    = abap_true )
                                                                bottom_right = VALUE #( column       = 1
                                                                                        column_fixed = abap_true
                                                                                        row          = 1
                                                                                        row_fixed    = abap_true ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>decode_range_address_a1( 'A1:B1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1
                                                                                        row    = 1 )
                                                                bottom_right = VALUE #( column = 2
                                                                                        row    = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>decode_range_address_a1( 'A:A' )
                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1 )
                                                                bottom_right = VALUE #( column = 1 ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lcl_excelom_range=>decode_range_address_a1( '1:1' )
                                        exp = VALUE ty_address( top_left     = VALUE #( row = 1 )
                                                                bottom_right = VALUE #( row = 1 ) ) ).
  ENDMETHOD.
ENDCLASS.
