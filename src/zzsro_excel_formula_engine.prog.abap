REPORT zzsro_excel_formula_engine.
*" EXCEL OBJECT MODEL (e.g. the one you can see in Excel VBA)
*
*" https://github.com/sandraros/excelom
*" EXCEL OBJECT MODEL (e.g. the one you can see in Excel VBA)
*" https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model
*
*"-------------------------------------------------------------------------------------------------------------------------
*"
*" Excel saves formulas always with A1 reference-style addresses (even if displayed/entered in R1C1 style).
*"
*"-------------------------------------------------------------------------------------------------------------------------
*"
*" AND : logical boolean function
*" CHOOSE : conditional function, the first argument is an integer, if it's 1 the result will be the second argument, if it's 2 the result will be the third argument, etc.
*" COLUMN
*" CONCATENATE
*" FILTER
*" FLOOR.MATH
*" IFS
*" LEFT
*" MID
*" MOD
*" VLOOKUP
*
*CLASS lcl_xlom                    DEFINITION DEFERRED.
*CLASS lcl_xlom_application        DEFINITION DEFERRED.
*CLASS lcl_xlom_range              DEFINITION DEFERRED.
*CLASS lcl_xlom_sheet              DEFINITION DEFERRED.
*CLASS lcl_xlom_workbook           DEFINITION DEFERRED.
*CLASS lcl_xlom_workbooks          DEFINITION DEFERRED.
*CLASS lcl_xlom_worksheet          DEFINITION DEFERRED.
*CLASS lcl_xlom_worksheets         DEFINITION DEFERRED.
*
*CLASS lcl_xlom__ex_ut             DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_ut_lexer       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_ut_operator    DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_ut_parser      DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_ut_parser_item DEFINITION DEFERRED.
*
*CLASS lcl_xlom__ex_el_array       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_el_boolean     DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_el_empty_arg   DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_el_error       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_el_number      DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_el_range       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_el_string      DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_address     DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_cell        DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_find        DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_if          DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_iferror     DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_index       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_indirect    DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_len         DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_match       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_offset      DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_right       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_row         DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_fu_t           DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_op_ampersand   DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_op_colon       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_op_equal       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_op_minus       DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_op_mult        DEFINITION DEFERRED.
*CLASS lcl_xlom__ex_op_plus        DEFINITION DEFERRED.
*
*CLASS lcl_xlom__ut_eval_context   DEFINITION DEFERRED.
*
*CLASS lcl_xlom__va                DEFINITION DEFERRED.
*CLASS lcl_xlom__va_array          DEFINITION DEFERRED.
*CLASS lcl_xlom__va_boolean        DEFINITION DEFERRED.
*CLASS lcl_xlom__va_empty          DEFINITION DEFERRED.
*CLASS lcl_xlom__va_error          DEFINITION DEFERRED.
*CLASS lcl_xlom__va_number         DEFINITION DEFERRED.
*CLASS lcl_xlom__va_string         DEFINITION DEFERRED.
*
*CLASS lcx_xlom__ex_ut_parser      DEFINITION DEFERRED.
*CLASS lcx_xlom__va                DEFINITION DEFERRED.
*CLASS lcx_xlom_todo               DEFINITION DEFERRED.
*CLASS lcx_xlom_unexpected         DEFINITION DEFERRED.
*
*INTERFACE lif_xlom__ex_array       DEFERRED.
*INTERFACE lif_xlom__va_array       DEFERRED.
*INTERFACE lif_xlom__ut_all_friends DEFERRED.
*INTERFACE lif_xlom__ex             DEFERRED.
*INTERFACE lif_xlom__va             DEFERRED.
*
*CLASS ltc_evaluate                DEFINITION DEFERRED.
*CLASS ltc_lexer                   DEFINITION DEFERRED.
*CLASS ltc_parser                  DEFINITION DEFERRED.
*CLASS ltc_range                   DEFINITION DEFERRED.
*
*
*CLASS lcx_xlom__ex_ut_parser DEFINITION INHERITING FROM cx_static_check.
*  PUBLIC SECTION.
*    METHODS constructor
*      IMPORTING !text     TYPE csequence OPTIONAL
*                msgv1     TYPE csequence OPTIONAL
*                msgv2     TYPE csequence OPTIONAL
*                msgv3     TYPE csequence OPTIONAL
*                msgv4     TYPE csequence OPTIONAL
*                textid    LIKE textid    OPTIONAL
*                !previous LIKE previous  OPTIONAL.
*
*    METHODS get_text     REDEFINITION.
*    METHODS get_longtext REDEFINITION.
*
*  PRIVATE SECTION.
*    DATA text  TYPE string.
*    DATA msgv1 TYPE string.
*    DATA msgv2 TYPE string.
*    DATA msgv3 TYPE string.
*    DATA msgv4 TYPE string.
*ENDCLASS.
*
*
*CLASS lcx_xlom__va DEFINITION INHERITING FROM cx_static_check.
*  PUBLIC SECTION.
*    DATA result_error TYPE ref to lcl_xlom__va_error READ-ONLY.
*    METHODS constructor
*      IMPORTING result_error TYPE ref to lcl_xlom__va_error.
*ENDCLASS.
*
*
*CLASS lcx_xlom_todo DEFINITION INHERITING FROM cx_no_check.
*ENDCLASS.
*
*
*CLASS lcx_xlom_unexpected DEFINITION INHERITING FROM cx_no_check.
*ENDCLASS.
*
*
*INTERFACE lif_xlom__ut_all_friends.
*ENDINTERFACE.
*
*
*CLASS lcl_xlom DEFINITION.
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*    "! xlApplicationInternational
*    TYPES ty_application_international TYPE i.
*    "! xlCalculation
*    TYPES ty_calculation TYPE i.
*    "! not an Excel constant
*    TYPES ty_country TYPE i.
*    "! xlReferenceType
*    TYPES ty_reference_style TYPE i.
*
*    TYPES:
*      BEGIN OF ts_range_address_one_cell,
*        column TYPE i,
*        row    TYPE i,
*      END OF ts_range_address_one_cell.
*    TYPES:
*      BEGIN OF ts_range_address,
*        top_left     TYPE ts_range_address_one_cell,
*        bottom_right TYPE ts_range_address_one_cell,
*      END OF ts_range_address.
*
*    CONSTANTS:
*      BEGIN OF c_application_international,
*        country_code TYPE ty_application_international VALUE 1,
*      END OF c_application_international.
*
*    CONSTANTS:
*      BEGIN OF c_calculation,
*        automatic     TYPE ty_calculation VALUE -4105,
*        manual        TYPE ty_calculation VALUE -4135,
*        semiautomatic TYPE ty_calculation VALUE 2,
*      END OF c_calculation.
*
*    CONSTANTS:
*      BEGIN OF c_country,
*        brazil         TYPE ty_country VALUE 55,
*        czech_republic TYPE ty_country VALUE 420,
*        denmark        TYPE ty_country VALUE 45,
*        estonia        TYPE ty_country VALUE 372,
*        finland        TYPE ty_country VALUE 358,
*        france         TYPE ty_country VALUE 33,
*        germany        TYPE ty_country VALUE 49,
*        greece         TYPE ty_country VALUE 30,
*        hungary        TYPE ty_country VALUE 36,
*        indonesia      TYPE ty_country VALUE 62,
*        italy          TYPE ty_country VALUE 39,
*        japan          TYPE ty_country VALUE 81,
*        malaysia       TYPE ty_country VALUE 60,
*        netherlands    TYPE ty_country VALUE 31,
*        norway         TYPE ty_country VALUE 47,
*        poland         TYPE ty_country VALUE 48,
*        portugal       TYPE ty_country VALUE 351,
*        russia         TYPE ty_country VALUE 7,
*        slovenia       TYPE ty_country VALUE 386,
*        spain          TYPE ty_country VALUE 34,
*        sweden         TYPE ty_country VALUE 46,
*        turkey         TYPE ty_country VALUE 90,
*        ukraine        TYPE ty_country VALUE 380,
*        usa            TYPE ty_country VALUE 1,
*      END OF c_country.
*
*    CONSTANTS:
*      BEGIN OF c_reference_style,
*        a1    TYPE ty_reference_style VALUE 1,
*        r1_c1 TYPE ty_reference_style VALUE -4150,
*      END OF c_reference_style.
*ENDCLASS.
*
*
*INTERFACE lif_xlom__ex.
*  TYPES ty_expression_type TYPE i.
*  TYPES:
*    BEGIN OF ts_operand_result,
*      name                     TYPE string,
*      object                   TYPE REF TO lif_xlom__va,
*      "! <ul>
*      "! <li>'X': the argument isn't changed when the formula is expanded for Array Evaluation
*      "! e.g. the argument Array of the function INDEX: if A1 contains =INDEX(C1:D2,{1,2},{1,2}),
*      "! A2 and A2 values are the same as if they contain =INDEX(C1:D2,1,1) and =INDEX(C1:D2,2,2).</li>
*      "! <li>' ': the argument is changed when the formula is expanded for Array Evaluation
*      "! e.g.the argument Text of the function RIGHT: if A1 contains =RIGHT(A1:A2,{1;2}),
*      "! A1 and A2 values are the same as if they contain =RIGHT(A1,1) and =RIGHT(A2,2).</li>
*      "! </ul>
*      not_part_of_result_array TYPE abap_bool,
*    END OF ts_operand_result.
*  TYPES tt_operand_result TYPE SORTED TABLE OF ts_operand_result WITH UNIQUE KEY name.
*  TYPES:
*    BEGIN OF ts_operand_expr,
*      name                     TYPE string,
*      object                   TYPE REF TO lif_xlom__ex,
*      not_part_of_result_array TYPE abap_bool,
*    END OF ts_operand_expr.
*  TYPES tt_operand_expr TYPE SORTED TABLE OF ts_operand_expr WITH UNIQUE KEY name.
*  TYPES:
*    BEGIN OF ts_evaluate_array_operands,
*      result          TYPE REF TO lif_xlom__va,
*      operand_results TYPE tt_operand_result,
*    END OF ts_evaluate_array_operands.
*
*  CONSTANTS:
*    BEGIN OF c_type,
*      array          TYPE ty_expression_type VALUE 1,
*      boolean        TYPE ty_expression_type VALUE 2,
*      empty_argument TYPE ty_expression_type VALUE 3,
*      error          TYPE ty_expression_type VALUE 4,
*      number         TYPE ty_expression_type VALUE 5,
*      range          TYPE ty_expression_type VALUE 6,
*      string         TYPE ty_expression_type VALUE 7,
*      BEGIN OF function,
*        address  TYPE ty_expression_type VALUE 100,
*        cell     TYPE ty_expression_type VALUE 101,
*        countif  TYPE ty_expression_type VALUE 102,
*        find     TYPE ty_expression_type VALUE 103,
*        if       TYPE ty_expression_type VALUE 104,
*        iferror  TYPE ty_expression_type VALUE 105,
*        index    TYPE ty_expression_type VALUE 106,
*        indirect TYPE ty_expression_type VALUE 107,
*        len      TYPE ty_expression_type VALUE 108,
*        match    TYPE ty_expression_type VALUE 109,
*        offset   TYPE ty_expression_type VALUE 110,
*        right    TYPE ty_expression_type VALUE 111,
*        row      TYPE ty_expression_type VALUE 112,
*        t        TYPE ty_expression_type VALUE 113,
*      END OF function,
*      BEGIN OF operation,
*        ampersand   TYPE ty_expression_type VALUE 10,
*        equal       TYPE ty_expression_type VALUE 11,
*        minus       TYPE ty_expression_type VALUE 12,
*        minus_unary TYPE ty_expression_type VALUE 13,
*        mult        TYPE ty_expression_type VALUE 14,
*        plus        TYPE ty_expression_type VALUE 15,
*      END OF operation,
*    END OF c_type.
*
*  DATA type                 TYPE ty_expression_type        READ-ONLY.
*  DATA result_of_evaluation TYPE REF TO lif_xlom__va READ-ONLY.
*
*  METHODS evaluate
*    IMPORTING !context      TYPE REF TO lcl_xlom__ut_eval_context
*    RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*
*  METHODS evaluate_single
*    IMPORTING arguments     TYPE tt_operand_result
*              !context      TYPE REF TO lcl_xlom__ut_eval_context
*    RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*
*  METHODS is_equal
*    IMPORTING expression    TYPE REF TO lif_xlom__ex
*    RETURNING VALUE(result) TYPE abap_bool.
*
*  METHODS set_result
*    IMPORTING value         TYPE REF TO lif_xlom__va
*    RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*ENDINTERFACE.
*
*
*INTERFACE lif_xlom__ex_array.
*  INTERFACES lif_xlom__ex.
*
*  DATA row_count    TYPE i       READ-ONLY.
*  DATA column_count TYPE i       READ-ONLY.
*ENDINTERFACE.
*
*
*INTERFACE lif_xlom__va.
*  TYPES ty_type TYPE i.
*
*  CONSTANTS:
*    BEGIN OF c_type,
*      boolean TYPE ty_type VALUE 1,
*      array   TYPE ty_type VALUE 2,
*      empty   TYPE ty_type VALUE 3,
*      error   TYPE ty_type VALUE 4,
*      number  TYPE ty_type VALUE 5,
*      range   TYPE ty_type VALUE 6,
*      string  TYPE ty_type VALUE 7,
*    END OF c_type.
*
*  DATA type         TYPE ty_type READ-ONLY.
*
*  METHODS get_value
*    RETURNING VALUE(result) TYPE REF TO data.
*
*  METHODS is_array
*    RETURNING VALUE(result) TYPE abap_bool.
*
*  METHODS is_boolean
*    RETURNING VALUE(result) TYPE abap_bool.
*
*  "! Checks whether the current result has the exact same type as the input result,
*  "! and the same values. For instance, lcl_xlom__va_string=>create( '1'
*  "! )->is_equal( lcl_xlom__va_string=>create( '1' ) ) is true, but
*  "! lcl_xlom__va_number=>create( 1
*  "! )->is_equal( lcl_xlom__va_string=>create( '1' ) ) is false.
*  METHODS is_equal
*    IMPORTING input_result  TYPE REF TO lif_xlom__va
*    RETURNING VALUE(result) TYPE abap_bool.
*
*  METHODS is_error
*    RETURNING VALUE(result) TYPE abap_bool.
*
*  METHODS is_number
*    RETURNING VALUE(result) TYPE abap_bool.
*
*  METHODS is_string
*    RETURNING VALUE(result) TYPE abap_bool.
*ENDINTERFACE.
*
*
*INTERFACE lif_xlom__va_array.
*  INTERFACES lif_xlom__va.
*
*  TYPES tt_column TYPE STANDARD TABLE OF REF TO lif_xlom__va WITH EMPTY KEY.
*  TYPES:
*    BEGIN OF ts_row,
*      columns_of_row TYPE tt_column,
*    END OF ts_row.
*  TYPES tt_row TYPE STANDARD TABLE OF ts_row WITH EMPTY KEY.
*  TYPES:
*    BEGIN OF ts_address_one_cell,
*      "! 0 means that the address is the whole row defined in ROW
*      column       TYPE i,
*      column_fixed TYPE abap_bool,
*      "! 0 means that the address is the whole column defined in COLUMN
*      row          TYPE i,
*      row_fixed    TYPE abap_bool,
*    END OF ts_address_one_cell.
*  TYPES:
*    BEGIN OF ts_address,
*      "! Can also be an internal ID like "1" ([1]Sheet1!A1)
*      workbook_name  TYPE string,
*      worksheet_name TYPE string,
*      range_name     TYPE string,
*      top_left       TYPE ts_address_one_cell,
*      bottom_right   TYPE ts_address_one_cell,
*    END OF ts_address.
*
*  DATA row_count    TYPE i READ-ONLY.
*  DATA column_count TYPE i READ-ONLY.
*
*  METHODS get_array_value
**    IMPORTING top_left      TYPE ts_address_one_cell
**              bottom_right  TYPE ts_address_one_cell
*    IMPORTING top_left      TYPE lcl_xlom=>ts_range_address_one_cell
*              bottom_right  TYPE lcl_xlom=>ts_range_address_one_cell
*    RETURNING VALUE(result) TYPE REF TO lif_xlom__va_array.
*
*  "!
*  "! @parameter column | Start from 1
*  "! @parameter row    | Start from 1
*  METHODS get_cell_value
*    IMPORTING !column       TYPE i
*              !row          TYPE i
*    RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*
*  METHODS set_array_value
*    IMPORTING !rows TYPE tt_row.
*
*  METHODS set_cell_value
*    IMPORTING !column    TYPE i
*              !row       TYPE i
*              !value     TYPE REF TO lif_xlom__va
*              formula    TYPE REF TO lif_xlom__ex OPTIONAL
*              calculated TYPE abap_bool               OPTIONAL.
*ENDINTERFACE.
*
*
*CLASS lcl_xlom__va DEFINITION.
*  public SECTION.
*    interfaces lif_xlom__ut_all_friends.
*
*    CLASS-METHODS to_boolean
*      IMPORTING !input         TYPE REF TO lif_xlom__va
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_boolean.
*
*    CLASS-METHODS to_number
*      IMPORTING !input        TYPE REF TO lif_xlom__va
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_number
*      RAISING
*        lcx_xlom__va.
*
*    CLASS-METHODS to_string
*      IMPORTING  input        TYPE REF TO lif_xlom__va
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_string
*      RAISING
*        lcx_xlom__va.
*    CLASS-METHODS to_range
*      IMPORTING
*        input         TYPE REF TO lif_xlom__va
*      RETURNING
*        value(result) TYPE REF TO lcl_xlom_range
*      RAISING
*        lcx_xlom__va.
*    CLASS-METHODS to_array
*      IMPORTING
*        input         TYPE REF TO lif_xlom__va
*      RETURNING
*        value(result) TYPE REF TO lif_xlom__va_array
*      RAISING
*        lcx_xlom__va.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut DEFINITION.
*  PUBLIC SECTION.
*    CLASS-METHODS are_equal
*      IMPORTING expression_1  TYPE REF TO lif_xlom__ex
*                expression_2  TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE abap_bool.
*
*    CLASS-METHODS evaluate_array_operands
*      IMPORTING expression    TYPE REF TO lif_xlom__ex
*                !context      TYPE REF TO lcl_xlom__ut_eval_context
*                operands      TYPE lif_xlom__ex=>tt_operand_expr
*      RETURNING VALUE(result) TYPE lif_xlom__ex=>ts_evaluate_array_operands.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_lexer DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    types TY_token_type type string.
*
*    TYPES:
*      BEGIN OF ts_token,
*        value TYPE string,
*        type  TYPE TY_token_type,
*      END OF ts_token.
*    TYPES tt_token TYPE STANDARD TABLE OF ts_token WITH EMPTY KEY.
*
*    TYPES:
*      BEGIN OF ts_parenthesis_group,
*        from_token TYPE i,
*        to_token   TYPE i,
*        level      TYPE i,
*        last_subgroup_token type i,
*      END OF ts_parenthesis_group.
*    TYPES tt_parenthesis_group TYPE STANDARD TABLE OF ts_parenthesis_group WITH EMPTY KEY.
*
*    TYPES:
*      BEGIN OF ts_result_lexe,
*        tokens             TYPE tt_token,
*      END OF ts_result_lexe.
*
*    CONSTANTS:
*      BEGIN OF c_type,
*        comma                      TYPE ty_token_type VALUE ',',
*        comma_space                TYPE ty_token_type VALUE `, `,
*        curly_bracket_close        TYPE ty_token_type VALUE '}',
*        curly_bracket_open         TYPE ty_token_type VALUE '{',
*        "! In RIGHT("hello",) the second argument is empty, interpreted as 0,
*        "! which is different from RIGHT("hello"), where the second argument
*        "! is interpreted as being 1.
*        empty_argument             TYPE ty_token_type VALUE '∅',
*        "! #N/A!, etc.
*        error_name                 TYPE ty_token_type VALUE '#',
*        "! LEN(...), etc.
*        function_name              TYPE ty_token_type VALUE 'F',
*        number                     TYPE ty_token_type VALUE 'N',
*        operator                   TYPE ty_token_type VALUE 'O',
*        parenthesis_close          TYPE ty_token_type VALUE ')',
*        parenthesis_open           TYPE ty_token_type VALUE '(',
*        semicolon                  TYPE ty_token_type VALUE ';',
*        square_bracket_close       TYPE ty_token_type VALUE ']',
*        square_bracket_space_close TYPE ty_token_type VALUE ' ]',
*        square_bracket_open        TYPE ty_token_type VALUE '[',
*        square_brackets_open_close TYPE ty_token_type VALUE '[]',
*        symbol_name                TYPE ty_token_type VALUE 'W',
*        table_name                 TYPE ty_token_type VALUE 'T',
*        text_literal               TYPE ty_token_type VALUE '"',
*      END OF c_type.
*
*    CLASS-METHODS create
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_ut_lexer.
*
*    METHODS lexe IMPORTING !text         TYPE csequence
*                 RETURNING VALUE(result) TYPE tt_token. "ts_result_lexe.
*
*  PRIVATE SECTION.
*    "! Insert the parts of the text in "FIND ... IN text ..." for which there was no match.
*    METHODS complete_with_non_matches
*      IMPORTING i_string  TYPE string
*      CHANGING  c_matches TYPE match_result_tab.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_parser_item DEFINITION FINAL
*  CREATE PRIVATE
*  FRIENDS lcl_xlom__ex_ut_parser.
*
*  PRIVATE SECTION.
*    TYPES tt_item TYPE STANDARD TABLE OF REF TO lcl_xlom__ex_ut_parser_item WITH EMPTY KEY.
*
*    DATA type       TYPE lcl_xlom__ex_ut_lexer=>ts_token-type.
*    DATA value      TYPE lcl_xlom__ex_ut_lexer=>ts_token-value.
*    DATA expression TYPE REF TO lif_xlom__ex.
*    DATA subitems   TYPE tt_item.
*
*    CLASS-METHODS create
*      IMPORTING type       TYPE lcl_xlom__ex_ut_lexer=>ts_token-type
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_ut_parser_item.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_operator DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    TYPES tt_operand_position_offset TYPE STANDARD TABLE OF i WITH EMPTY KEY.
*    TYPES tt_expression              TYPE STANDARD TABLE OF REF TO lif_xlom__ex WITH EMPTY KEY.
*
**    CLASS-DATA multiply TYPE REF TO lcl_xlom__ex_ut_operator READ-ONLY.
**    CLASS-DATA plus TYPE REF TO lcl_xlom__ex_ut_operator READ-ONLY.
*
*    CLASS-METHODS class_constructor.
*
*    CLASS-METHODS create
*      IMPORTING !name                    TYPE string
*                unary                    TYPE abap_bool
*                operand_position_offsets TYPE tt_operand_position_offset
*                !priority                TYPE i
*                description              TYPE csequence
*      RETURNING VALUE(result)            TYPE REF TO lcl_xlom__ex_ut_operator.
*
*    METHODS create_expression
*      IMPORTING operands      TYPE tt_expression
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__ex.
*
*    CLASS-METHODS get
*      IMPORTING operator TYPE string
*                unary    TYPE abap_bool
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_ut_operator.
*
*    "! <ul>
*    "! <li>1 : Reference operators ":" (colon), " " (single space), "," (comma)</li>
*    "! <li>2 : – (as in –1) and + (as in +1)</li>
*    "! <li>3 : % (as in =50%)</li>
*    "! <li>4 : ^ Exponentiation (as in 2^8)</li>
*    "! <li>5 : * and / Multiplication and division                    </li>
*    "! <li>6 : + and – Addition and subtraction                       </li>
*    "! <li>7 : & Connects two strings of text (concatenation)         </li>
*    "! <li>8 : = < > <= >= <> Comparison</li>
*    "! </ul>
*    "!
*    "! @parameter result | .
*    METHODS get_priority
*      RETURNING VALUE(result) TYPE i.
*
*    "! 1 : predecessor operand only (% e.g. 10%)
*    "! 2 : before and after operand only (+ - * / ^ & e.g. 1+1)
*    "! 3 : successor operand only (unary + and - e.g. +5)
*    "!
*    "! @parameter result | .
*    METHODS get_operand_position_offsets
*      RETURNING VALUE(result) TYPE tt_operand_position_offset.
*
*  PRIVATE SECTION.
*    TYPES:
*      "! operator precedence
*      "! Get operator priorities
*      BEGIN OF ts_operator,
*        name                     TYPE string,
*        "! +1 for unary operators (e.g. -1)
*        "! -1 and +1 for binary operators (e.g. 1*2)
*        "! -1 for postfix operators (e.g. 10%)
*        operand_position_offsets TYPE tt_operand_position_offset,
*        "! To distinguish unary from binary operators + and -
*        unary                    TYPE abap_bool,
*        "! <ul>
*        "! <li>1 : Reference operators ":" (colon), " " (single space), "," (comma)</li>
*        "! <li>2 : – (as in –1) and + (as in +1)</li>
*        "! <li>3 : % (as in =50%)</li>
*        "! <li>4 : ^ Exponentiation (as in 2^8)</li>
*        "! <li>5 : * and / Multiplication and division                    </li>
*        "! <li>6 : + and – Addition and subtraction                       </li>
*        "! <li>7 : & Connects two strings of text (concatenation)         </li>
*        "! <li>8 : = < > <= >= <> Comparison</li>
*        "! </ul>
*        priority                 TYPE i,
**        "! % is the only postfix operator e.g. 10% (=0.1)
**        postfix           TYPE abap_bool,
*        desc                     TYPE string,
*        handler                  TYPE REF TO lcl_xlom__ex_ut_operator,
*      END OF ts_operator.
*    TYPES tt_operator TYPE SORTED TABLE OF ts_operator WITH UNIQUE KEY name unary.
*
*    CLASS-DATA operators TYPE tt_operator.
*
*    DATA name                     TYPE string.
*    "! +1 for unary operators (e.g. -1)
*    "! -1 and +1 for binary operators (e.g. 1*2)
*    "! -1 for postfix operators (e.g. 10%)
*    DATA operand_position_offsets TYPE tt_operand_position_offset.
*    "! <ul>
*    "! <li>1 : Reference operators ":" (colon), " " (single space), "," (comma)</li>
*    "! <li>2 : – (as in –1) and + (as in +1)</li>
*    "! <li>3 : % (as in =50%)</li>
*    "! <li>4 : ^ Exponentiation (as in 2^8)</li>
*    "! <li>5 : * and / Multiplication and division                    </li>
*    "! <li>6 : + and – Addition and subtraction                       </li>
*    "! <li>7 : & Connects two strings of text (concatenation)         </li>
*    "! <li>8 : = < > <= >= <> Comparison</li>
*    "! </ul>
*    DATA priority                 TYPE i.
*    "! Unary operators are + and - (like in --A1 or +5)
*    DATA unary                    TYPE abap_bool.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_parser DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    TYPES tt_expr TYPE STANDARD TABLE OF REF TO lif_xlom__ex WITH EMPTY KEY.
*
*    CLASS-METHODS create
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_ut_parser.
*
*    METHODS parse
*      IMPORTING !tokens            TYPE lcl_xlom__ex_ut_lexer=>tt_token
*      RETURNING VALUE(result)      TYPE REF TO lif_xlom__ex
*      RAISING   lcx_xlom__ex_ut_parser.
*
*  PRIVATE SECTION.
*    DATA formula_offset      TYPE i.
*    DATA current_token_index TYPE sytabix.
*    DATA tokens              TYPE lcl_xlom__ex_ut_lexer=>tt_token.
*
*    METHODS get_expression_from_curly_brac
*      IMPORTING arguments     TYPE lcl_xlom__ex_ut_parser_item=>tt_item
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__ex.
*
*    METHODS get_expression_from_error
*      IMPORTING error_name    TYPE string
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__ex.
*
*    METHODS get_expression_from_function
*      IMPORTING function_name TYPE string
*                arguments     TYPE tt_expr
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__ex.
*
*    METHODS get_expression_from_symbol_nam
*      IMPORTING token_value   TYPE string
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__ex.
*
*    "! Transform parentheses and operators into items
*    METHODS parse_expression_item
*      CHANGING item TYPE REF TO lcl_xlom__ex_ut_parser_item.
*
*    "! Merge function item with its next item item (arguments in parentheses)
*    METHODS parse_expression_item_1
*      CHANGING item TYPE REF TO lcl_xlom__ex_ut_parser_item.
*
*    METHODS parse_expression_item_2
*      CHANGING item TYPE REF TO lcl_xlom__ex_ut_parser_item.
*
*    METHODS parse_expression_item_3
*      CHANGING item TYPE REF TO lcl_xlom__ex_ut_parser_item.
*
*    METHODS parse_expression_item_4
*      CHANGING item TYPE REF TO lcl_xlom__ex_ut_parser_item.
*
*    METHODS parse_expression_item_5
*      CHANGING item TYPE REF TO lcl_xlom__ex_ut_parser_item.
*
*    METHODS parse_tokens_until
*      IMPORTING  until TYPE csequence
*      CHANGING  item   TYPE REF TO lcl_xlom__ex_ut_parser_item.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ut_eval_context DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    TYPES:
*      BEGIN OF ts_containing_cell,
*        row    TYPE i,
*        column TYPE i,
*      END OF ts_containing_cell.
*
*    DATA worksheet       TYPE REF TO lcl_xlom_worksheet READ-ONLY.
*    DATA containing_cell TYPE ts_containing_cell           READ-ONLY.
*
*    CLASS-METHODS create
*      IMPORTING worksheet       TYPE REF TO lcl_xlom_worksheet
*                containing_cell TYPE ts_containing_cell
*      RETURNING VALUE(result)   TYPE REF TO lcl_xlom__ut_eval_context.
*
*    METHODS set_containing_cell
*      IMPORTING
*        value type ts_containing_cell.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_array DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex_array.
*
*    TYPES tt_column TYPE STANDARD TABLE OF REF TO lif_xlom__ex WITH EMPTY KEY.
*    TYPES:
*      BEGIN OF ts_row,
*        columns_of_row TYPE tt_column,
*      END OF ts_row.
*    TYPES tt_row TYPE STANDARD TABLE OF ts_row WITH EMPTY KEY.
*
*    CLASS-METHODS create
*      IMPORTING !rows         TYPE tt_row
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_el_array.
*
*  PRIVATE SECTION.
*    DATA rows TYPE tt_row.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_boolean DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING boolean_value TYPE abap_bool
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_el_boolean.
*
*  PRIVATE SECTION.
*    DATA boolean_value TYPE abap_bool.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_empty_arg DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_el_empty_arg.
*
*  PRIVATE SECTION.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_error DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    TYPES ty_error_number TYPE i.
*
*    CLASS-DATA blocked                    TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    CLASS-DATA calc                       TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    CLASS-DATA connect                    TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! #DIV/0! Is produced by =1/0
*    CLASS-DATA division_by_zero           TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    CLASS-DATA field                      TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    CLASS-DATA getting_data               TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! #N/A. Is produced by =ERROR.TYPE(1) or if C1 contains =A1:A2+B1:B3 -> C3=#N/A.
*    CLASS-DATA na_not_applicable          TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! #NAME! Is produced by =XXXX if XXXX is not an existing range name.
*    CLASS-DATA name                       TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    CLASS-DATA null                       TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! #NUM! Is produced by =1E+240*1E+240
*    CLASS-DATA num                        TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! TODO #PYTHON! internal error number is not 2222, what is it?
*    CLASS-DATA python                     TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! #REF! Is produced by =INDEX(A1,2,1)
*    CLASS-DATA ref                        TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! #SPILL! Is produced by A1 containing ={1,2} and B1 containing a value -> A1=#SPILL!
*    CLASS-DATA spill                      TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    CLASS-DATA unknown                    TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*    "! #VALUE! Is produced by =1+"a". #VALUE! in English, #VALEUR! in French.
*    CLASS-DATA value_cannot_be_calculated TYPE REF TO lcl_xlom__ex_el_error READ-ONLY.
*
*    "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
*    DATA english_error_name    TYPE string          READ-ONLY.
*    "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
*    DATA internal_error_number TYPE ty_error_number READ-ONLY.
*
*    CLASS-METHODS class_constructor.
*    CLASS-METHODS get_from_error_name
*      IMPORTING
*        error_name    TYPE csequence
*      RETURNING
*        value(result) TYPE REF TO lif_xlom__ex.
*
*  PRIVATE SECTION.
*    TYPES:
*      BEGIN OF ts_error,
*        "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
*        english_error_name    TYPE string,
*        "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
*        internal_error_number TYPE ty_error_number,
*        object                TYPE REF TO lcl_xlom__ex_el_error,
*      END OF ts_error.
*    TYPES tt_error TYPE STANDARD TABLE OF ts_error WITH EMPTY KEY.
*
*    CLASS-DATA errors TYPE tt_error.
*
*    CLASS-METHODS create
*      IMPORTING english_error_name    TYPE ts_error-english_error_name
*                internal_error_number TYPE ts_error-internal_error_number
*      RETURNING VALUE(result)         TYPE REF TO lcl_xlom__ex_el_error.
*ENDCLASS.
*
*
*"! ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
*"! https://support.microsoft.com/en-us/office/address-function-d0c26c0d-3991-446b-8de4-ab46431d4f89
*CLASS lcl_xlom__ex_fu_address DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING row_num    TYPE REF TO lif_xlom__ex
*                column_num TYPE REF TO lif_xlom__ex
*                ABS_num    TYPE REF TO lif_xlom__ex OPTIONAL
*                a1         TYPE REF TO lif_xlom__ex OPTIONAL
*                Sheet_text TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_address.
*
*  PRIVATE SECTION.
*    DATA row_num    TYPE REF TO lif_xlom__ex.
*    DATA column_num TYPE REF TO lif_xlom__ex.
*    DATA abs_num    TYPE REF TO lif_xlom__ex.
*    DATA a1         TYPE REF TO lif_xlom__ex.
*    DATA sheet_text TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! CELL(info_type, [reference])
*"! CELL("filename",A1) => https://myoffice.accenture.com/personal/xxxxxxxxxxxxxxxxxxxxx/Documents/[activity log.xlsx]Log
*"! In cell B1, formula =CELL("address",A1:A6) is the same result as =CELL("address",A1), which is $A$1 in cell B1;
*"! the cells B2 to B6 are not initialized with $A$2, $A$3, etc.
*"! https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf
*CLASS lcl_xlom__ex_fu_cell DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*    INTERFACES lif_xlom__ex.
*
*    "! @parameter reference | The parameter is technically OPTIONAL but should be passed, as explained at
*    "!                        https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf:
*    "!                        "<em>Important: Although technically reference is <strong>optional</strong>, including it in your formula is encouraged,
*    "!                        unless you understand the effect its absence has on your formula result and want that effect in place.
*    "!                        Omitting the reference argument does not reliably produce information about a specific cell, for the following reasons:</em>"
*    "!                        <ul>
*    "!                        <li>"<em>In automatic calculation mode, when a cell is modified by a user the calculation may be triggered
*    "!                            before or after the selection has progressed, depending on the platform you're using for Excel.
*    "!                            For example, Excel for Windows currently triggers calculation before selection changes, but Excel
*    "!                            for the web triggers it afterward.</em></li>
*    "!                        <li>"<em>When Co-Authoring with another user who makes an edit, this function will report your active cell rather than the editor's.</em>"</li>
*    "!                        <li>"<em>Any recalculation, for instance pressing F9, will cause the function to return a new result even though no cell edit has occurred.</em>"</li>
*    "!                        </ul>
*    CLASS-METHODS create
*      IMPORTING info_type TYPE REF TO lcl_xlom__ex_el_string
*                reference TYPE REF TO lcl_xlom__ex_el_range OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_cell.
*
*  PRIVATE SECTION.
*    TYPES ty_info_type TYPE i.
*
*    CONSTANTS:
*      BEGIN OF c_info_type,
*        filename TYPE string VALUE 'filename',
*      END OF c_info_type.
*
*    DATA info_type TYPE REF TO lcl_xlom__ex_el_string.
*    DATA reference TYPE REF TO lcl_xlom__ex_el_range.
*ENDCLASS.
*
*
*"! COUNTIF(range, criteria)
*"! https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34
*CLASS lcl_xlom__ex_fu_countif DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING range    TYPE REF TO lif_xlom__ex
*                criteria TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_countif.
*
*  PRIVATE SECTION.
*    DATA range    TYPE REF TO lif_xlom__ex.
*    DATA criteria TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! FIND(find_text, within_text, [start_num])
*"! https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628
*CLASS lcl_xlom__ex_fu_find DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING find_text   TYPE REF TO lif_xlom__ex
*                within_text TYPE REF TO lif_xlom__ex
*                start_num   TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_FIND.
*
*  PRIVATE SECTION.
*    DATA find_text   TYPE REF TO lif_xlom__ex.
*    DATA within_text TYPE REF TO lif_xlom__ex.
*    DATA start_num   TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_if DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING condition     TYPE REF TO lif_xlom__ex
*                expr_if_true  TYPE REF TO lif_xlom__ex
*                expr_if_false TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_if.
*
*  PRIVATE SECTION.
*    DATA condition     TYPE REF TO lif_xlom__ex.
*    DATA expr_if_true  TYPE REF TO lif_xlom__ex.
*    DATA expr_if_false TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! IFERROR(value, value_if_error)
*"! IFERROR(#N/A,"1") returns "1"
*"! https://support.microsoft.com/en-us/office/iferror-function-c526fd07-caeb-47b8-8bb6-63f3e417f611
*CLASS lcl_xlom__ex_fu_iferror DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING !value         TYPE REF TO lif_xlom__ex
*                value_if_error TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result)  TYPE REF TO lcl_xlom__ex_fu_iferror.
*
*  PRIVATE SECTION.
*    DATA value          TYPE REF TO lif_xlom__ex.
*    DATA value_if_error TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! INDEX(array, row_num, [column_num])
*"! If row_num is omitted, column_num is required.
*"! If column_num is omitted, row_num is required.
*"! row_num = 0 is interpreted the same way as row_num = 1. Same remark for column_num.
*"! row_num < 0 or column_num < 0 lead to #VALUE!
*"! row_num and column_num with values outside the array lead to #REF!
*"! https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
*CLASS lcl_xlom__ex_fu_index DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING array         TYPE REF TO lif_xlom__ex
*                row_num       TYPE REF TO lif_xlom__ex OPTIONAL
*                column_num    TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_index.
*
*  PRIVATE SECTION.
*    DATA array      TYPE REF TO lif_xlom__ex.
*    DATA row_num    TYPE REF TO lif_xlom__ex.
*    DATA column_num TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! INDIRECT(ref_text, [a1])
*"! https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261
*CLASS lcl_xlom__ex_fu_indirect DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    "!
*    "! @parameter ref_text | Range address
*    "! @parameter a1 | Optional. A logical value that specifies what type of reference is contained in the cell ref_text.
*    "!                 <ul>
*    "!                 <li>If a1 is TRUE or omitted, ref_text is interpreted as an A1-style reference.</li>
*    "!                 <li>If a1 is FALSE, ref_text is interpreted as an R1C1-style reference.</li>
*    "!                 </ul>
*    "! @parameter result | Range
*    CLASS-METHODS create
*      IMPORTING ref_text      TYPE REF TO lif_xlom__ex
*                a1            TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_indirect.
*
*  PRIVATE SECTION.
*    DATA ref_text  TYPE REF TO lif_xlom__ex.
*    DATA a1        TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! LEN(text)
*"! https://support.microsoft.com/en-us/office/len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb
*CLASS lcl_xlom__ex_fu_len DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING !text         TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_len.
*
*  PRIVATE SECTION.
*    DATA text TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! MATCH(lookup_value, lookup_array, [match_type])
*"! https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a
*CLASS lcl_xlom__ex_fu_match DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*    INTERFACES lif_xlom__ex.
*
*    "! MATCH(lookup_value, lookup_array, [match_type])
*    "! https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a
*    "! MATCH returns the position of the matched value within lookup_array, not the value itself. For example, MATCH("b",{"a","b","c"},0) returns 2, which is the relative position of "b" within the array {"a","b","c"}.
*    "! MATCH does not distinguish between uppercase and lowercase letters when matching text values.
*    "! If MATCH is unsuccessful in finding a match, it returns the #N/A error value.
*    "! If match_type is 0 and lookup_value is a text string, you can use the wildcard characters — the question mark (?) and asterisk (*) — in the lookup_value argument.
*    "! A question mark matches any single character; an asterisk matches any sequence of characters.
*    "! If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
*    "! @parameter lookup_value | Required. The value that you want to match in lookup_array.
*    "!                           For example, when you look up someone's number in a telephone book,
*    "!                           you are using the person's name as the lookup value, but the telephone number is the value you want.
*    "!                           The lookup_value argument can be a value (number, text, or logical value)
*    "!                           or a cell reference to a number, text, or logical value.
*    "! @parameter lookup_array | Required. The range of cells being searched.
*    "! @parameter match_type   | Optional. The number -1, 0, or 1. The match_type argument specifies how Excel
*    "!                           matches lookup_value with values in lookup_array. The default value for this argument is 1.
*    "!                           <ul>
*    "!                           <li>1 or omitted: MATCH finds the largest value that is less than or equal to lookup_value.
*    "!                                             The values in the lookup_array argument must be placed in ascending order, for example:
*    "!                                             ...-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE.</li>
*    "!                           <li>0: MATCH finds the first value that is exactly equal to lookup_value.
*    "!                                  The values in the lookup_array argument can be in any order.</li>
*    "!                           <li>-1: MATCH finds the smallest value that is greater than or equal tolookup_value. The values in the lookup_array argument
*    "!                                   must be placed in descending order, for example: TRUE, FALSE, Z-A, ...2, 1, 0, -1, -2, ..., and so on.</li>
*    "!                           </ul>
*    "! @parameter result       | MATCH returns the position of the matched value within lookup_array, not the value itself.
*    "!                           For example, MATCH("b",{"a","b","c"},0) returns 2, which is the relative position of "b" within the array {"a","b","c"}.
*    CLASS-METHODS create
*      IMPORTING lookup_value  TYPE REF TO lif_xlom__ex
*                lookup_array  TYPE REF TO lif_xlom__ex OPTIONAL
*                match_type    TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_match.
*
*  PRIVATE SECTION.
*    DATA lookup_value TYPE REF TO lif_xlom__ex.
*    DATA lookup_array TYPE REF TO lif_xlom__ex.
*    DATA match_type   TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! OFFSET(reference, rows, cols, [height], [width])
*"! OFFSET($A$1,0,0,5,0) is equivalent to $A$1:$A$5
*"! https://support.microsoft.com/en-us/office/offset-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66
*CLASS lcl_xlom__ex_fu_offset DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING !reference    TYPE REF TO lif_xlom__ex
*                !rows         TYPE REF TO lif_xlom__ex
*                cols          TYPE REF TO lif_xlom__ex
*                height        TYPE REF TO lif_xlom__ex OPTIONAL
*                !width        TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_offset.
*
*  PRIVATE SECTION.
*    DATA reference TYPE REF TO lif_xlom__ex.
*    DATA rows      TYPE REF TO lif_xlom__ex.
*    DATA cols      TYPE REF TO lif_xlom__ex.
*    DATA height    TYPE REF TO lif_xlom__ex.
*    DATA width     TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! RIGHT(text,[num_chars])
*"! A1=RIGHT({"hello","world"},{2,3}) -> A1="lo", B1="rld"
*"! https://support.microsoft.com/en-us/office/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f
*CLASS lcl_xlom__ex_fu_right DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING text          TYPE REF TO lif_xlom__ex
*                num_chars     TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_right.
*
*  PRIVATE SECTION.
*    DATA text      TYPE REF TO lif_xlom__ex.
*    DATA num_chars TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! ROW([reference])
*"! https://support.microsoft.com/en-us/office/row-function-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d
*CLASS lcl_xlom__ex_fu_row DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING !reference    TYPE REF TO lif_xlom__ex OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_row.
*
*  PRIVATE SECTION.
*    DATA reference TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! T(value)
*"! If value is or refers to text, T returns value. If value does not refer to text, T returns "" (empty text).
*"! Examples: T("text") = "text", T(1) = "".
*"! https://support.microsoft.com/en-us/office/t-function-fb83aeec-45e7-4924-af95-53e073541228
*CLASS lcl_xlom__ex_fu_t DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING value TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_fu_t.
*
*  PRIVATE SECTION.
*    DATA value TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_number DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING !number       TYPE f
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_el_number.
*
*  PRIVATE SECTION.
*    DATA number TYPE f.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_ampersand DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING left_operand  TYPE REF TO lif_xlom__ex
*                right_operand TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_op_ampersand.
*
*  PRIVATE SECTION.
*    DATA left_operand  TYPE REF TO lif_xlom__ex.
*    DATA right_operand TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*"! Operator colon (e.g. A1:A2, OFFSET(A1,1,1):OFFSET(A1,2,2), my.B1:my.C1 (range names), etc.)
*CLASS lcl_xlom__ex_op_colon DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*    INTERFACES lif_xlom__ex_array.
*
*    CLASS-METHODS create
*      IMPORTING left_operand  TYPE REF TO lif_xlom__ex
*                right_operand TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_op_colon.
*
*  PRIVATE SECTION.
*    DATA left_operand  TYPE REF TO lif_xlom__ex.
*    DATA right_operand TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_equal DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING left_operand  TYPE REF TO lif_xlom__ex
*                right_operand TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_op_equal.
*
*  PRIVATE SECTION.
*    DATA left_operand  TYPE REF TO lif_xlom__ex.
*    DATA right_operand TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_minus DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING left_operand  TYPE REF TO lif_xlom__ex
*                right_operand TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_op_minus.
*
*  PRIVATE SECTION.
*    DATA left_operand  TYPE REF TO lif_xlom__ex.
*    DATA right_operand TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_minus_unry DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING operand       TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_op_minus_unry.
*
*  PRIVATE SECTION.
*    DATA operand TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_mult DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING left_operand  TYPE REF TO lif_xlom__ex
*                right_operand TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_op_mult.
*
*  PRIVATE SECTION.
*    DATA left_operand  TYPE REF TO lif_xlom__ex.
*    DATA right_operand TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_plus DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING left_operand  TYPE REF TO lif_xlom__ex
*                right_operand TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_op_plus.
*
*  PRIVATE SECTION.
*    DATA left_operand  TYPE REF TO lif_xlom__ex.
*    DATA right_operand TYPE REF TO lif_xlom__ex.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_range DEFINITION FINAL
*  CREATE PRIVATE
*  FRIENDS lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*    INTERFACES lif_xlom__ex_array.
*
*    CLASS-METHODS create
*      IMPORTING address_or_name TYPE string
*      RETURNING VALUE(result)   TYPE REF TO lcl_xlom__ex_el_range.
*
*  PRIVATE SECTION.
*    DATA _address_or_name TYPE string.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_string DEFINITION FINAL
*  CREATE PRIVATE
*  FRIENDS lif_xlom__ex.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ex.
*
*    CLASS-METHODS create
*      IMPORTING !text         TYPE csequence
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__ex_el_string.
*
*  PRIVATE SECTION.
*    DATA string TYPE string.
*ENDCLASS.
*
*
*CLASS lcl_xlom_range DEFINITION
*  CREATE PROTECTED
*  FRIENDS lif_xlom__ut_all_friends
*          ltc_evaluate.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*    INTERFACES lif_xlom__va_array.
*
*    DATA application TYPE REF TO lcl_xlom_application READ-ONLY.
*    DATA parent      TYPE REF TO lcl_xlom_worksheet   READ-ONLY.
*
*    "! Address (RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo)
*    "! https://learn.microsoft.com/en-us/office/vba/api/excel.range.address
*    "!
*    "! @parameter row_absolute | True to return the row part of the reference as an absolute reference.
*    "! @parameter column_absolute | True to return the column part of the reference as an absolute reference.
*    "! @parameter reference_style | In A1 or R1C1 format.
*    "! @parameter external | True to return an external reference. False to return a local reference.
*    "! @parameter relative_to | If RowAbsolute and ColumnAbsolute are False, and ReferenceStyle is xlR1C1, you must include a starting point for the relative reference. This argument is a Range object that defines the starting point.
*    "!                          NOTE: Testing with Excel VBA 7.1 shows that an explicit starting point is not mandatory. There appears to be a default reference of $A$1.
*    "! @parameter result | Returns the address of the range, e.g. "A1", "$A$1", etc.
*    METHODS address
*      IMPORTING row_absolute    TYPE abap_bool                       DEFAULT abap_true
*                column_absolute TYPE abap_bool                       DEFAULT abap_true
*                reference_style TYPE lcl_xlom=>ty_reference_style DEFAULT lcl_xlom=>c_reference_style-a1
*                external        TYPE abap_bool                       DEFAULT abap_false
*                relative_to     TYPE REF TO lcl_xlom_range        OPTIONAL
*      RETURNING VALUE(result)   TYPE string.
*
*    METHODS calculate.
*
*    "! Use either both row and column, or item alone.
*    "! @parameter row | Start from 1
*    "! @parameter column | Start from 1.
*    "! @parameter item | Item number from 1, 16385 is the same as row = 2 column = 1.
*    METHODS cells
*      IMPORTING !row          TYPE i    OPTIONAL
*                !column       TYPE i    OPTIONAL
*                item          TYPE int8 OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    METHODS columns
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    METHODS count
*      RETURNING VALUE(result) TYPE i.
*
*    "! Called by the Worksheet.Range property.
*    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
*    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
*    "! @parameter result | .
*    CLASS-METHODS create
*      IMPORTING cell1         TYPE REF TO lcl_xlom_range
*                cell2         TYPE REF TO lcl_xlom_range OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    CLASS-METHODS create_from_address_or_name
*      IMPORTING address       TYPE clike
*                relative_to   TYPE REF TO lcl_xlom_worksheet
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range
*      RAISING
*        lcx_xlom__va.
*
*    "! Range with integer row and column coordinates
*    "! @parameter row | Start from 1
*    "! @parameter column | Start from 1
*    CLASS-METHODS create_from_row_column
*      IMPORTING worksheet     TYPE REF TO lcl_xlom_worksheet
*                !row          TYPE i
*                !column       TYPE i
*                row_size      TYPE i DEFAULT 1
*                column_size   TYPE i DEFAULT 1
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    CLASS-METHODS create_from_expr_range
*      IMPORTING expr_range    TYPE REF TO lcl_xlom__ex_el_range
*                relative_to   TYPE REF TO lcl_xlom_worksheet
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range
*      RAISING
*        lcx_xlom__va.
*
*    METHODS formula2
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__ex.
*
*    "! Offset (RowOffset, ColumnOffset)
*    "! https://learn.microsoft.com/fr-fr/office/vba/api/excel.range.offset
*    "! @parameter row_offset | Start from 0
*    "! @parameter column_offset | Start from 0
*    METHODS offset
*      IMPORTING !row_offset    TYPE i
*                !column_offset TYPE i
*      RETURNING VALUE(result)  TYPE REF TO lcl_xlom_range.
*
*    "! Resize (RowSize, ColumnSize)
*    "! https://learn.microsoft.com/en-us/office/vba/api/excel.range.resize
*    "! @parameter row_size | Start from 1
*    "! @parameter column_size | Start from 1
*    METHODS resize
*      IMPORTING !row_size     TYPE i
*                !column_size  TYPE i
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    METHODS row
*      RETURNING VALUE(result) TYPE i.
*
*    METHODS rows
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    METHODS set_formula2
*      IMPORTING !value TYPE string
*      RAISING   lcx_xlom__ex_ut_parser.
*
*    METHODS set_value
*      IMPORTING !value TYPE REF TO lif_xlom__va.
*
*    METHODS value
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*
*  PRIVATE SECTION.
*    TYPES:
*      BEGIN OF ts_formula_buffer_line,
*        formula TYPE string,
*        object  TYPE REF TO lif_xlom__ex,
*      END OF ts_formula_buffer_line.
*    TYPES tt_formula_buffer TYPE HASHED TABLE OF ts_formula_buffer_line WITH UNIQUE KEY formula.
*
*    "! By default, the "count" method counts the number of cells.
*    "! It's possible to make it count only the columns or rows
*    "! when the range is created by the methods "columns" or "rows".
*    TYPES ty_column_row_collection TYPE i.
*
*    TYPES:
*      BEGIN OF ts_range_buffer_line,
*        worksheet             TYPE REF TO lcl_xlom_worksheet,
*        address               TYPE lcl_xlom=>ts_range_address,
**        address               TYPE lif_xlom_result_array=>ts_address,
*        column_row_collection TYPE ty_column_row_collection,
*        object                TYPE REF TO lcl_xlom_range,
*      END OF ts_range_buffer_line.
*    TYPES tt_range_buffer TYPE HASHED TABLE OF ts_range_buffer_line WITH UNIQUE KEY worksheet address column_row_collection.
*
*    TYPES:
*      BEGIN OF ts_range_name_or_coords,
*        range_name TYPE string,
*        column     TYPE i,
*        row        TYPE i,
*      END OF ts_range_name_or_coords.
*
*    CONSTANTS:
*      "! By default, the "count" method counts the number of cells.
*      "! It's possible to make it count only the columns or rows
*      "! when the range is created by the methods "columns" or "rows".
*      BEGIN OF c_column_row_collection,
*        none    TYPE ty_column_row_collection VALUE 1,
*        columns TYPE ty_column_row_collection VALUE 2,
*        rows    TYPE ty_column_row_collection VALUE 3,
*      END OF c_column_row_collection.
*
*    CLASS-DATA _formula_buffer TYPE tt_formula_buffer.
*    CLASS-DATA _range_buffer   TYPE tt_range_buffer.
*
*    DATA _address TYPE lcl_xlom=>ts_range_address.
**    DATA _address TYPE lif_xlom_result_array=>ts_address.
*
*    CLASS-METHODS convert_column_a_xfd_to_number
*      IMPORTING roman_letters TYPE csequence
*      RETURNING VALUE(result) TYPE i.
*
*    CLASS-METHODS convert_column_number_to_a_xfd
*      IMPORTING !number       TYPE i
*      RETURNING VALUE(result) TYPE string.
*
*    "! @parameter column_row_collection | By default, the "count" method counts the number of cells.
*    "!                                    It's possible to make it count only the columns or rows
*    "!                                    when the range is created by the methods "columns" or "rows".
*    CLASS-METHODS create_from_top_left_bottom_ri
*      IMPORTING worksheet             TYPE REF TO lcl_xlom_worksheet
*                top_left              TYPE lcl_xlom=>ts_range_address-top_left
**                top_left              TYPE lif_xlom_result_array=>ts_address-top_left
*                bottom_right          TYPE lcl_xlom=>ts_range_address-bottom_right
**                bottom_right          TYPE lif_xlom_result_array=>ts_address-bottom_right
*                column_row_collection TYPE ty_column_row_collection DEFAULT c_column_row_collection-none
*      RETURNING VALUE(result)         TYPE REF TO lcl_xlom_range.
*
*    CLASS-METHODS decode_range_address
*      IMPORTING address       TYPE string
**                relative_to   TYPE REF TO lcl_xlom_worksheet
**      RETURNING VALUE(result) TYPE lcl_xlom=>ts_range_address.
*      RETURNING VALUE(result) TYPE lif_xlom__va_array=>ts_address.
*
*    CLASS-METHODS decode_range_address_a1
*      IMPORTING address       TYPE string
*      RETURNING VALUE(result) TYPE lif_xlom__va_array=>ts_address.
*
*    CLASS-METHODS decode_range_coords
*      IMPORTING words         TYPE string_table
*                !from         TYPE i
*                !to           TYPE i
*      RETURNING VALUE(result) TYPE lif_xlom__va_array=>ts_address_one_cell.
*
*    CLASS-METHODS decode_range_name_or_coords
*      IMPORTING range_name_or_coords TYPE string
*      RETURNING VALUE(result)        TYPE lcl_xlom=>ts_range_address_one_cell.
*
*    METHODS _offset_resize
*      IMPORTING row_offset    TYPE i
*                column_offset TYPE i
*                row_size      TYPE i
*                column_size   TYPE i
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    CLASS-METHODS optimize_array_if_range
*      IMPORTING array         TYPE REF TO lif_xlom__va_array
*      RETURNING VALUE(result) TYPE lcl_xlom=>ts_range_address.
*ENDCLASS.
*
*
*CLASS lcl_xlom_columns DEFINITION FINAL
*    INHERITING FROM lcl_xlom_range
*    FRIENDS lcl_xlom_range.
*  PUBLIC SECTION.
*    METHODS count REDEFINITION.
*ENDCLASS.
*
*
*CLASS lcl_xlom_rows DEFINITION FINAL
*    INHERITING FROM lcl_xlom_range
*    FRIENDS lcl_xlom_range.
*  PUBLIC SECTION.
*    METHODS count REDEFINITION.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_array DEFINITION FINAL
*  CREATE PRIVATE
*  FRIENDS lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__va_array.
*    INTERFACES lif_xlom__ut_all_friends.
*
*  TYPES:
*    BEGIN OF ts_used_range_one_cell,
*      column TYPE i,
*      row    TYPE i,
*    END OF ts_used_range_one_cell.
*  TYPES:
*    BEGIN OF ts_used_range,
*      top_left     TYPE ts_used_range_one_cell,
*      bottom_right TYPE ts_used_range_one_cell,
*    END OF ts_used_range.
*
*  DATA used_range TYPE ts_used_range READ-ONLY.
*
*    CLASS-METHODS class_constructor.
*
*    CLASS-METHODS create_from_range
*      IMPORTING !range        TYPE REF TO lcl_xlom_range
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_array.
*
*    CLASS-METHODS create_initial
*      IMPORTING row_count    TYPE i
*                column_count TYPE i
*                rows type lif_xlom__va_array=>tt_row optional
*      RETURNING VALUE(result)     TYPE REF TO lcl_xlom__va_array.
*
*  PRIVATE SECTION.
*    "! Internal type of cell value (empty, number, string, boolean, error, array, compound data)
*    TYPES ty_value_type TYPE i.
*    TYPES:
*      BEGIN OF ts_cell,
*        "! Start from 1
*        column     TYPE i,
*        "! Start from 1
*        row        TYPE i,
*        formula    TYPE REF TO lif_xlom__ex,
*        calculated TYPE abap_bool,
*        value      TYPE REF TO lif_xlom__va,
**        "! Type of cell value, among empty, number, text, boolean, error, compound data. For NUMBER, BOOLEAN and ERROR, the value is defined by VALUE2-DOUBLE.
**        "! For TEXT, the value is defined by VALUE2-STRING.
**        value_type TYPE ty_value_type,
**        "! Number, Error, Boolean: <ul>
**        "! <li>If TYPE = C_TYPE-BOOLEAN, the possible values are the constants C_BOOLEAN-TRUE (-1) and C_BOOLEAN-FALSE (0).</li>
**        "! <li>If TYPE = C_TYPE-ERROR, the possible values are the internal numbers defined in lcl_xlom__va_ERROR</li>
**        "! </ul>
**        double     TYPE f,
**        string     TYPE string,
*      END OF ts_cell.
*    TYPES tt_cell TYPE SORTED TABLE OF ts_cell WITH UNIQUE KEY row column.
*
*    CONSTANTS:
*      "! Type of cell value needed by Excel Object Model
*      BEGIN OF c_value_type,
*        "! Needed by ISBLANK formula function. IT CANNOT be replaced with "empty = xsdbool( not line_exists( _cells[ row = ... column = ... ] ) )"
*        "! because a cell may exist for data other than value, like number format, background color, and so on. Optionally, there could be two
*        "! internal tables, _cells only for values.
*        empty         TYPE ty_value_type VALUE 1,
*        number        TYPE ty_value_type VALUE 2,
*        string        TYPE ty_value_type VALUE 3,
*        "! Cell containing the value TRUE or FALSE.
*        "! Needed by TYPE formula function (4 = logical value)
*        boolean       TYPE ty_value_type VALUE 4,
*        "! Needed by TYPE formula function (16 = error)
*        error         TYPE ty_value_type VALUE 5,
*        "! Needed by TYPE formula function (64 = array)
*        array         TYPE ty_value_type VALUE 6,
*        "! Needed by TYPE formula function (128 = compound data)
*        compound_data TYPE ty_value_type VALUE 7,
*      END OF c_value_type.
*
*    CLASS-DATA initial_used_range TYPE lcl_xlom__va_array=>ts_used_range.
*
*    DATA _cells TYPE tt_cell.
*
*    "!
*    "! @parameter row | Start from 1
*    "! @parameter column | Start from 1
*    METHODS set_cell_value_single
*      IMPORTING !row       TYPE i
*                !column    TYPE i
*                !value     TYPE REF TO lif_xlom__va
*                formula    TYPE REF TO lif_xlom__ex OPTIONAL
*                calculated TYPE abap_bool               OPTIONAL.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_boolean DEFINITION FINAL
*  CREATE PRIVATE
*  FRIENDS lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__va.
*
*    CLASS-DATA false TYPE REF TO lcl_xlom__va_boolean READ-ONLY.
*    CLASS-DATA true  TYPE REF TO lcl_xlom__va_boolean READ-ONLY.
*
*    DATA boolean_value TYPE abap_bool READ-ONLY.
*
*    CLASS-METHODS class_constructor.
*
*    CLASS-METHODS get
*      IMPORTING boolean_value TYPE abap_bool
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_boolean.
*
*  PRIVATE SECTION.
*    DATA number        TYPE f.
*
*    CLASS-METHODS create
*      IMPORTING boolean_value TYPE abap_bool
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_boolean.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_empty DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__va.
*
*    CLASS-METHODS get_singleton
*        RETURNING VALUE(result) type ref to lcl_xlom__va_empty.
*  PRIVATE SECTION.
*    CLASS-DATA singleton type ref to lcl_xlom__va_empty.
*ENDCLASS.
*
*
*"! https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/cell-error-values
*"! NB: many errors are missing, the list of the other errors can be found in xlCVError enumeration.
*"! #VALUE! in English, #VALEUR! in French, etc.
*"!
*"! You can insert a cell error value into a cell or test the value of a cell for an error value by
*"! using the CVErr function. The cell error values can be one of the following xlCVError constants.
*"! <ul>
*"! <li>Constant . .Error number . .Cell error value</li>
*"! <li>xlErrDiv0 . 2007 . . . . . .#DIV/0!         </li>
*"! <li>xlErrNA . . 2042 . . . . . .#N/A            </li>
*"! <li>xlErrName . 2029 . . . . . .#NAME?          </li>
*"! <li>xlErrNull . 2000 . . . . . .#NULL!          </li>
*"! <li>xlErrNum . .2036 . . . . . .#NUM!           </li>
*"! <li>xlErrRef . .2023 . . . . . .#REF!           </li>
*"! <li>xlErrValue .2015 . . . . . .#VALUE!         </li>
*"! </ul>
*"! VB example:
*"! <ul>
*"! <li>If IsError(ActiveCell.Value) Then            </li>
*"! <li>. If ActiveCell.Value = CVErr(xlErrDiv0) Then</li>
*"! <li>. End If                                     </li>
*"! <li>End If                                       </li>
*"! </ul>
*"! NB:
*"! <ul>
*"! <li>CVErr(xlErrDiv0) is of type Variant/Error and Locals/Watches shows: Error 2007</li>
*"! <li>There is no Error data type, only Variant can be used.                        </li>
*"! </ul>
*CLASS lcl_xlom__va_error DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__va.
*
*    TYPES ty_error_number TYPE i.
*
*    CLASS-DATA blocked                    TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    CLASS-DATA calc                       TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    CLASS-DATA connect                    TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! #DIV/0! Is produced by =1/0
*    CLASS-DATA division_by_zero           TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    CLASS-DATA field                      TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    CLASS-DATA getting_data               TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! #N/A. Is produced by =ERROR.TYPE(1) or if C1 contains =A1:A2+B1:B3 -> C3=#N/A.
*    CLASS-DATA na_not_applicable          TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! #NAME! Is produced by =XXXX if XXXX is not an existing range name.
*    CLASS-DATA name                       TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    CLASS-DATA null                       TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! #NUM! Is produced by =1E+240*1E+240
*    CLASS-DATA num                        TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! TODO #PYTHON! internal error number is not 2222, what is it?
*    CLASS-DATA python                     TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! #REF! Is produced by =INDEX(A1,2,1)
*    CLASS-DATA ref                        TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! #SPILL! Is produced by A1 containing ={1,2} and B1 containing a value -> A1=#SPILL!
*    CLASS-DATA spill                      TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    CLASS-DATA unknown                    TYPE REF TO lcl_xlom__va_error READ-ONLY.
*    "! #VALUE! Is produced by =1+"a". #VALUE! in English, #VALEUR! in French.
*    CLASS-DATA value_cannot_be_calculated TYPE REF TO lcl_xlom__va_error READ-ONLY.
*
*    "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
*    DATA english_error_name    TYPE string          READ-ONLY.
*    "! Example how the error is obtained
*    DATA description           TYPE string          READ-ONLY.
*    "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
*    DATA internal_error_number TYPE ty_error_number READ-ONLY.
*    "! Result of formula function ERROR.TYPE e.g. 3 for =ERROR.TYPE(#VALUE!)
*    DATA formula_error_number  TYPE ty_error_number READ-ONLY.
*
*    CLASS-METHODS class_constructor.
*
*    CLASS-METHODS get_by_error_number
*      IMPORTING !type         TYPE ty_error_number
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_error.
*
*  PRIVATE SECTION.
*    TYPES:
*      BEGIN OF ts_error,
*        "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
*        english_error_name    TYPE string,
*        "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
*        internal_error_number TYPE ty_error_number,
*        "! Result of formula function ERROR.TYPE e.g. 3 for =ERROR.TYPE(#VALUE!)
*        formula_error_number  TYPE ty_error_number,
*        object                TYPE REF TO lcl_xlom__va_error,
*      END OF ts_error.
*    TYPES tt_error TYPE STANDARD TABLE OF ts_error WITH EMPTY KEY.
*
*    CLASS-DATA errors TYPE tt_error.
*
*    "!
*    "! @parameter english_error_name | English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
*    "! @parameter internal_error_number | Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
*    "! @parameter formula_error_number | Result of formula function ERROR.TYPE e.g. 3 for =ERROR.TYPE(#VALUE!)
*    "! @parameter description | Example how the error is obtained
*    CLASS-METHODS create
*      IMPORTING english_error_name    TYPE string
*                internal_error_number TYPE ty_error_number
*                formula_error_number  TYPE ty_error_number
*                !description          TYPE string OPTIONAL
*      RETURNING VALUE(result)         TYPE REF TO lcl_xlom__va_error.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_number DEFINITION FINAL
*  CREATE PRIVATE.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__va.
*
*    CLASS-METHODS create
*      IMPORTING !number       TYPE f
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_number.
*
*    CLASS-METHODS get
*      IMPORTING !number       TYPE f
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_number.
*
*    METHODS get_integer
*      RETURNING VALUE(result) TYPE i.
*
*    METHODS get_number
*      RETURNING VALUE(result) TYPE f.
*
*  PRIVATE SECTION.
*    TYPES:
*      BEGIN OF ts_buffer_line,
*        number TYPE f,
*        object TYPE REF TO lcl_xlom__va_number,
*      END OF ts_buffer_line.
*    TYPES tt_buffer TYPE SORTED TABLE OF ts_buffer_line WITH UNIQUE KEY number.
*
*    CLASS-DATA buffer TYPE tt_buffer.
*
*    DATA number TYPE f.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_string DEFINITION FINAL
*  CREATE PRIVATE
*  friends lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__va.
*
*    CLASS-METHODS create
*      IMPORTING !string       TYPE csequence
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_string.
*
*    CLASS-METHODS get
*      IMPORTING !string       TYPE csequence
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom__va_string.
*
*    METHODS get_string
*      RETURNING VALUE(result) TYPE string.
*
*  PRIVATE SECTION.
*    TYPES:
*      BEGIN OF ts_buffer_line,
*        string TYPE string,
*        object TYPE REF TO lcl_xlom__va_string,
*      END OF ts_buffer_line.
*    TYPES tt_buffer TYPE HASHED TABLE OF ts_buffer_line WITH UNIQUE KEY string.
*
*    CLASS-DATA buffer TYPE tt_buffer.
*
*    DATA string TYPE string.
*ENDCLASS.
*
*
*CLASS lcl_xlom_sheet DEFINITION.
*ENDCLASS.
*
*
*CLASS lcl_xlom_worksheet DEFINITION FINAL
*  CREATE PRIVATE
*  INHERITING FROM lcl_xlom_sheet
*  FRIENDS lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*    TYPES ty_name TYPE c LENGTH 31.
*
*    DATA application TYPE REF TO lcl_xlom_application READ-ONLY.
*    "! worksheet name TODO
*    DATA name        TYPE string                         READ-ONLY.
*    DATA parent      TYPE REF TO lcl_xlom_workbook    READ-ONLY.
*
*    "! Worksheet.Calculate method (Excel).
*    "! Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.
*    "! <p>expression.Calculate</p>
*    "! expression A variable that represents a Worksheet object.
*    "! https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.calculate(method)
*    METHODS calculate.
*
*    "! Use either both row and column, or item alone.
*    "! @parameter row | Start from 1
*    "! @parameter column | Start from 1.
*    "! @parameter item | Item number from 1, 16385 is the same as row = 2 column = 1.
*    METHODS cells
*      IMPORTING !row          TYPE i    OPTIONAL
*                !column       TYPE i    OPTIONAL
*                item          TYPE int8 OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*    CLASS-METHODS create
*      IMPORTING workbook      TYPE REF TO lcl_xlom_workbook
*                !name         TYPE csequence
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_worksheet.
*
**    METHODS decode_range_address
**      IMPORTING address       TYPE string
**      RETURNING VALUE(result) TYPE lif_xlom_result_array=>ts_address.
*
*    METHODS range
*      IMPORTING cell1_string  TYPE string OPTIONAL
*                cell2_string  TYPE string OPTIONAL
*                cell1_range   TYPE REF TO lcl_xlom_range OPTIONAL
*                cell2_range   TYPE REF TO lcl_xlom_range OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range
*      RAISING
*        lcx_xlom__va.
*
*    METHODS used_range
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*
*  PRIVATE SECTION.
*    CONSTANTS max_rows    TYPE i VALUE 1048576.
*    CONSTANTS max_columns TYPE i VALUE 16384.
*
*    DATA _array TYPE REF TO lcl_xlom__va_array.
*
*    "! Worksheet.Range property. Returns a Range object that represents a cell or a range of cells.
*    "! <p>expression.Range (Cell1, Cell2)</p>
*    "! expression A variable that represents a Worksheet object.
*    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
*    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
*    "! @parameter result | .
*    METHODS range_from_address
*      IMPORTING cell1         TYPE string
*                cell2         TYPE string OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range
*      RAISING
*        lcx_xlom__va.
*
*    "! Worksheet.Range property. Returns a Range object that represents a cell or a range of cells.
*    "! <p>expression.Range (Cell1, Cell2)</p>
*    "! expression A variable that represents a Worksheet object.
*    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
*    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
*    "! @parameter result | .
*    METHODS range_from_two_ranges
*      IMPORTING cell1         TYPE REF TO lcl_xlom_range
*                cell2         TYPE REF TO lcl_xlom_range
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_range.
*ENDCLASS.
*
*
*CLASS lcl_xlom_worksheets DEFINITION FRIENDS lif_xlom__ut_all_friends.
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*    DATA application TYPE REF TO lcl_xlom_application READ-ONLY.
*    DATA count       TYPE i                              READ-ONLY.
*    DATA workbook    TYPE REF TO lcl_xlom_workbook    READ-ONLY.
*
*    METHODS add
*      IMPORTING !name         TYPE lcl_xlom_worksheet=>ty_name
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_worksheet.
*
*    CLASS-METHODS create
*      IMPORTING workbook      TYPE REF TO lcl_xlom_workbook
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_worksheets.
*
*    "! @parameter index  | Required    Variant The name or index number of the object.
*    METHODS item
*      IMPORTING !index        TYPE simple
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_worksheet
*      RAISING
*        lcx_xlom__va.
*
*  PRIVATE SECTION.
*    TYPES:
*      BEGIN OF ty_worksheet,
*        name   TYPE lcl_xlom_worksheet=>ty_name,
*        object TYPE REF TO lcl_xlom_worksheet,
*      END OF ty_worksheet.
*    TYPES ty_worksheets TYPE SORTED TABLE OF ty_worksheet WITH UNIQUE KEY name.
*
*    DATA worksheets TYPE ty_worksheets.
*ENDCLASS.
*
*
*CLASS lcl_xlom_workbook DEFINITION
*  FRIENDS lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*    TYPES ty_name TYPE string.
*
*    DATA application TYPE REF TO lcl_xlom_application READ-ONLY.
*    "! workbook name TODO
*    DATA name        TYPE string                         READ-ONLY.
*    "! workbook path TODO
*    DATA path        TYPE string                         READ-ONLY.
*    DATA worksheets  TYPE REF TO lcl_xlom_worksheets  READ-ONLY.
*
*    CLASS-METHODS create
*      IMPORTING !application  TYPE REF TO lcl_xlom_application
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_workbook.
*
*    "! SaveAs (FileName, FileFormat, Password, WriteResPassword,
*    "!         ReadOnlyRecommended, CreateBackup, AccessMode,
*    "!         ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)
*    "! https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
*    "!
*    "! @parameter file_name | A string that indicates the name of the file to be saved. You can include
*    "!                        a full path; if you don't, Microsoft Excel saves the file in the current folder.
*    METHODS save_as
*      IMPORTING file_name TYPE csequence.
*
*  PRIVATE SECTION.
*    EVENTS saved.
*ENDCLASS.
*
*
*CLASS lcl_xlom_application DEFINITION
*  FRIENDS lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*    DATA active_sheet    TYPE REF TO lcl_xlom_sheet READ-ONLY.
*    DATA calculation     TYPE lcl_xlom=>ty_calculation VALUE lcl_xlom=>c_calculation-automatic READ-ONLY.
*    DATA reference_style TYPE lcl_xlom=>ty_reference_style VALUE lcl_xlom=>c_reference_style-a1 READ-ONLY.
*    DATA workbooks       TYPE REF TO lcl_xlom_workbooks READ-ONLY.
*
*    METHODS calculate.
*
*    CLASS-METHODS create
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_application.
*
*    METHODS evaluate
*      IMPORTING !name         TYPE csequence
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*
*    METHODS international
*      IMPORTING !index        TYPE lcl_xlom=>ty_application_international
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*
*    METHODS intersect
*      IMPORTING arg1          TYPE REF TO lcl_xlom_range
*                arg2          TYPE REF TO lcl_xlom_range
*                arg3          TYPE REF TO lcl_xlom_range OPTIONAL
*                arg4          TYPE REF TO lcl_xlom_range OPTIONAL
*                arg5          TYPE REF TO lcl_xlom_range OPTIONAL
*                arg6          TYPE REF TO lcl_xlom_range OPTIONAL
*                arg7          TYPE REF TO lcl_xlom_range OPTIONAL
*                arg8          TYPE REF TO lcl_xlom_range OPTIONAL
*                arg9          TYPE REF TO lcl_xlom_range OPTIONAL
*                arg10         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg11         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg12         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg13         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg14         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg15         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg16         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg17         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg18         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg19         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg20         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg21         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg22         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg23         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg24         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg25         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg26         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg27         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg28         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg29         TYPE REF TO lcl_xlom_range OPTIONAL
*                arg30         TYPE REF TO lcl_xlom_range OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lif_xlom__va.
*
*    METHODS set_calculation
*      IMPORTING value TYPE lcl_xlom=>ty_calculation DEFAULT lcl_xlom=>c_calculation-automatic.
*
*  PRIVATE SECTION.
*    " 1 = US
*    DATA _country_code TYPE i VALUE 1.
*
*    CLASS-METHODS _intersect_2
*      IMPORTING arg1          TYPE REF TO lcl_xlom_range
*                arg2          TYPE REF TO lcl_xlom_range
*      RETURNING VALUE(result) TYPE lcl_xlom=>ts_range_address.
*
*    CLASS-METHODS _intersect_2_basis
*      IMPORTING arg1          TYPE lcl_xlom=>ts_range_address
*                arg2          TYPE lcl_xlom=>ts_range_address
*      RETURNING VALUE(result) TYPE lcl_xlom=>ts_range_address.
*
*    CLASS-METHODS type
*      IMPORTING any_data_object TYPE any
*      RETURNING VALUE(result)   TYPE abap_typekind.
*ENDCLASS.
*
*
*CLASS lcl_xlom_workbooks DEFINITION
*  FRIENDS lif_xlom__ut_all_friends.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*    DATA application TYPE REF TO lcl_xlom_application READ-ONLY.
*    DATA count       TYPE i                              READ-ONLY.
*
*    CLASS-METHODS create
*      IMPORTING application TYPE REF TO lcl_xlom_application
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_workbooks.
*
*    "! Add (Template)
*    "! https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.add
*    METHODS add
*      IMPORTING template      TYPE any OPTIONAL
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_workbook.
*
*    "!
*    "! @parameter index  | Required    Variant The name or index number of the object.
*    "! @parameter result | .
*    METHODS item
*      IMPORTING !index        TYPE simple
*      RETURNING VALUE(result) TYPE REF TO lcl_xlom_workbook.
*
*  PRIVATE SECTION.
*    TYPES:
*      BEGIN OF ty_workbook,
*        name   TYPE lcl_xlom_workbook=>ty_name,
*        object TYPE REF TO lcl_xlom_workbook,
*      END OF ty_workbook.
*    TYPES ty_workbooks TYPE SORTED TABLE OF ty_workbook WITH NON-UNIQUE KEY name
*                        WITH UNIQUE SORTED KEY by_object COMPONENTS object.
*
*    DATA workbooks TYPE ty_workbooks.
*
*    METHODS on_saved FOR EVENT saved OF lcl_xlom_workbook
*      IMPORTING sender.
*ENDCLASS.
*
*
*
*
*CLASS lcl_xlom IMPLEMENTATION.
*ENDCLASS.
*
*
*CLASS lcl_xlom_application IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom_application( ).
*    result->workbooks = lcl_xlom_workbooks=>create( result ).
*  ENDMETHOD.
*
*  METHOD calculate.
*    DATA(workbook_number) = 1.
*    WHILE workbook_number <= workbooks->count.
*      DATA(workbook) = workbooks->item( workbook_number ).
*
*      DATA(worksheet_number) = 1.
*      WHILE worksheet_number <= workbook->worksheets->count.
*        TRY.
*            DATA(worksheet) = workbook->worksheets->item( worksheet_number ).
*          CATCH lcx_xlom__va INTO DATA(error). " TODO: variable is assigned but never used (ABAP cleaner)
*            RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*        ENDTRY.
*        worksheet->calculate( ).
*        worksheet_number = worksheet_number + 1.
*      ENDWHILE.
*
*      workbook_number = workbook_number + 1.
*    ENDWHILE.
*  ENDMETHOD.
*
*  METHOD evaluate.
*    DATA(lexer) = lcl_xlom__ex_ut_lexer=>create( ).
*    DATA(lexer_tokens) = lexer->lexe( name ).
*    DATA(parser) = lcl_xlom__ex_ut_parser=>create( ).
*
*    TRY.
*    DATA(expression) = parser->parse( lexer_tokens ).
*    CATCH lcx_xlom__ex_ut_parser.
*      result = lcl_xlom__va_error=>value_cannot_be_calculated.
*      RETURN.
*    ENDTRY.
*
*    result = expression->evaluate( context = lcl_xlom__ut_eval_context=>create( worksheet       = CAST #( active_sheet )
*                                                                                     containing_cell = VALUE #( row = 1 column = 1 ) ) ).
*  ENDMETHOD.
*
*  METHOD international.
*    CASE index.
*      WHEN lcl_xlom=>c_application_international-country_code.
*        result = lcl_xlom__va_number=>get( EXACT #( _country_code ) ).
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD intersect.
*    TYPES tt_range TYPE STANDARD TABLE OF REF TO lcl_xlom_range WITH EMPTY KEY.
*
*    DATA(args) = VALUE tt_range( ( arg1 ) ( arg2 ) ( arg3 ) ( arg4 ) ( arg5 ) ( arg6 ) ( arg7 ) ( arg8 ) ( arg9 ) ( arg10 )
*                                 ( arg11 ) ( arg12 ) ( arg13 ) ( arg14 ) ( arg15 ) ( arg16 ) ( arg17 ) ( arg18 ) ( arg19 ) ( arg20 )
*                                 ( arg21 ) ( arg22 ) ( arg23 ) ( arg24 ) ( arg25 ) ( arg26 ) ( arg27 ) ( arg28 ) ( arg29 ) ( arg30 ) ).
**    DATA(args) = VALUE tt_range( ( arg1 ) ( arg2 ) ( arg3 ) ( arg4 ) ( arg5 ) ( arg6 ) ( arg7 ) ( arg8 ) ( arg9 ) ( arg10 )
**                                 ( arg11 ) ( arg12 ) ( arg13 ) ( arg14 ) ( arg15 ) ( arg16 ) ( arg17 ) ( arg18 ) ( arg19 ) ( arg20 )
**                                 ( arg21 ) ( arg22 ) ( arg23 ) ( arg24 ) ( arg25 ) ( arg26 ) ( arg27 ) ( arg28 ) ( arg29 ) ( arg30 ) ).
*
*    DATA(temp_intersect_range_address) = value lcl_xlom=>ts_range_address( ).
*    LOOP AT args INTO DATA(arg)
*        WHERE table_line IS BOUND.
*      temp_intersect_range_address = _intersect_2_basis( arg1 = temp_intersect_range_address
*                                         arg2 = VALUE #( top_left-column     = arg->_address-top_left-column
*                                                   top_left-row        = arg->_address-top_left-row
*                                                   bottom_right-column = arg->_address-bottom_right-column
*                                                   bottom_right-row    = arg->_address-bottom_right-row ) ).
*      IF temp_intersect_range_address IS INITIAL.
*        " Empty intersection
*        RETURN.
*      ENDIF.
*    ENDLOOP.
*
*    result = lcl_xlom_range=>create_from_top_left_bottom_ri( worksheet    = arg1->parent
*                                                                top_left     = VALUE #( column = temp_intersect_range_address-top_left-column
*                                                                                        row    = temp_intersect_range_address-top_left-row )
*                                                                bottom_right = VALUE #( column = temp_intersect_range_address-bottom_right-column
*                                                                                        row    = temp_intersect_range_address-bottom_right-row ) ).
*  ENDMETHOD.
*
*  METHOD _intersect_2.
*    TYPES tt_range TYPE STANDARD TABLE OF REF TO lcl_xlom_range WITH EMPTY KEY.
*
*    DATA(args) = VALUE tt_range( ( arg1 ) ( arg2 ) ).
*
*    LOOP AT args INTO DATA(arg)
*        WHERE table_line IS BOUND.
*      result = _intersect_2_basis( arg1 = result
*                                   arg2 = VALUE #( top_left-column     = arg->_address-top_left-column
*                                                   top_left-row        = arg->_address-top_left-row
*                                                   bottom_right-column = arg->_address-bottom_right-column
*                                                   bottom_right-row    = arg->_address-bottom_right-row ) ).
*      IF result IS INITIAL.
*        " Empty intersection
*        RETURN.
*      ENDIF.
*    ENDLOOP.
*  ENDMETHOD.
*
*  METHOD _intersect_2_basis.
*    result = COND #( WHEN arg1 IS NOT INITIAL
*                     THEN arg1
*                     ELSE VALUE #( top_left-column     = 0
*                                   top_left-row        = 0
*                                   bottom_right-column = lcl_xlom_worksheet=>max_columns + 1
*                                   bottom_right-row    = lcl_xlom_worksheet=>max_rows + 1 ) ).
*
*    IF arg2-top_left-column > result-top_left-column.
*      result-top_left-column = arg2-top_left-column.
*    ENDIF.
*    IF arg2-top_left-row > result-top_left-row.
*      result-top_left-row = arg2-top_left-row.
*    ENDIF.
*    IF arg2-bottom_right-column < result-bottom_right-column.
*      result-bottom_right-column = arg2-bottom_right-column.
*    ENDIF.
*    IF arg2-bottom_right-row < result-bottom_right-row.
*      result-bottom_right-row = arg2-bottom_right-row.
*    ENDIF.
*
*    IF    result-top_left-column > result-bottom_right-column
*       OR result-top_left-row    > result-bottom_right-row.
*      " Empty intersection
*      result = VALUE #( ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD set_calculation.
*    calculation = value.
*  ENDMETHOD.
*
*  METHOD type.
*    DESCRIBE FIELD any_data_object TYPE result.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom_columns IMPLEMENTATION.
*  METHOD count.
*    result = lif_xlom__va_array~column_count.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ut_eval_context IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ut_eval_context( ).
*    result->worksheet       = worksheet.
*    result->containing_cell = containing_cell.
*  ENDMETHOD.
*
*  METHOD set_containing_cell.
*    me->containing_cell = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut IMPLEMENTATION.
*  METHOD are_equal.
*    result = xsdbool(    (     expression_1 IS NOT BOUND
*                           AND expression_2 IS NOT BOUND )
*                      OR (     expression_1 IS BOUND
*                           AND expression_2 IS BOUND
*                           AND expression_1->is_equal( expression_2 ) ) ).
*  ENDMETHOD.
*
*  METHOD evaluate_array_operands.
*    "  formula with operation on arrays        result
*    "
*    "  -> if the left operand is one line high, the line is replicated till max lines of the right operand.
*    "  -> if the left operand is one column wide, the column is replicated till max columns of the right operand.
*    "  -> if the right operand is one line high, the line is replicated till max lines of the left operand.
*    "  -> if the right operand is one column wide, the column is replicated till max columns of the left operand.
*    "
*    "  -> if the left operand has less lines than the right operand, additional lines are added with #N/A.
*    "  -> if the left operand has less columns than the right operand, additional columns are added with #N/A.
*    "  -> if the right operand has less lines than the left operand, additional lines are added with #N/A.
*    "  -> if the right operand has less columns than the left operand, additional columns are added with #N/A.
*    "
*    "  -> target array size = max lines of both operands + max columns of both operands.
*    "  -> each target cell of the target array is calculated like this:
*    "     T(1,1) = L(1,1) op R(1,1)
*    "     T(2,1) = L(2,1) op R(2,1)
*    "     etc.
*    "     If the left cell or right cell is #N/A, the target cell is also #N/A.
*    "
*    "  Examples where one of the two operands has 1 cell, 1 line or 1 column
*    "
*    "  a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
*    "
*    "  a | b | c   op   k                      a op k | b op k | c op k
*    "  d | e | f                               d op k | e op k | f op k
*    "  g | h | i                               g op k | h op k | i op k
*    "
*    "  a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
*    "  d | e | f                               d op k | e op l | f op m | #N/A
*    "  g | h | i                               g op k | h op l | i op m | #N/A
*    "
*    "  a | b | c   op   k                      a op k | b op k | c op k
*    "  d | e | f        l                      d op l | e op l | f op l
*    "  g | h | i        m                      g op m | h op m | i op m
*    "                   n                      #N/A   | #N/A   | #N/A
*    "
*    "  a | b | c   op   k                      a op k | b op k | c op k
*    "  d | e | f        l                      d op l | e op l | f op l
*    "  g | h | i                               #N/A   | #N/A   | #N/A
*    "
*    "  a | b | c   op   k                      a op k | b op k | c op k
*    "                   l                      a op l | b op l | c op l
*    "                   m                      a op m | b op m | c op m
*    "
*    "  Both operands have more than 1 line and more than 1 column
*    "
*    "  a | b | c   op   k | n                  a op k | b op n | #N/A
*    "  d | e | f        l | o                  d op l | e op o | #N/A
*    "  g | h | i                               #N/A   | #N/A   | #N/A
*    "
*    "  a | b | c   op   k | n                  a op k | b op n | #N/A
*    "  d | e | f        l | o                  d op l | e op o | #N/A
*    "                   m | p                  #N/A   | #N/A   | #N/A
*    "
*    "  a | b       op   k | n | q              a op k | b op n | #N/A
*    "  d | e            l | o | r              d op l | e op o | #N/A
*    "  g | h                                   #N/A   | #N/A   | #N/A
*    DATA(at_least_one_array_or_range) = abap_false.
*    LOOP AT operands REFERENCE INTO DATA(operand).
*      IF operand->object IS NOT BOUND.
*        " e.g. NUM_CHARS not passed to function RIGHT
*        INSERT VALUE #( name                     = operand->name
*                        object                   = VALUE #( )
*                        not_part_of_result_array = operand->not_part_of_result_array )
*               INTO TABLE result-operand_results.
*      ELSE.
*        " Evaluate the operand
*        INSERT VALUE #( name                     = operand->name
*                        object                   = operand->object->evaluate( context )
*                        not_part_of_result_array = operand->not_part_of_result_array )
*               INTO TABLE result-operand_results
*               REFERENCE INTO DATA(operand_result).
*        IF     operand_result->not_part_of_result_array = abap_false
*           AND (    operand_result->object->type = operand_result->object->c_type-array
*                 OR operand_result->object->type = operand_result->object->c_type-range )
*           AND (    CAST lif_xlom__va_array( operand_result->object )->row_count    > 1
*                 OR CAST lif_xlom__va_array( operand_result->object )->column_count > 1 ).
*          " Perform array evaluation on more than 1 cell
*          at_least_one_array_or_range = abap_true.
*        ENDIF.
*      ENDIF.
*    ENDLOOP.
*
*    " Continue only if there's at least one array or one range with more than one cell.
*    IF at_least_one_array_or_range = abap_false.
*      RETURN.
*    ENDIF.
*
*    DATA(max_row_count) = 1.
*    DATA(max_column_count) = 1.
*    LOOP AT result-operand_results REFERENCE INTO operand_result
*         WHERE     not_part_of_result_array  = abap_false
*               AND object                   IS BOUND.
*      CASE operand_result->object->type.
*        WHEN operand_result->object->c_type-array
*          OR operand_result->object->c_type-range.
*          max_row_count = nmax( val1 = max_row_count
*                                val2 = CAST lif_xlom__va_array( operand_result->object )->row_count ).
*          max_column_count = nmax( val1 = max_column_count
*                                   val2 = CAST lif_xlom__va_array( operand_result->object )->column_count ).
*      ENDCASE.
*    ENDLOOP.
*
*    DATA(target_array) = lcl_xlom__va_array=>create_initial( row_count    = max_row_count
*                                                                   column_count = max_column_count ).
*    DATA(row) = 1.
*    DO max_row_count TIMES.
*
*      DATA(column) = 1.
*      DO max_column_count TIMES.
*
*        DATA(single_cell_operands) = VALUE lif_xlom__ex=>tt_operand_result( ).
*        LOOP AT result-operand_results REFERENCE INTO operand_result.
*          IF     operand_result->not_part_of_result_array  = abap_false
*             AND operand_result->object                   IS BOUND.
*            IF operand_result->object->type = operand_result->object->c_type-array.
*              DATA(operand_result_array) = CAST lcl_xlom__va_array( operand_result->object ).
*              DATA(cell) = operand_result_array->lif_xlom__va_array~get_cell_value( column = column
*                                                                                          row    = row ).
*            ELSEIF operand_result->object->type = operand_result->object->c_type-range.
*              DATA(operand_result_range) = CAST lcl_xlom_range( operand_result->object ).
*              cell = operand_result_range->cells( row    = row
*                                                  column = column ).
*            ELSE.
*              cell = operand_result->object.
*            ENDIF.
*          ELSE.
*            cell = operand_result->object.
*          ENDIF.
*          INSERT VALUE #( name   = operand_result->name
*                          object = cell )
*                 INTO TABLE single_cell_operands.
*        ENDLOOP.
*        DATA(single_cell_result) = expression->evaluate_single( arguments = single_cell_operands
*                                                                context   = context ).
*        target_array->lif_xlom__va_array~set_cell_value( row    = row
*                                                               column = column
*                                                               value  = single_cell_result ).
*
*        column = column + 1.
*      ENDDO.
*
*      row = row + 1.
*    ENDDO.
*    result-result = target_array.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_lexer IMPLEMENTATION.
*  METHOD complete_with_non_matches.
*    DATA(last_offset) = 0.
*    LOOP AT c_matches ASSIGNING FIELD-SYMBOL(<match>).
*      IF <match>-offset > last_offset.
*        INSERT VALUE match_result( offset = last_offset
*                                   length = <match>-offset - last_offset ) INTO c_matches INDEX sy-tabix.
*      ENDIF.
*      last_offset = <match>-offset + <match>-length.
*    ENDLOOP.
*    IF strlen( i_string ) > last_offset.
*      APPEND VALUE match_result( offset = last_offset
*                                 length = strlen( i_string ) - last_offset ) TO c_matches.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD create.
*    result = NEW lcl_xlom__ex_ut_lexer( ).
*  ENDMETHOD.
*
*  METHOD lexe.
*    TYPES ty_ref_to_parenthesis_group TYPE REF TO ts_parenthesis_group.
*
*    " Note about `[ ` and ` ]` (https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e):
*    "   > Use the space character to improve readability in a structured reference
*    "   > You can use space characters to improve the readability of a structured reference.
*    "   > For example: =DeptSales[ [Sales Person]:[Region] ] or =DeptSales[[#Headers], [#Data], [% Commission]]"
*    "   > It’s recommended to use one space:
*    "   >   - After the first left bracket ([)
*    "   >   - Preceding the last right bracket (]).
*    "   >   - After a comma.
*    "
*    " Between `[` and `]`, the escape character is `'` e.g. `['[value']]` for the column header `[value]`.
*    "
*    " Note: -- is not an operator, it's a chain of the unary "-" operator (there could be even 3 or more subsequent unary operators); + can also be a unary operator,
*    "       hence the formula +--++-1 is a valid formula which simply means -1. https://stackoverflow.com/questions/3286197/what-does-do-in-excel-formulas
*    FIND ALL OCCURRENCES OF REGEX '(?:'
*                                & '\('
*                                & '|\{'
*                                & '|\[ '             " opening bracket after table name
*                                & '|\['              " table column name, each character can be:
*                                    & '(?:''.'       "   either one single quote (escape) with next character
*                                    & '|[^\[\]]'       "   or any other character except [ and ]
*                                    & ')+'
*                                    & '\]'
*                                & '|\['              " opening bracket after table name
*                                & '|\)'
*                                & '|\}'
*                                & '| ?\]'
*                                & '|, ?'
*                                & '|;'
*                                & '|:'
*                                & '|<>'
*                                & '|<='
*                                & '|>='
*                                & '|<'
*                                & '|>'
*                                & '|='
*                                & '|\+'
*                                & '|-'
*                                & '|\*'
*                                & '|/'
*                                & '|\^'
*                                & '|&'
*                                & '|%'
*                                & '|"(?:""|[^"])*"'  " string literal
*                                & '|#[A-Z0-9/!?]+'      " error name (#DIV/0!, #N/A, #VALUE!, #GETTING_DATA!, #NAME?, etc.)
*                                & ')'
*            IN text RESULTS DATA(matches).
*
*    complete_with_non_matches( EXPORTING i_string  = text
*                               CHANGING  c_matches = matches ).
*
*    DATA(token_values) = value string_table( ).
*    LOOP AT matches REFERENCE INTO DATA(match).
*      INSERT substring( val = text
*                        off = match->offset
*                        len = match->length )
*             INTO TABLE token_values.
*    ENDLOOP.
*
*    DATA(current_parenthesis_group) = VALUE ty_ref_to_parenthesis_group( ).
*    DATA(parenthesis_group) = VALUE ts_parenthesis_group( ).
*    DATA(parenthesis_level) = 0.
*    DATA(table_specification) = abap_false.
*    DATA(token) = VALUE ts_token( ).
*    DATA(token_number) = 1.
*    LOOP AT token_values REFERENCE INTO DATA(token_value).
*      " is comma a separator or a union operator?
*      " https://techcommunity.microsoft.com/t5/excel/does-the-union-operator-exist/m-p/2590110
*      " With argument-list functions, there is no union. Example: A1 contains 1, both =SUM(A1,A1) and =SUM((A1,A1)) return 2.
*      " With no-argument-list functions, there is a union. Example: =LARGE((A1,B1),2) (=LARGE(A1,B1,2) is invalid, too many arguments)
*      CASE token_value->*.
*        WHEN '('
*          OR '[]'
*          OR '['
*          OR `[ `
*          OR '{'
*          OR ')'
*          OR '}'
*          OR ']'
*          OR ` ]`
*          OR ',' " separator or union operator?
*          OR `, `
*          OR ';'.
*          token = VALUE #( value = condense( token_value->* )
*                           type  = condense( token_value->* ) ).
*        WHEN ` `
*          OR ':' " =B1:A1:B2:B3:A1:B2:B2:B3:B2 is same as =A1:B3
*          OR '<>'
*          OR '<='
*          OR '>='
*          OR '<'
*          OR '>'
*          OR '='
*          OR '+'
*          OR '-'
*          OR '*'
*          OR '/'  " 10/2 = 5
*          OR '^'  " 10^2 = 100
*          OR '&'  " "A"&"B" = "AB"
*          OR '%'. " 10% = 0.1
*          token = VALUE #( value = token_value->*
*                           type  = 'O' ).
*        WHEN OTHERS.
*          DATA(first_character) = substring( val = token_value->* len = 1 ).
*          IF first_character = '"'.
*            " text literal
*            token = VALUE #( value = replace( val  = substring( val = token_value->*
*                                                                off = 1
*                                                                len = strlen( token_value->* ) - 2 )
*                                              sub  = '""'
*                                              with = '"'
*                                              occ  = 0 )
*                             type  = c_type-text_literal ).
*          ELSEIF first_character = '['.
*            " table argument
*            token = VALUE #( value = token_value->*
*                             type  = c_type-square_bracket_open ).
*          ELSEIF first_character = '#'.
*            " error name
*            token = VALUE #( value = token_value->*
*                             type  = c_type-error_name ).
*          ELSEIF first_character CA '0123456789.-+'.
*            " number
*            token = VALUE #( value = token_value->*
*                             type  = c_type-number ).
*          ELSE.
*            " function name, --, cell reference, table name, name of named range, constant (TRUE, FALSE)
*            TYPES ty_ref_to_string TYPE REF TO string.
*            DATA(next_token_value) = COND ty_ref_to_string( WHEN token_number < lines( token_values )
*                                                            THEN REF #( token_values[
*                                                                            token_number + 1 ] ) ).
*            DATA(token_type) = c_type-symbol_name.
*            IF next_token_value IS BOUND.
*              DATA(next_token_first_character) = substring( val = next_token_value->*
*                                                            len = 1 ).
*              CASE next_token_first_character.
*                WHEN '('.
*                  token_type = c_type-function_name.
*                  DELETE token_values INDEX token_number + 1.
*                WHEN '['.
*                  token_type = c_type-table_name.
*                  CASE next_token_value->*.
*                    WHEN `[` OR `[ `.
*                      DELETE token_values INDEX token_number + 1.
*                    WHEN OTHERS.
*                      IF strlen( next_token_value->* ) > 2.
*                        " Excel formula "table[column]"; 1 token "[column]" becomes 2 tokens "[column]" and "]".
*                        INSERT ']' INTO token_values INDEX token_number + 2.
*                      ENDIF.
*                  ENDCASE.
*              ENDCASE.
*            ENDIF.
*            token = VALUE #( value = token_value->*
*                             type  = token_type ).
*          ENDIF.
*      ENDCASE.
*
*      INSERT token INTO TABLE result.
*      token_number = token_number + 1.
*    ENDLOOP.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_operator IMPLEMENTATION.
*  METHOD class_constructor.
*    LOOP AT VALUE tt_operator(
*        ( name = ':'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 1 desc = 'range A1:A2 or A1:A2:A2' )
*        ( name = ` `                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 1 desc = 'intersection A1 A2' )
*        ( name = ','                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 1 desc = 'union A1,A2' )
*        ( name = '-' unary = abap_true operand_position_offsets = VALUE #( ( +1 ) )        priority = 2 desc = '-1' )
*        ( name = '+' unary = abap_true operand_position_offsets = VALUE #( ( +1 ) )        priority = 2 desc = '+1' )
*        ( name = '%'                   operand_position_offsets = VALUE #( ( -1 ) )        priority = 3 desc = 'percent e.g. 10%' )
*        ( name = '^'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 4 desc = 'exponent 2^8' )
*        ( name = '*'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 5 desc = '2*2' )
*        ( name = '/'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 5 desc = '2/2' )
*        ( name = '+'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 6 desc = '2+2' )
*        ( name = '-'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 6 desc = '2-2' )
*        ( name = '&'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 7 desc = 'concatenate "A"&"B"' )
*        ( name = '='                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1=1' )
*        ( name = '<'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1<1' )
*        ( name = '>'                   operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1>1' )
*        ( name = '<='                  operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1<=1' )
*        ( name = '>='                  operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1>=1' )
*        ( name = '<>'                  operand_position_offsets = VALUE #( ( -1 ) ( +1 ) ) priority = 8 desc = 'A1<>1' ) )
*         REFERENCE INTO DATA(operator).
*      create( name                     = operator->name
*              unary                    = operator->unary
*              operand_position_offsets = operator->operand_position_offsets
*              priority                 = operator->priority
*              description              = operator->desc ).
*    ENDLOOP.
*  ENDMETHOD.
*
*  METHOD create.
*    result = NEW lcl_xlom__ex_ut_operator( ).
*    result->name                     = name.
*    result->operand_position_offsets = operand_position_offsets.
*    result->priority                 = priority.
*    result->unary                    = unary.
*    INSERT VALUE #( name     = name
*                    unary    = unary
*                    priority = priority
*                    desc     = description
*                    handler  = result )
*           INTO TABLE operators.
*  ENDMETHOD.
*
*  METHOD create_expression.
*    CASE name.
*      WHEN '+'.
*        IF unary = abap_true.
*          RAISE EXCEPTION TYPE lcx_xlom_todo.
*        ENDIF.
*        result = lcl_xlom__ex_op_plus=>create( left_operand  = operands[ 1 ]
*                                                   right_operand = operands[ 2 ] ).
*      WHEN '-'.
*        IF unary = abap_false.
*          result = lcl_xlom__ex_op_minus=>create( left_operand  = operands[ 1 ]
*                                                      right_operand = operands[ 2 ] ).
*        ELSE.
*          result = lcl_xlom__ex_op_minus_unry=>create( operand = operands[ 1 ] ).
*        ENDIF.
*      WHEN '*'.
*        result = lcl_xlom__ex_op_mult=>create( left_operand  = operands[ 1 ]
*                                                   right_operand = operands[ 2 ] ).
*      WHEN '='.
*        result = lcl_xlom__ex_op_equal=>create( left_operand  = operands[ 1 ]
*                                                    right_operand = operands[ 2 ] ).
*      WHEN '&'.
*        result = lcl_xlom__ex_op_ampersand=>create( left_operand  = operands[ 1 ]
*                                                        right_operand = operands[ 2 ] ).
*      WHEN ':'.
*        result = lcl_xlom__ex_op_colon=>create( left_operand  = operands[ 1 ]
*                                                        right_operand = operands[ 2 ] ).
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD get.
*    result = VALUE #( operators[ name  = operator
*                                 unary = unary ]-handler OPTIONAL ).
*    IF result IS NOT BOUND.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD get_operand_position_offsets.
*    result = operand_position_offsets.
*  ENDMETHOD.
*
*  METHOD get_priority.
*    result = priority.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_parser IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_ut_parser( ).
*  ENDMETHOD.
*
*  METHOD get_expression_from_curly_brac.
*    result = lcl_xlom__ex_el_array=>create(
*                 rows = VALUE #( FOR <argument> IN arguments
*                                 ( columns_of_row = VALUE #( FOR <subitem> IN <argument>->subitems
*                                                             ( <subitem>->expression ) ) ) ) ).
*  ENDMETHOD.
*
*  METHOD get_expression_from_error.
*    result = lcl_xlom__ex_el_error=>get_from_error_name( error_name ).
*  ENDMETHOD.
*
*  METHOD get_expression_from_function.
*    CASE function_name.
*      WHEN 'ADDRESS'.
*        result = lcl_xlom__ex_fu_address=>create(
*                     row_num    = arguments[ 1 ]
*                     column_num = arguments[ 2 ]
*                     abs_num    = COND #( WHEN line_exists( arguments[ 3 ] ) THEN CAST #( arguments[ 3 ] ) )
*                     a1         = COND #( WHEN line_exists( arguments[ 4 ] ) THEN CAST #( arguments[ 4 ] ) )
*                     sheet_text = COND #( WHEN line_exists( arguments[ 5 ] ) THEN CAST #( arguments[ 5 ] ) ) ).
*      WHEN 'CELL'.
*        result = lcl_xlom__ex_fu_cell=>create(
*                     info_type = CAST #( arguments[ 1 ] )
*                     reference = COND #( WHEN line_exists( arguments[ 2 ] ) THEN CAST #( arguments[ 2 ] ) ) ).
*      WHEN 'COUNTIF'.
*        result = lcl_xlom__ex_fu_countif=>create( range    = arguments[ 1 ]
*                                                        criteria = arguments[ 2 ] ).
*      WHEN 'FIND'.
*        result = lcl_xlom__ex_fu_find=>create(
*                     find_text   = arguments[ 1 ]
*                     within_text = arguments[ 2 ]
*                     start_num   = COND #( WHEN line_exists( arguments[ 3 ] ) THEN arguments[ 3 ] ) ).
*      WHEN 'IF'.
*        result = lcl_xlom__ex_fu_if=>create( condition     = arguments[ 1 ]
*                                                   expr_if_true  = arguments[ 2 ]
*                                                   expr_if_false = arguments[ 3 ] ).
*      WHEN 'IFERROR'.
*        result = lcl_xlom__ex_fu_iferror=>create( value          = arguments[ 1 ]
*                                                        value_if_error = arguments[ 2 ] ).
*      WHEN 'INDEX'.
*        result = lcl_xlom__ex_fu_index=>create(
*                   array      = arguments[ 1 ]
*                   row_num    = arguments[ 2 ]
*                   column_num = arguments[ 3 ] ).
*      WHEN 'INDIRECT'.
*        result = lcl_xlom__ex_fu_indirect=>create( ref_text = arguments[ 1 ]
*                                                         a1       = COND #( WHEN line_exists( arguments[ 2 ] ) THEN arguments[ 2 ] ) ).
*      WHEN 'LEN'.
*        result = lcl_xlom__ex_fu_len=>create( text = arguments[ 1 ] ).
*      WHEN 'MATCH'.
*        result = lcl_xlom__ex_fu_match=>create(
*                   lookup_value = arguments[ 1 ]
*                   lookup_array = arguments[ 2 ]
*                   match_type   = COND #( WHEN line_exists( arguments[ 3 ] ) THEN arguments[ 3 ] ) ).
*      WHEN 'OFFSET'.
*        result = lcl_xlom__ex_fu_offset=>create(
*                     reference = arguments[ 1 ]
*                     rows      = arguments[ 2 ]
*                     cols      = arguments[ 3 ]
*                     height    = COND #( WHEN line_exists( arguments[ 4 ] ) THEN arguments[ 4 ] )
*                     width     = COND #( WHEN line_exists( arguments[ 5 ] ) THEN arguments[ 5 ] ) ).
*      WHEN 'RIGHT'.
*        result = lcl_xlom__ex_fu_right=>create( text      = arguments[ 1 ]
*                                                      num_chars = COND #( WHEN line_exists( arguments[ 2 ] ) THEN arguments[ 2 ] ) ).
*      WHEN 'ROW'.
*        result = lcl_xlom__ex_fu_row=>create( reference = COND #( WHEN line_exists( arguments[ 1 ] ) THEN arguments[ 1 ] ) ).
*      WHEN 'T'.
*        result = lcl_xlom__ex_fu_t=>create( value = arguments[ 1 ] ).
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD get_expression_from_symbol_nam.
*    IF token_value CP 'true'.
*      result = lcl_xlom__ex_el_boolean=>create( boolean_value = abap_true ).
*    ELSEIF token_value CP 'false'.
*      result = lcl_xlom__ex_el_boolean=>create( boolean_value = abap_false ).
*    ELSE.
*      result = lcl_xlom__ex_el_range=>create( token_value ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD parse.
*    current_token_index = 1.
*    me->tokens             = tokens.
*
*    DATA(initial_item) = lcl_xlom__ex_ut_parser_item=>create( type = '(' ).
*    current_token_index = 0.
*
*    " Determine the groups for the parentheses (expression grouping,
*    " function arguments), curly brackets (arrays) and square brackets (tables).
*    parse_tokens_until( EXPORTING until = lcl_xlom__ex_ut_lexer=>c_type-parenthesis_close
*                        CHANGING  item  = initial_item ).
*
*    " Determine the items for the operators.
*    parse_expression_item_2( CHANGING item = initial_item ).
*
**    " Merge function item with its next item (arguments in parentheses) and ignore commas between arguments.
*    " ignore commas between arguments of functions and tables.
*    parse_expression_item_1( CHANGING item = initial_item ).
*
*    " arrays
*    parse_expression_item_5( CHANGING item = initial_item ).
*
*    " Remove useless items of one item.
*    parse_expression_item_3( CHANGING item = initial_item ).
*
*    " Determine the expressions for each item.
*    parse_expression_item_4( CHANGING item = initial_item ).
*
*    result = initial_item->expression.
*  ENDMETHOD.
*
*  METHOD parse_expression_item.
*    WHILE current_token_index < lines( tokens ).
*      current_token_index = current_token_index + 1.
*      DATA(current_token) = REF #( tokens[ current_token_index ] ).
*      DATA(subitem) = NEW lcl_xlom__ex_ut_parser_item( ).
*      subitem->type = current_token->type.
*      subitem->value = current_token->value.
*      CASE current_token->type.
*        WHEN lcl_xlom__ex_ut_lexer=>c_type-parenthesis_open
*          OR lcl_xlom__ex_ut_lexer=>c_type-function_name.
*          parse_tokens_until( EXPORTING until = lcl_xlom__ex_ut_lexer=>c_type-parenthesis_close
*                              CHANGING  item  = subitem ).
*        WHEN lcl_xlom__ex_ut_lexer=>c_type-curly_bracket_open.
*          parse_tokens_until( EXPORTING until = lcl_xlom__ex_ut_lexer=>c_type-curly_bracket_close
*                              CHANGING  item  = subitem ).
*        WHEN lcl_xlom__ex_ut_lexer=>c_type-table_name.
*          parse_tokens_until( EXPORTING until = lcl_xlom__ex_ut_lexer=>c_type-square_bracket_close
*                              CHANGING  item  = subitem ).
*      ENDCASE.
*      INSERT subitem INTO TABLE item->subitems.
*    ENDWHILE.
*  ENDMETHOD.
*
*  METHOD parse_tokens_until.
*    WHILE current_token_index < lines( tokens ).
*      current_token_index = current_token_index + 1.
*      DATA(current_token) = REF #( tokens[ current_token_index ] ).
*      DATA(subitem) = NEW lcl_xlom__ex_ut_parser_item( ).
*      subitem->type = current_token->type.
*      subitem->value = current_token->value.
*      CASE current_token->type.
*        WHEN lcl_xlom__ex_ut_lexer=>c_type-parenthesis_open
*          OR lcl_xlom__ex_ut_lexer=>c_type-function_name.
*          parse_tokens_until( EXPORTING until = lcl_xlom__ex_ut_lexer=>c_type-parenthesis_close
*                              CHANGING  item  = subitem ).
*        WHEN lcl_xlom__ex_ut_lexer=>c_type-curly_bracket_open.
*          parse_tokens_until( EXPORTING until = lcl_xlom__ex_ut_lexer=>c_type-curly_bracket_close
*                              CHANGING  item  = subitem ).
*        WHEN lcl_xlom__ex_ut_lexer=>c_type-table_name.
*          parse_tokens_until( EXPORTING until = lcl_xlom__ex_ut_lexer=>c_type-square_bracket_close
*                              CHANGING  item  = subitem ).
*        WHEN until.
*          RETURN.
*      ENDCASE.
*      INSERT subitem INTO TABLE item->subitems.
*    ENDWHILE.
*  ENDMETHOD.
*
*  METHOD parse_expression_item_1.
*    LOOP AT item->subitems INTO DATA(subitem).
*      parse_expression_item_1( CHANGING item = subitem ).
*    ENDLOOP.
*
*    CASE item->type.
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-function_name
*            OR lcl_xlom__ex_ut_lexer=>c_type-table_name.
*        LOOP AT item->subitems INTO subitem.
*          DATA(subitem_index) = sy-tabix.
*          IF subitem->type = lcl_xlom__ex_ut_lexer=>c_type-comma.
*            IF    subitem_index = lines( item->subitems )
*               OR (     subitem_index < lines( item->subitems )
*                    AND item->subitems[ subitem_index + 1 ]->type = lcl_xlom__ex_ut_lexer=>c_type-comma ).
*              " RIGHT("hello",) (means RIGHT("hello",0)) -> arguments "hello" and empty
*              " NB: RIGHT("hello") means RIGHT("hello",1)
*              INSERT lcl_xlom__ex_ut_parser_item=>create( type = lcl_xlom__ex_ut_lexer=>c_type-empty_argument ) INTO item->subitems INDEX subitem_index + 1.
*            ENDIF.
*            DELETE item->subitems USING KEY loop_key.
*          ENDIF.
*        ENDLOOP.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD parse_expression_item_2.
*    TYPES to_expression TYPE REF TO lif_xlom__ex.
*    TYPES:
*      BEGIN OF ts_work,
*        position   TYPE sytabix,
*        token      TYPE REF TO lcl_xlom__ex_ut_lexer=>ts_token,
*        expression TYPE REF TO lif_xlom__ex,
*        operator   TYPE REF TO lcl_xlom__ex_ut_operator,
*        priority   TYPE i,
*      END OF ts_work.
*    TYPES tt_work TYPE SORTED TABLE OF ts_work WITH NON-UNIQUE KEY position
*                    WITH NON-UNIQUE SORTED KEY by_priority COMPONENTS priority position.
*    TYPES:
*      BEGIN OF ts_subitem_by_priority,
*        operator TYPE REF TO lcl_xlom__ex_ut_operator,
*        priority TYPE i,
*        subitem_index type i,
*        subitem  TYPE REF TO lcl_xlom__ex_ut_parser_item,
*      END OF ts_subitem_by_priority.
*    TYPES tt_subitem_by_priority TYPE SORTED TABLE OF ts_subitem_by_priority WITH NON-UNIQUE KEY priority.
*    TYPES tt_operand_positions TYPE STANDARD TABLE OF i WITH EMPTY KEY.
*
*    DATA priorities TYPE SORTED TABLE OF i WITH UNIQUE KEY table_line.
*    DATA subitems_by_priority TYPE tt_subitem_by_priority.
*    DATA item_index TYPE syst-tabix.
*
*    LOOP AT item->subitems INTO DATA(subitem).
*      parse_expression_item_2( CHANGING item = subitem ).
*    ENDLOOP.
*
*    DATA(work_table) = VALUE tt_work( ).
*    LOOP AT item->subitems INTO subitem
*         WHERE table_line->type = lcl_xlom__ex_ut_lexer=>c_type-operator.
*      item_index = sy-tabix.
*      DATA(unary) = COND abap_bool( WHEN item_index = 1
*                                    THEN abap_true
*                                    ELSE SWITCH #( item->subitems[ item_index - 1 ]->type
*                                                   WHEN lcl_xlom__ex_ut_lexer=>c_type-operator
*                                                     OR lcl_xlom__ex_ut_lexer=>c_type-comma
*                                                     OR lcl_xlom__ex_ut_lexer=>c_type-parenthesis_open
*                                                     OR lcl_xlom__ex_ut_lexer=>c_type-semicolon
*                                                   THEN abap_true ) ).
*      DATA(operator) = lcl_xlom__ex_ut_operator=>get( operator = subitem->value
*                                                        unary    = unary ).
*      DATA(priority) = operator->get_priority( ).
*      INSERT priority INTO TABLE priorities.
*      INSERT VALUE #( priority      = priority
*                      operator      = operator
*                      subitem_index = item_index
*                      subitem       = subitem )
*             INTO TABLE subitems_by_priority.
*    ENDLOOP.
*
*    " Process operators with priority 1 first, then 2, etc.
*    " The priority 0 corresponds to functions, tables, boolean values, numeric literals and text literals.
*    LOOP AT priorities INTO priority.
*      LOOP AT subitems_by_priority REFERENCE INTO data(subitem_by_priority)
*           WHERE priority = priority.
*
*        DATA(operand_position_offsets) = subitem_by_priority->operator->get_operand_position_offsets( ).
*
*        subitem_by_priority->subitem->subitems = VALUE #( ).
*        LOOP AT operand_position_offsets INTO DATA(operand_position_offset).
*          INSERT item->subitems[ subitem_by_priority->subitem_index + operand_position_offset ]
*              INTO TABLE subitem_by_priority->subitem->subitems.
*        ENDLOOP.
*
*        DATA(positions_of_operands_to_delet) = VALUE tt_operand_positions(
*                                                         FOR <operand_position_offset> IN operand_position_offsets
*                                                         ( subitem_by_priority->subitem_index + <operand_position_offset> ) ).
*        SORT positions_of_operands_to_delet BY table_line DESCENDING.
*        LOOP AT positions_of_operands_to_delet INTO DATA(position).
*          DELETE item->subitems index position.
*        ENDLOOP.
*      ENDLOOP.
*    ENDLOOP.
*  ENDMETHOD.
*
*  METHOD parse_expression_item_3.
*    LOOP AT item->subitems REFERENCE INTO DATA(subitem).
*      parse_expression_item_3( CHANGING item = subitem->* ).
*    ENDLOOP.
*
*    IF     item->type = lcl_xlom__ex_ut_lexer=>c_type-parenthesis_open
*       AND lines( item->subitems ) = 1.
*      item = item->subitems[ 1 ].
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD parse_expression_item_4.
*    LOOP AT item->subitems INTO DATA(subitem).
*      parse_expression_item_4( CHANGING item = subitem ).
*    ENDLOOP.
*
*    CASE item->type.
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-curly_bracket_open.
*        item->expression = get_expression_from_curly_brac( arguments = item->subitems ).
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-empty_argument.
*        item->expression = lcl_xlom__ex_el_empty_arg=>create( ).
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-error_name.
*        item->expression = get_expression_from_error( error_name = item->value ).
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-function_name.
*        " FUNCTION NAME
*        item->expression = get_expression_from_function( function_name = item->value
*                                                         arguments     = VALUE #( FOR <subitem> IN item->subitems
*                                                                                  ( <subitem>->expression ) ) ).
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-number.
*        " NUMBER
*        item->expression = lcl_xlom__ex_el_number=>create( CONV #( item->value ) ).
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-operator.
*        " OPERATOR
*        item->expression = lcl_xlom__ex_ut_operator=>get( operator = item->value
*                                         unary    = xsdbool( lines( item->subitems ) = 1 ) )->create_expression(
*                                             operands = VALUE #( FOR <subitem> IN item->subitems
*                                                                 ( <subitem>->expression ) ) ).
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-semicolon.
*        " Array row separator. No special processing, it's handled inside curly brackets.
*        ASSERT 1 = 1. " Debug helper to set a break-point
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-symbol_name.
*        " SYMBOL NAME (range address, range name, boolean constant)
*        item->expression = get_expression_from_symbol_nam( item->value ).
*      WHEN lcl_xlom__ex_ut_lexer=>c_type-text_literal.
*        " TEXT LITERAL
*        item->expression = lcl_xlom__ex_el_string=>create( item->value ).
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD parse_expression_item_5.
*      types ty_ref_item type ref to lcl_xlom__ex_ut_parser_item.
*
*    LOOP AT item->subitems INTO DATA(subitem).
*      parse_expression_item_5( CHANGING item = subitem ).
*    ENDLOOP.
*
*    LOOP AT item->subitems INTO DATA(array)
*         WHERE table_line->type = lcl_xlom__ex_ut_lexer=>c_type-curly_bracket_open.
*      DATA(new_array_subitems) = VALUE lcl_xlom__ex_ut_parser_item=>tt_item( ).
*      DATA(row) = VALUE ty_ref_item( ).
*      LOOP AT array->subitems INTO DATA(array_subitem).
*        CASE array_subitem->type.
*          WHEN lcl_xlom__ex_ut_lexer=>c_type-comma.
*          WHEN lcl_xlom__ex_ut_lexer=>c_type-semicolon.
*            INSERT row INTO TABLE new_array_subitems.
*            row = VALUE #( ).
*          WHEN OTHERS.
*            IF row IS NOT BOUND.
*              row = NEW lcl_xlom__ex_ut_parser_item( ).
*              row->type  = lcl_xlom__ex_ut_lexer=>c_type-semicolon.
*              row->value = lcl_xlom__ex_ut_lexer=>c_type-semicolon.
*            ENDIF.
*            INSERT array_subitem INTO TABLE row->subitems.
*        ENDCASE.
*      ENDLOOP.
*      INSERT row INTO TABLE new_array_subitems.
*      array->subitems = new_array_subitems.
*    ENDLOOP.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_ut_parser_item IMPLEMENTATION.
*  METHOD create.
*    RESULT = new lcl_xlom__ex_ut_parser_item( ).
*    result->type = type.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_array IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_el_array( ).
*    result->lif_xlom__ex~type = result->lif_xlom__ex~c_type-array.
*    result->rows = rows.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    result = lif_xlom__ex~set_result( lcl_xlom__va_array=>create_initial(
*               row_count    = lines( rows )
*               column_count = REDUCE #( init n = 0
*                                             FOR <row> in rows
*                                             next n = nmax( val1 = n val2 = lines( <row>-columns_of_row ) ) )
*               rows              = VALUE #( for <row> in rows
*                                            ( columns_of_row = VALUE #( for <column> in <row>-columns_of_row
*                                                                        ( <column>->evaluate( context ) ) ) ) ) ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    DATA(array) = CAST lcl_xlom__ex_el_array( expression ).
*    IF lines( rows ) <> lines( array->rows ).
*      RETURN.
*    ENDIF.
*
*    DATA(row_tabix) = 1.
*    WHILE row_tabix <= lines( rows ).
*
*      DATA(ref_columns) = REF #( rows[ row_tabix ]-columns_of_row ).
*      DATA(ref_array_columns) = REF #( array->rows[ row_tabix ]-columns_of_row ).
*      IF lines( ref_columns->* ) <> lines( ref_array_columns->* ).
*        result = abap_false.
*        RETURN.
*      ENDIF.
*
*      DATA(column_tabix) = 1.
*      WHILE column_tabix <= lines( ref_columns->* ).
*        IF abap_false = ref_columns->*[ column_tabix ]->is_equal( ref_array_columns->*[ column_tabix ] ).
*          RETURN.
*        ENDIF.
*        column_tabix = column_tabix + 1.
*      ENDWHILE.
*
*      row_tabix = row_tabix + 1.
*    ENDWHILE.
*
*    result = abap_true.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_boolean IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_el_boolean( ).
*    result->boolean_value = boolean_value.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-boolean.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    result = lif_xlom__ex~set_result( lcl_xlom__va_boolean=>get( boolean_value ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_empty_arg IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_el_empty_arg( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-empty_argument.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    result = lif_xlom__ex~set_result( lcl_xlom__va_empty=>get_singleton( ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    result = xsdbool( expression->type = expression->c_type-empty_argument ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_error IMPLEMENTATION.
*  METHOD class_constructor.
*    blocked                    = lcl_xlom__ex_el_error=>create( english_error_name    = '#BLOCKED!     '
*                                                                 internal_error_number = 2047 ).
*    calc                       = lcl_xlom__ex_el_error=>create( english_error_name    = '#CALC!        '
*                                                                 internal_error_number = 2050 ).
*    connect                    = lcl_xlom__ex_el_error=>create( english_error_name    = '#CONNECT!     '
*                                                                 internal_error_number = 2046 ).
*    division_by_zero           = lcl_xlom__ex_el_error=>create( english_error_name    = '#DIV/0!       '
*                                                                 internal_error_number = 2007 ).
*    field                      = lcl_xlom__ex_el_error=>create( english_error_name    = '#FIELD!       '
*                                                                 internal_error_number = 2049 ).
*    getting_data               = lcl_xlom__ex_el_error=>create( english_error_name    = '#GETTING_DATA!'
*                                                                 internal_error_number = 2043 ).
*    na_not_applicable          = lcl_xlom__ex_el_error=>create( english_error_name    = '#N/A          '
*                                                                 internal_error_number = 2042 ).
*    name                       = lcl_xlom__ex_el_error=>create( english_error_name    = '#NAME?        '
*                                                                 internal_error_number = 2029 ).
*    null                       = lcl_xlom__ex_el_error=>create( english_error_name    = '#NULL!        '
*                                                                 internal_error_number = 2000 ).
*    num                        = lcl_xlom__ex_el_error=>create( english_error_name    = '#NUM!         '
*                                                                 internal_error_number = 2036 ).
*    python                     = lcl_xlom__ex_el_error=>create( english_error_name    = '#PYTHON!      '
*                                                                 internal_error_number = 2222 ).
*    ref                        = lcl_xlom__ex_el_error=>create( english_error_name    = '#REF!         '
*                                                                 internal_error_number = 2023 ).
*    spill                      = lcl_xlom__ex_el_error=>create( english_error_name    = '#SPILL!       '
*                                                                 internal_error_number = 2045 ).
*    unknown                    = lcl_xlom__ex_el_error=>create( english_error_name    = '#UNKNOWN!     '
*                                                                 internal_error_number = 2048 ).
*    value_cannot_be_calculated = lcl_xlom__ex_el_error=>create( english_error_name    = '#VALUE!       '
*                                                                 internal_error_number = 2015 ).
*  ENDMETHOD.
*
*  METHOD create.
*    result = NEW lcl_xlom__ex_el_error( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-error.
*    result->english_error_name    = english_error_name.
*    result->internal_error_number = internal_error_number.
*    INSERT VALUE #( english_error_name    = english_error_name
*                    internal_error_number = internal_error_number
*                    object                = result )
*           INTO TABLE errors.
*  ENDMETHOD.
*
*  METHOD get_from_error_name.
*    result = errors[ english_error_name = error_name ]-object.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    result = lif_xlom__ex~set_result( lcl_xlom__va_error=>get_by_error_number( internal_error_number ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_address IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_address( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-address.
*    result->row_num    = row_num.
*    result->column_num = column_num.
*    result->abs_num    = abs_num.
*    result->a1         = a1.
*    result->sheet_text = sheet_text.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_cell IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_cell( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-cell.
*    result->info_type = info_type.
*    result->reference = reference.
*    if reference is not bound.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    endif.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA temp_result TYPE REF TO lcl_xlom__va_string.
*
*    TRY.
*    " In cell B1, formula =CELL("address",A1:A6) is the same result as =CELL("address",A1), which is $A$1 in cell B1;
*    " the cells B2 to B6 are not initialized with $A$2, $A$3, etc.
*    DATA(info_type_result) = lcl_xlom__va=>to_string( info_type->lif_xlom__ex~evaluate( context ) )->get_string( ).
*    CASE info_type_result.
*      WHEN c_info_type-filename.
*        " Retourne par exemple "C:\temp\[Book1.xlsx]Sheet1"
*        temp_result = lcl_xlom__va_string=>create( context->worksheet->parent->path
*                                                    && |\\[{ context->worksheet->parent->name }]|
*                                                    && context->worksheet->name ).
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*    result = lif_xlom__ex~set_result( temp_result ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_countif IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_countif( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-countif.
*    result->range    = range.
*    result->criteria = criteria.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'RANGE   ' object = range    not_part_of_result_array = abap_true )
*                                                       ( name = 'CRITERIA' object = criteria ) ) ).
*    result = lif_xlom__ex~set_result( COND #( WHEN array_evaluation-result IS BOUND
*                                                  THEN array_evaluation-result
*                                                  ELSE lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                                                         context   = context ) ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    DATA(result_of_range) = CAST lif_xlom__va_array( arguments[ name = 'RANGE' ]-object ).
*    DATA(result_of_criteria) = lcl_xlom__va=>to_string( arguments[ name = 'CRITERIA' ]-object )->get_string( ).
*    " criteria may start with the comparison operator =, <, <=, >, >=, <>.
*    " It's followed by the string to compare.
*    " Anything else will be equivalent to the same string prefixed with the = comparison operator.
*    IF strlen( result_of_criteria ) >= 1.
*      CASE substring( val = result_of_criteria
*                      len = 1 ).
*        WHEN '=' OR '<' OR '>'.
*          RAISE EXCEPTION TYPE lcx_xlom_todo.
*      ENDCASE.
*    ENDIF.
*    DATA(count_cells) = 0.
*    DATA(row_number) = 1.
*    WHILE row_number <= result_of_range->row_count.
*      DATA(column_number) = 1.
*      WHILE column_number <= result_of_range->column_count.
*        DATA(cell_value_string) = lcl_xlom__va=>to_string( result_of_range->get_cell_value(
*                                                                               column = column_number
*                                                                               row    = row_number ) )->get_string( ).
*        IF cell_value_string CP result_of_criteria.
*          count_cells = count_cells + 1.
*        ENDIF.
*        column_number = column_number + 1.
*      ENDWHILE.
*      row_number = row_number + 1.
*    ENDWHILE.
*    result = lcl_xlom__va_number=>create( CONV #( count_cells ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_find IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_find( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-find.
*    result->find_text   = find_text.
*    result->within_text = within_text.
*    result->start_num   = start_num.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'FIND_TEXT  ' object = find_text   )
*                                                       ( name = 'WITHIN_TEXT' object = within_text )
*                                                       ( name = 'START_NUM  ' object = start_num   ) ) ).
*    result = lif_xlom__ex~set_result( COND #( WHEN array_evaluation-result IS BOUND
*                                                  THEN array_evaluation-result
*                                                  ELSE lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                                                         context   = context ) ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    DATA(result_of_find_text) = lcl_xlom__va=>to_string( arguments[ name = 'FIND_TEXT' ]-object )->get_string( ).
*    DATA(result_of_within_text) = lcl_xlom__va=>to_string( arguments[ name = 'WITHIN_TEXT' ]-object )->get_string( ).
*    DATA(result_of_start_num) = CAST lcl_xlom__va_number( arguments[ name = 'START_NUM' ]-object ).
*    DATA(start_offset) = COND i( WHEN result_of_start_num IS BOUND THEN result_of_start_num->get_number( ) ).
*    IF start_offset > strlen( result_of_within_text ).
*      result = lcl_xlom__va_error=>value_cannot_be_calculated.
*    ELSE.
*      DATA(result_offset) = COND #( WHEN result_of_find_text IS INITIAL
*                                    THEN 1
*                                    ELSE find( val = result_of_within_text
*                                               sub = result_of_find_text
*                                               off = start_offset ) + 1 ).
*      IF result_offset = 0.
*        result = lcl_xlom__va_error=>value_cannot_be_calculated.
*      ELSE.
*        result = lcl_xlom__va_number=>create( CONV #( result_offset ) ).
*      ENDIF.
*    ENDIF.
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_if IMPLEMENTATION.
*  METHOD create.
*    IF    condition     IS NOT BOUND
*       OR expr_if_true  IS NOT BOUND
*       OR expr_if_false IS NOT BOUND.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    result = NEW lcl_xlom__ex_fu_if( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-if.
*    result->condition     = condition.
*    result->expr_if_true  = expr_if_true.
*    result->expr_if_false = expr_if_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(condition_evaluation) = lcl_xlom__va=>to_boolean( condition->evaluate( context ) ).
*    result = lif_xlom__ex~set_result( COND #( WHEN condition_evaluation = lcl_xlom__va_boolean=>true
*                                                  THEN expr_if_true->evaluate( context )
*                                                  ELSE expr_if_false->evaluate( context ) ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF expression->type = lif_xlom__ex=>c_type-function-if.
*      DATA(if) = CAST lcl_xlom__ex_fu_if( expression ).
*      IF     condition->is_equal( if->condition )
*         AND expr_if_true->is_equal( if->expr_if_true )
*         AND expr_if_false->is_equal( if->expr_if_false ).
*        result = abap_true.
*      ELSE.
*        result = abap_false.
*      ENDIF.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_iferror IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_iferror( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-iferror.
*    result->value                 = value.
*    result->value_if_error        = value_if_error.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'VALUE         ' object = value )
*                                                       ( name = 'VALUE_IF_ERROR' object = value_if_error ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    result = lif_xlom__ex~set_result( COND #( LET value_result = arguments[ name = 'VALUE' ]-object
*                                                  IN
*                                                  WHEN value_result->type = lif_xlom__va=>c_type-error
*                                                  THEN arguments[ name = 'VALUE_IF_ERROR' ]-object
*                                                  ELSE value_result ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_index IMPLEMENTATION.
*  METHOD create.
*    IF      array     IS NOT BOUND
*       OR ( row_num    IS NOT BOUND
*       and  column_num IS NOT BOUND ).
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    result = NEW lcl_xlom__ex_fu_index( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-index.
*    result->array      = array.
*    result->row_num    = row_num.
*    result->column_num = column_num.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    " INDEX(A1:D4,{1,2;3,4},{1,2;3,4}) will return {A1,B2;C3,D4}
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'ARRAY     ' object = array      not_part_of_result_array = abap_true )
*                                                       ( name = 'ROW_NUM   ' object = row_num    )
*                                                       ( name = 'COLUMN_NUM' object = column_num ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*        DATA(array_or_range) = lcl_xlom__va=>to_array( arguments[ name = 'ARRAY' ]-object ).
*        DATA(row) = lcl_xlom__va=>to_number( arguments[ name = 'ROW_NUM' ]-object )->get_number( ).
*        DATA(column) = lcl_xlom__va=>to_number( arguments[ name = 'COLUMN_NUM' ]-object )->get_number( ).
*        result = lif_xlom__ex~set_result( array_or_range->get_cell_value( column = EXACT #( column )
*                                                                           row    = EXACT #( row ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_indirect IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_indirect( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-indirect.
*    result->ref_text     = ref_text.
*    result->a1    = a1.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    TRY.
*    " INDIRECT("A1:D4")
*    DATA(ref_text_result) = lcl_xlom__va=>to_string( ref_text->evaluate( context ) )->get_string( ).
*    result = lif_xlom__ex~set_result( lcl_xlom_range=>create_from_address_or_name(
*                                           address     = ref_text_result
*                                           relative_to = context->worksheet ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_len IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_len( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-len.
*    result->text                  = text.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*        expression = me
*        context    = context
*        operands   = VALUE #( ( name = 'TEXT' object = text ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    result = lif_xlom__ex~set_result( lcl_xlom__va_number=>create(
*                                              strlen( lcl_xlom__va=>to_string(
*                                                          arguments[ name = 'TEXT' ]-object )->get_string( ) ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF expression->type <> lif_xlom__ex=>c_type-function-len.
*      RETURN.
*    ENDIF.
*    DATA(len) = CAST lcl_xlom__ex_fu_len( expression ).
*    result = xsdbool( text->is_equal( len->text ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_match IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_match( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-match.
*    result->lookup_value          = lookup_value.
*    result->lookup_array          = lookup_array.
*    result->match_type            = match_type.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'LOOKUP_VALUE' object = lookup_value )
*                                                       ( name = 'LOOKUP_ARRAY' object = lookup_array not_part_of_result_array = abap_true )
*                                                       ( name = 'MATCH_TYPE  ' object = match_type ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    DATA temp_result TYPE REF TO lif_xlom__va.
*
*    TRY.
*        DATA(lookup_value_result) = arguments[ name = 'LOOKUP_VALUE' ]-object.
*        DATA(lookup_array_result) = lcl_xlom__va=>to_array( arguments[ name = 'LOOKUP_ARRAY' ]-object ).
*        DATA(match_type_result) = COND i( LET result_num_chars = arguments[ name = 'MATCH_TYPE' ]-object IN
*                                          WHEN result_num_chars IS BOUND
*                                          THEN COND #( WHEN result_num_chars->type = result_num_chars->c_type-empty
*                                                       THEN 1
*                                                       ELSE lcl_xlom__va=>to_number( result_num_chars )->get_integer( ) ) ).
*        IF match_type_result <> 0.
*          RAISE EXCEPTION TYPE lcx_xlom_todo.
*        ENDIF.
*        IF     lookup_array_result->row_count    > 1
*           AND lookup_array_result->column_count > 1.
*          " MATCH cannot lookup a two-dimension array, it can search either one row or one column.
*          temp_result = lcl_xlom__va_error=>na_not_applicable.
*        ELSE.
*          DATA(optimized_lookup_array) = lcl_xlom_range=>optimize_array_if_range( lookup_array_result ).
*          IF optimized_lookup_array IS NOT INITIAL.
*            DATA(row_number) = optimized_lookup_array-top_left-row.
*            WHILE row_number <= optimized_lookup_array-bottom_right-row.
*              DATA(column_number) = optimized_lookup_array-top_left-column.
*              WHILE column_number <= optimized_lookup_array-bottom_right-column.
*                DATA(cell_value) = lookup_array_result->get_cell_value( column = column_number
*                                                                        row    = row_number ).
*                DATA(equal) = abap_false.
*                DATA(lookup_value_result_2) = SWITCH #( lookup_value_result->type
*                                                        WHEN lif_xlom__va=>c_type-array
*                                                          OR lif_xlom__va=>c_type-range
*                                                        THEN CAST lif_xlom__va( CAST lif_xlom__va_array( lookup_value_result )->get_cell_value(
*                                                                                       column = 1
*                                                                                       row    = 1 ) )
*                                                        ELSE lookup_value_result ).
*                IF    lookup_value_result_2->type = lif_xlom__va=>c_type-string
*                   OR cell_value->type            = lif_xlom__va=>c_type-string.
*                  equal = xsdbool( lcl_xlom__va=>to_string( lookup_value_result_2 )->get_string( )
*                                   = lcl_xlom__va=>to_string( cell_value )->get_string( ) ).
*                ELSEIF    lookup_value_result_2->type = lif_xlom__va=>c_type-number
*                       OR cell_value->type            = lif_xlom__va=>c_type-number.
*                  equal = xsdbool( lcl_xlom__va=>to_number( lookup_value_result_2 )->get_number( )
*                                   = lcl_xlom__va=>to_number( cell_value )->get_number( ) ).
*                ELSEIF    lookup_value_result_2->type = lif_xlom__va=>c_type-boolean
*                       OR cell_value->type            = lif_xlom__va=>c_type-boolean.
*                  equal = xsdbool( lcl_xlom__va=>to_boolean( lookup_value_result_2 )->boolean_value
*                                   = lcl_xlom__va=>to_boolean( cell_value )->boolean_value ).
*                ELSEIF     lookup_value_result_2->type = lif_xlom__va=>c_type-empty
*                       AND cell_value->type            = lif_xlom__va=>c_type-empty.
*                  equal = abap_true.
*                ELSE.
*                  RAISE EXCEPTION TYPE lcx_xlom_todo.
*                ENDIF.
*                IF equal = abap_true.
*                  IF lookup_array_result->row_count > 1.
*                    temp_result = lcl_xlom__va_number=>create( EXACT #( row_number ) ).
*                  ELSE.
*                    temp_result = lcl_xlom__va_number=>create( EXACT #( column_number ) ).
*                  ENDIF.
*                  " Dummy code to exit the two loops
*                  row_number = lookup_array_result->row_count.
*                  column_number = lookup_array_result->column_count.
*                ENDIF.
*                column_number = column_number + 1.
*              ENDWHILE.
*              row_number = row_number + 1.
*            ENDWHILE.
*          ENDIF.
*          IF temp_result IS NOT BOUND.
*            " no match found
*            temp_result = lcl_xlom__va_error=>na_not_applicable.
*          ENDIF.
*        ENDIF.
*        result = lif_xlom__ex~set_result( temp_result ).
*      CATCH lcx_xlom__va INTO DATA(error).
*        result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_offset IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_offset( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-offset.
*    result->reference             = reference.
*    result->rows                  = rows.
*    result->cols                  = cols.
*    result->height                = height.
*    result->width                 = width.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'REFERENCE' object = reference not_part_of_result_array = abap_true )
*                                                       ( name = 'ROWS     ' object = rows      )
*                                                       ( name = 'COLS     ' object = cols      )
*                                                       ( name = 'HEIGHT   ' object = height    )
*                                                       ( name = 'WIDTH    ' object = width     ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*        DATA(rows_result) = lcl_xlom__va=>to_number( arguments[ name = 'ROWS' ]-object )->get_integer( ).
*        DATA(cols_result) = lcl_xlom__va=>to_number( arguments[ name = 'COLS' ]-object )->get_integer( ).
*        DATA(reference_result) = lcl_xlom__va=>to_range( input = arguments[ name = 'REFERENCE' ]-object ).
*        DATA(height_result) = COND #( WHEN height       IS BOUND
*                                       AND height->type <> height->c_type-empty_argument
*                                      THEN lcl_xlom__va=>to_number( arguments[
*                                                                                     name = 'HEIGHT' ]-object )->get_integer( )
*                                      ELSE reference_result->rows( )->count( ) ).
*        DATA(width_result) = COND #( WHEN width       IS BOUND
*                                      AND width->type <> width->c_type-empty_argument
*                                     THEN lcl_xlom__va=>to_number( arguments[
*                                                                                    name = 'WIDTH' ]-object )->get_integer( )
*                                     ELSE reference_result->columns( )->count( ) ).
*        result = lif_xlom__ex~set_result( reference_result->offset( row_offset    = rows_result
*                                                                     column_offset = cols_result
*                                                             )->resize( row_size    = height_result
*                                                                        column_size = width_result ) ).
*      CATCH lcx_xlom__va INTO DATA(error).
*        result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF expression->type <> lif_xlom__ex=>c_type-function-offset.
*      RETURN.
*    ENDIF.
*    DATA(compare_offset) = CAST lcl_xlom__ex_fu_offset( expression ).
*
*    result = xsdbool(     reference->is_equal( compare_offset->reference )
*                      AND rows->is_equal( compare_offset->rows )
*                      AND cols->is_equal( compare_offset->cols )
*                      AND lcl_xlom__ex_ut=>are_equal( expression_1 = height
*                                                        expression_2 = compare_offset->height )
*                      AND lcl_xlom__ex_ut=>are_equal( expression_1 = width
*                                                        expression_2 = compare_offset->width ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_right IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_right( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-right.
*    result->text                  = text.
*    result->num_chars             = num_chars.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'TEXT'      object = text )
*                                                       ( name = 'NUM_CHARS' object = num_chars ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    DATA right       TYPE string.
*    DATA temp_result TYPE REF TO lif_xlom__va.
*
*    TRY.
*    DATA(text) = lcl_xlom__va=>to_string( arguments[ name = 'TEXT' ]-object )->get_string( ).
*    DATA(result_num_chars) = arguments[ name = 'NUM_CHARS' ]-object.
*    DATA(number_num_chars) = COND #( WHEN result_num_chars IS BOUND
*                                          AND result_num_chars->type <> result_num_chars->c_type-empty
*                                     THEN lcl_xlom__va=>to_number( result_num_chars )->get_number( ) ).
*    IF number_num_chars < 0.
*      temp_result = lcl_xlom__va_error=>value_cannot_be_calculated.
*    ELSE.
*      IF text = ''.
*        right = ``.
*      ELSE.
*        DATA(off) = COND i( " Get the last character
*                            WHEN result_num_chars IS NOT BOUND     THEN strlen( text ) - 1
*                            " Get the whole text
*                            WHEN number_num_chars > strlen( text ) THEN 0
*                            " Get exactly as many characters as defined in NUM_CHARS
*                            " (note that if NUM_CHARS = STRLEN( text ), the result is the empty string "")
*                            ELSE                                        strlen( text ) - number_num_chars ).
*        right = substring( val = text
*                           off = off ).
*      ENDIF.
*      temp_result = lcl_xlom__va_string=>create( right ).
*    ENDIF.
*    result = lif_xlom__ex~set_result( temp_result ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF expression->type <> lif_xlom__ex=>c_type-function-right.
*      RETURN.
*    ENDIF.
*    DATA(compare_right) = CAST lcl_xlom__ex_fu_right( expression ).
*
*    result = xsdbool(     text->is_equal( compare_right->text )
*                      AND lcl_xlom__ex_ut=>are_equal( expression_1 = num_chars
*                                                        expression_2 = compare_right->num_chars ) ).
**                      AND (    (     num_chars                IS NOT BOUND
**                                 AND  IS NOT BOUND )
**                            OR (     num_chars                IS BOUND
**                                 AND compare_right->num_chars IS BOUND
**                                 AND num_chars->is_equal( compare_right->num_chars ) ) ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_row IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_row( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-row.
*    result->reference             = reference.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'REFERENCE' object = reference ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    DATA temp_result TYPE REF TO lif_xlom__va.
*
*    IF reference IS NOT BOUND.
*      temp_result = lcl_xlom__va_number=>create( EXACT #( context->containing_cell-row ) ).
*    ELSE.
*      DATA(reference_result) = CAST lcl_xlom_range( arguments[ name = 'REFERENCE' ]-object ).
*      temp_result = lcl_xlom__va_number=>create( reference_result->row( ) ).
*    ENDIF.
*    result = lif_xlom__ex~set_result( temp_result ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_fu_t IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_fu_t( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-function-t.
*    result->value                 = value.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'VALUE' object = value ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    DATA(value_result) = arguments[ name = 'VALUE' ]-object.
*    result = lif_xlom__ex~set_result(
*                 lcl_xlom__va_string=>create( COND #( WHEN value_result->type = value_result->c_type-string
*                                                            THEN lcl_xlom__va=>to_string( value_result )->get_string( )
*                                                            ELSE '' ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_number IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_el_number( ).
*    result->number = number.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-number.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    result = lif_xlom__ex~set_result( lcl_xlom__va_number=>create( number ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression->type = lif_xlom__ex=>c_type-number
*       AND number           = CAST lcl_xlom__ex_el_number( expression )->number.
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_ampersand IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_op_ampersand( ).
*    result->left_operand  = left_operand.
*    result->right_operand = right_operand.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-operation-ampersand.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'LEFT'  object = left_operand )
*                                                       ( name = 'RIGHT' object = right_operand ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    result = lif_xlom__ex~set_result(
*                 lcl_xlom__va_string=>create(
*                     lcl_xlom__va=>to_string( arguments[ name = 'LEFT' ]-object )->get_string( )
*                  && lcl_xlom__va=>to_string( arguments[ name = 'RIGHT' ]-object )->get_string( ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression       IS BOUND
*       AND expression->type  = lif_xlom__ex=>c_type-operation-ampersand
*       AND left_operand->is_equal( CAST lcl_xlom__ex_op_ampersand( expression )->left_operand )
*       AND right_operand->is_equal( CAST lcl_xlom__ex_op_ampersand( expression )->right_operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_colon IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_op_colon( ).
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-operation-plus.
*    result->left_operand  = left_operand.
*    result->right_operand = right_operand.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(cell1) = SWITCH #( left_operand->type
*                            WHEN left_operand->c_type-number THEN
*                              |{ CAST lcl_xlom__va_number( left_operand->evaluate( context ) )->get_integer( ) }|
*                            WHEN left_operand->c_type-array
*                              OR left_operand->c_type-range THEN
*                              CAST lcl_xlom__ex_el_range( left_operand )->_address_or_name
*                            ELSE
*                              THROW lcx_xlom_todo( ) ).
*    DATA(cell2) = SWITCH #( right_operand->type
*                            WHEN right_operand->c_type-number THEN
*                              |{ CAST lcl_xlom__va_number( right_operand->evaluate( context ) )->get_integer( ) }|
*                            WHEN left_operand->c_type-array
*                              OR left_operand->c_type-range THEN
*                              CAST lcl_xlom__ex_el_range( right_operand )->_address_or_name
*                            ELSE
*                              THROW lcx_xlom_todo( ) ).
*    TRY.
*    result = lif_xlom__ex~set_result(
*                 lcl_xlom_range=>create( lcl_xlom_range=>create_from_address_or_name(
*                                                address     = |{ cell1 }:{ cell2 }|
*                                                relative_to = context->worksheet ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression IS BOUND
*       AND expression->type = lif_xlom__ex=>c_type-operation-plus
*       AND left_operand->is_equal( CAST lcl_xlom__ex_op_colon( expression )->left_operand )
*       AND right_operand->is_equal( CAST lcl_xlom__ex_op_colon( expression )->right_operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_equal IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_op_equal( ).
*    result->left_operand  = left_operand.
*    result->right_operand = right_operand.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-operation-equal.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'LEFT'  object = left_operand )
*                                                       ( name = 'RIGHT' object = right_operand ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    DATA temp_result TYPE REF TO lif_xlom__va.
*    DATA(left_result) = arguments[ name = 'LEFT' ]-object.
*    DATA(right_result) = arguments[ name = 'RIGHT' ]-object.
*    IF left_result->type <> right_result->type.
*      temp_result = lcl_xlom__va_boolean=>false.
*    ELSE.
*      DATA(ref_to_left_operand_value) = left_result->get_value( ).
*      DATA(ref_to_right_operand_value) = right_result->get_value( ).
*      ASSIGN ref_to_left_operand_value->* TO FIELD-SYMBOL(<left_operand_value>).
*      ASSIGN ref_to_right_operand_value->* TO FIELD-SYMBOL(<right_operand_value>).
*      temp_result = lcl_xlom__va_boolean=>get( xsdbool( <left_operand_value> = <right_operand_value> ) ).
*    ENDIF.
*    result = lif_xlom__ex~set_result( temp_result ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression       IS BOUND
*       AND expression->type  = lif_xlom__ex=>c_type-operation-equal
*       AND left_operand->is_equal( CAST lcl_xlom__ex_op_equal( expression )->left_operand )
*       AND right_operand->is_equal( CAST lcl_xlom__ex_op_equal( expression )->right_operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_minus IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_op_minus( ).
*    result->left_operand          = left_operand.
*    result->right_operand         = right_operand.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-operation-minus.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'LEFT'  object = left_operand )
*                                                       ( name = 'RIGHT' object = right_operand ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    result = lif_xlom__ex~set_result(
*                 lcl_xlom__va_number=>create(
*                     lcl_xlom__va=>to_number( arguments[ name = 'LEFT' ]-object )->get_number( )
*                   - lcl_xlom__va=>to_number( arguments[ name = 'RIGHT' ]-object )->get_number( ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression IS BOUND
*       AND expression->type = lif_xlom__ex=>c_type-operation-minus
*       AND left_operand->is_equal( CAST lcl_xlom__ex_op_minus( expression )->left_operand )
*       AND right_operand->is_equal( CAST lcl_xlom__ex_op_minus( expression )->right_operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_minus_unry IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_op_minus_unry( ).
*    result->operand               = operand.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-operation-minus_unary.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'OPERAND' object = operand ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    DATA(operand_result) = arguments[ name = 'OPERAND' ]-object.
*    result = lif_xlom__ex~set_result(
*                 lcl_xlom__va_number=>create(
*                   - ( lcl_xlom__va=>to_number( operand_result )->get_number( ) ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression       IS BOUND
*       AND expression->type  = lif_xlom__ex=>c_type-operation-minus_unary
*       AND operand->is_equal( CAST lcl_xlom__ex_op_minus_unry( expression )->operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_mult IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_op_mult( ).
*    result->left_operand          = left_operand.
*    result->right_operand         = right_operand.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-operation-mult.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'LEFT'  object = left_operand )
*                                                       ( name = 'RIGHT' object = right_operand ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    result = lif_xlom__ex~set_result(
*                 lcl_xlom__va_number=>create(
*                     lcl_xlom__va=>to_number( arguments[ name = 'LEFT' ]-object )->get_number( )
*                   * lcl_xlom__va=>to_number( arguments[ name = 'RIGHT' ]-object )->get_number( ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression       IS BOUND
*       AND expression->type  = lif_xlom__ex=>c_type-operation-mult
*       AND left_operand->is_equal( CAST lcl_xlom__ex_op_mult( expression )->left_operand )
*       AND right_operand->is_equal( CAST lcl_xlom__ex_op_mult( expression )->right_operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_op_plus IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_op_plus( ).
*    result->left_operand          = left_operand.
*    result->right_operand         = right_operand.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-operation-plus.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    DATA(array_evaluation) = lcl_xlom__ex_ut=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = VALUE #( ( name = 'LEFT'  object = left_operand )
*                                                       ( name = 'RIGHT' object = right_operand ) ) ).
*    IF array_evaluation-result IS BOUND.
*      result = array_evaluation-result.
*    ELSE.
*      result = lif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
*                                                 context   = context ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    TRY.
*    result = lif_xlom__ex~set_result(
*                 lcl_xlom__va_number=>create(
*                     lcl_xlom__va=>to_number( arguments[ name = 'LEFT' ]-object )->get_number( )
*                   + lcl_xlom__va=>to_number( arguments[ name = 'RIGHT' ]-object )->get_number( ) ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF     expression IS BOUND
*       AND expression->type = lif_xlom__ex=>c_type-operation-plus
*       AND left_operand->is_equal( CAST lcl_xlom__ex_op_plus( expression )->left_operand )
*       AND right_operand->is_equal( CAST lcl_xlom__ex_op_plus( expression )->right_operand ).
*      result = abap_true.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_range IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_el_range( ).
*    result->_address_or_name = address_or_name.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-range. " 16/10
**    result->lif_xlom_expr~type = lif_xlom_expr=>c_type-array. " 16/10
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    TRY.
*    result = lif_xlom__ex~set_result( lcl_xlom_range=>create_from_address_or_name(
*                                              address     = _address_or_name
*                                              relative_to = context->worksheet ) ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      result = error->result_error.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF expression->type <> lif_xlom__ex=>c_type-range.
*      RETURN.
*    ENDIF.
*    DATA(compare_range) = CAST lcl_xlom__ex_el_range( expression ).
*    result = xsdbool( _address_or_name = compare_range->_address_or_name ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__ex_el_string IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__ex_el_string( ).
*    result->string = text.
*    result->lif_xlom__ex~type = lif_xlom__ex=>c_type-string.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate.
*    result = lif_xlom__ex~set_result( lcl_xlom__va_string=>create( string ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~evaluate_single.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~is_equal.
*    IF expression->type <> lif_xlom__ex=>c_type-string.
*      RETURN.
*    ENDIF.
*    DATA(string_object) = CAST lcl_xlom__ex_el_string( expression ).
*    result = xsdbool( string = string_object->string ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__ex~set_result.
*    lif_xlom__ex~result_of_evaluation = value.
*    result = value.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom_range IMPLEMENTATION.
*  METHOD address.
*    IF reference_style <> lcl_xlom=>c_reference_style-a1.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    IF external = abap_true.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    IF relative_to IS BOUND.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*
*    IF _address-top_left-column = 1
*        AND _address-bottom_right-column = lcl_xlom_worksheet=>max_columns.
*      " Whole rows (e.g. "$1:$1")
*      result = |${ _address-top_left-row }:${ _address-bottom_right-row }|.
*    ELSEIF _address-top_left-row = 1
*        AND _address-bottom_right-row = lcl_xlom_worksheet=>max_rows.
*      " Whole columns (e.g. "$A:$A")
*      result = |${ lcl_xlom_range=>convert_column_number_to_a_xfd( _address-top_left-column )
*                }:${ lcl_xlom_range=>convert_column_number_to_a_xfd( _address-bottom_right-column ) }|.
*    ELSE.
*      " one cell (e.g. "$A$1") or several cells (e.g. "$A$1:$A$2")
*      result = |${ lcl_xlom_range=>convert_column_number_to_a_xfd( _address-top_left-column )
*               }${ _address-top_left-row
*               }{ COND #( WHEN _address-bottom_right <> _address-top_left
*                          THEN |:${ lcl_xlom_range=>convert_column_number_to_a_xfd( _address-bottom_right-column )
*                                }${ _address-bottom_right-row }| ) }|.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD calculate.
*    IF _address-top_left <> _address-bottom_right.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    DATA(cell) = REF #( parent->_array->_cells[ column = _address-top_left-column
*                                                row    = _address-top_left-row ] OPTIONAL ).
*    IF cell IS NOT BOUND.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    DATA(context) = lcl_xlom__ut_eval_context=>create(
*                        worksheet       = parent
*                        containing_cell = VALUE #( row    = _address-top_left-row
*                                                   column = _address-top_left-column ) ).
*    cell->formula->evaluate( context ).
*  ENDMETHOD.
*
*  METHOD cells.
*    " This will change Z20:
*    " Range("Z20:AA25").Cells(1, 1) = "C"
*    result = lcl_xlom_range=>create_from_row_column( worksheet = parent
*                                                        row       = _address-top_left-row + row - 1
*                                                        column    = _address-top_left-column + column - 1 ).
*  ENDMETHOD.
*
*  METHOD columns.
*    result = lcl_xlom_range=>create_from_top_left_bottom_ri(
*                 worksheet             = parent
*                 top_left              = _address-top_left
*                 bottom_right          = _address-bottom_right
*                 column_row_collection = c_column_row_collection-columns ).
*  ENDMETHOD.
*
*  METHOD convert_column_a_xfd_to_number.
*    DATA(offset) = 0.
*    WHILE offset < strlen( roman_letters ).
*      FIND roman_letters+offset(1) IN sy-abcde MATCH OFFSET DATA(offset_a_to_z).
*      IF sy-subrc <> 0.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      ENDIF.
*      result = ( result * 26 ) + offset_a_to_z + 1.
*      IF result > lcl_xlom_worksheet=>max_columns.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      ENDIF.
*      offset = offset + 1.
*    ENDWHILE.
*  ENDMETHOD.
*
*  METHOD convert_column_number_to_a_xfd.
*    IF number NOT BETWEEN 1 AND lcl_xlom_worksheet=>max_columns.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    DATA(work_number) = number.
*    DO.
*      DATA(lv_mod) = ( work_number - 1 ) MOD 26.
*      DATA(lv_div) = ( work_number - 1 ) DIV 26.
*      work_number = lv_div.
*      result = sy-abcde+lv_mod(1) && result.
*      IF work_number <= 0.
*        EXIT.
*      ENDIF.
*    ENDDO.
*  ENDMETHOD.
*
*  METHOD count.
*    result = lif_xlom__va_array~column_count * lif_xlom__va_array~row_count.
*  ENDMETHOD.
*
*  METHOD create.
*    IF cell1 IS NOT BOUND.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    IF cell2 IS NOT BOUND.
*      result = cell1.
*      RETURN.
*    ENDIF.
*    IF cell1->parent <> cell2->parent.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    " This will set "g" from Z20 to AD24:
*    " Range(Range("Z20:AA21"), Range("AC23:AD24")).Value = "g"
*    " This will set "H" from Z20 to AD25:
*    " Range(Range("Z20:AA25"), Range("AC23:AD24")).Value = "H"
*    " This will set "H" from Z20 to AD25:
*    " Range(Range("AA25:Z20"), Range("AD24:AC23")).Value = "H"
*    DATA(structured_address) = VALUE lcl_xlom=>ts_range_address(
**    DATA(structured_address) = VALUE lif_xlom_result_array=>ts_address(
*        top_left     = VALUE #( column = nmin( val1 = nmin( val1 = cell1->_address-top_left-column
*                                                            val2 = cell2->_address-top_left-column )
*                                               val2 = nmin( val1 = cell1->_address-bottom_right-column
*                                                            val2 = cell2->_address-bottom_right-column ) )
*                                row    = nmin( val1 = nmin( val1 = cell1->_address-top_left-row
*                                                            val2 = cell2->_address-top_left-row )
*                                               val2 = nmin( val1 = cell1->_address-bottom_right-row
*                                                            val2 = cell2->_address-bottom_right-row ) ) )
*        bottom_right = VALUE #( column = nmax( val1 = nmax( val1 = cell1->_address-top_left-column
*                                                            val2 = cell2->_address-top_left-column )
*                                               val2 = nmax( val1 = cell1->_address-bottom_right-column
*                                                            val2 = cell2->_address-bottom_right-column ) )
*                                row    = nmax( val1 = nmax( val1 = cell1->_address-top_left-row
*                                                            val2 = cell2->_address-top_left-row )
*                                               val2 = nmax( val1 = cell1->_address-bottom_right-row
*                                                            val2 = cell2->_address-bottom_right-row ) ) ) ).
*    result = create_from_top_left_bottom_ri( worksheet    = cell1->parent
*                                             top_left     = structured_address-top_left
*                                             bottom_right = structured_address-bottom_right ).
*  ENDMETHOD.
*
*  METHOD create_from_address_or_name.
**    DATA(structured_address) = decode_range_address( address     = address
**                                                     relative_to = relative_to ).
*    DATA(structured_address) = decode_range_address( address ).
**    DATA(structured_address) = relative_to->decode_range_address( address ).
*    result = create_from_top_left_bottom_ri( worksheet    = cond #( when structured_address-worksheet_name is initial
*                                                                    then relative_to
*                                                                    else relative_to->parent->worksheets->item( structured_address-worksheet_name ) )
*                                             top_left     = VALUE #( column = structured_address-top_left-column
*                                                                     row    = structured_address-top_left-row )
*                                             bottom_right = VALUE #( column = structured_address-bottom_right-column
*                                                                     row    = structured_address-bottom_right-row ) ).
*  ENDMETHOD.
*
*  METHOD create_from_expr_range.
*    result = create_from_address_or_name( address     = expr_range->_address_or_name
*                                          relative_to = relative_to ).
*  ENDMETHOD.
*
*  METHOD create_from_row_column.
*    result = create_from_top_left_bottom_ri( worksheet    = worksheet
*                                             top_left     = VALUE #( column = column
*                                                                     row    = row )
*                                             bottom_right = VALUE #( column = column + column_size - 1
*                                                                     row    = row + row_size - 1 ) ).
*  ENDMETHOD.
*
*  METHOD create_from_top_left_bottom_ri.
*    DATA range TYPE REF TO lcl_xlom_range.
*
*    " If row = 0, it means the range is a whole column (rows from 1 to 1048576).
*    " If column = 0, it means the range is a whole row (columns from 1 to 16384).
*    DATA(address) = VALUE lcl_xlom=>ts_range_address(
**    DATA(address) = VALUE lif_xlom_result_array=>ts_address(
*        top_left     = VALUE #( row    = COND #( WHEN top_left-row > 0
*                                                 THEN top_left-row
*                                                 ELSE 1 )
*                                column = COND #( WHEN top_left-column > 0
*                                                 THEN top_left-column
*                                                 ELSE 1 ) )
*        bottom_right = VALUE #( row    = COND #( WHEN bottom_right-row > 0
*                                                 THEN bottom_right-row
*                                                 ELSE lcl_xlom_worksheet=>max_rows )
*                                column = COND #( WHEN bottom_right-column > 0
*                                                 THEN bottom_right-column
*                                                 ELSE lcl_xlom_worksheet=>max_columns ) ) ).
*    DATA(range_buffer_line) = REF #( _range_buffer[ worksheet             = worksheet
*                                                    address               = address
*                                                    column_row_collection = column_row_collection ] OPTIONAL ).
*    IF range_buffer_line IS NOT BOUND.
*      CASE column_row_collection.
*        WHEN c_column_row_collection-columns.
*          range = NEW lcl_xlom_columns( ).
*        WHEN c_column_row_collection-rows.
*          range = NEW lcl_xlom_rows( ).
*        WHEN c_column_row_collection-none.
*          range = NEW lcl_xlom_range( ).
*        WHEN OTHERS.
*          RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*      ENDCASE.
*      range->lif_xlom__va~type               = lif_xlom__va=>c_type-range.
*      range->application                        = worksheet->application.
*      range->parent                             = worksheet.
*      range->_address                           = address.
*      range->lif_xlom__va_array~row_count    = address-bottom_right-row - address-top_left-row + 1.
*      range->lif_xlom__va_array~column_count = address-bottom_right-column - address-top_left-column + 1.
*      INSERT VALUE #( worksheet             = worksheet
*                      address               = address
*                      column_row_collection = column_row_collection
*                      object                = range )
*             INTO TABLE _range_buffer
*             REFERENCE INTO range_buffer_line.
*    ENDIF.
*    result = range_buffer_line->object.
*  ENDMETHOD.
*
*  METHOD decode_range_address.
*    " The range address should always be in A1 reference style.
**    IF relative_to->parent->application->reference_style = lcl_xlom=>c_reference_style-a1.
*      result = lcl_xlom_range=>decode_range_address_a1( address ).
**      IF decoded_range-worksheet_name IS INITIAL.
**        result-worksheet_name = relative_to->parent->name.
**      ENDIF.
**    ELSE.
**      RAISE EXCEPTION TYPE lcx_xlom_todo.
***      result = lcl_xlom_range=>decode_range_address_r1_c1( address ).
**    ENDIF.
*    IF result-top_left IS INITIAL.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
**     " address is an invalid range address so it's probably referring to an existing name.
**     " Find the name in the current worksheet
**     result-name = parent->parent->names[ worksheet = parent ].
**     " Find the name in the current workbook
**     result-name = parent->parent->names[ worksheet = parent ].
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD decode_range_address_a1.
*    " Special characters are ":", "$", "!", "'", "[" and "]".
*    " When "'" is found, all characters till the next "'" form
*    "   one word or two words if there are "[" and "]".
*    " When "[" is found, all characters till the next "]" form
*    "   the workbook name.
*    " Subsequent non-special characters form a word.
*    "
*    " Examples:
*    "   In the current worksheet:
*    "     NB: they are case-insensitive
*    "     A1 (relative column and row)             word
*    "     $A1 (absolute column, relative column)   $ word
*    "     A$1                                      word $ word
*    "     $A$1                                     $ word $ word
*    "     A1:A2                                    word : word
*    "     $A$A                                     $ word $ word
*    "     A:A                                      word : word
*    "     1:1                                      word : word
*    "     NAME                                     word
*    "   Other worksheet:
*    "     Sheet1!A1                                word ! word
*    "     'Sheet 1'!A1                             word ! word
*    "     [1]Sheet1!$A$3                           [word] word ! $ word $ word   (XLSX internal notation for workbooks)
*    "   Other workbook:
*    "     '[C:\workbook.xlsx]'!NAME                [word] ! word                 (workbook absolute path / name in the global scope)
*    "     '[workbook.xlsx]Sheet 1'!$A$1            [word] word ! word            (workbook relative path)
*    "     [1]!NAME                                 [word] ! word                 (XLSX internal notation for workbooks)
*    TYPES ty_state TYPE i.
*
*    CONSTANTS:
*      BEGIN OF c_state,
*        normal                        TYPE ty_state VALUE 1,
*        within_single_quotes          TYPE ty_state VALUE 2,
*        within_single_quotes_brackets TYPE ty_state VALUE 3,
*        within_brackets               TYPE ty_state VALUE 4,
*      END OF c_state.
*    DATA colon_position TYPE i.
*
*    DATA(words) = VALUE string_table( ).
*    INSERT INITIAL LINE INTO TABLE words REFERENCE INTO DATA(current_word).
*    DATA(state) = c_state-normal.
*    DATA(offset) = 0.
*    WHILE offset < strlen( address ).
*      DATA(character) = substring( val = address
*                                   off = offset
*                                   len = 1 ).
*
*      DATA(start_a_new_word) = abap_false.
*      DATA(store_dedicated_word) = abap_false.
*      CASE state.
*        WHEN c_state-normal.
*          CASE character.
*            WHEN ''''.
*              state = c_state-within_single_quotes.
*            WHEN '$'.
*              store_dedicated_word = abap_true.
*            WHEN '!'.
*              store_dedicated_word = abap_true.
*            WHEN ':'.
*              store_dedicated_word = abap_true.
*            WHEN '['.
*              DATA(square_bracket_position) = 1.
*              state = c_state-within_brackets.
*            WHEN OTHERS.
*              current_word->* = current_word->* && character.
*          ENDCASE.
*        WHEN c_state-within_single_quotes.
*          CASE character.
*            WHEN ''''.
*              start_a_new_word = abap_true.
*              state = c_state-normal.
*            WHEN '['.
*              square_bracket_position = 1.
*              state = c_state-within_single_quotes_brackets.
*            WHEN OTHERS.
*              current_word->* = current_word->* && character.
*          ENDCASE.
*        WHEN c_state-within_single_quotes_brackets.
*          CASE character.
*            WHEN ']'.
*              start_a_new_word = abap_true.
*              state = c_state-within_single_quotes.
*            WHEN OTHERS.
*              current_word->* = current_word->* && character.
*          ENDCASE.
*        WHEN c_state-within_brackets.
*          CASE character.
*            WHEN ']'.
*              start_a_new_word = abap_true.
*              state = c_state-normal.
*            WHEN OTHERS.
*              current_word->* = current_word->* && character.
*          ENDCASE.
*      ENDCASE.
*      IF    start_a_new_word     = abap_true
*         OR store_dedicated_word = abap_true.
*        IF current_word->* IS NOT INITIAL.
*          INSERT INITIAL LINE INTO TABLE words REFERENCE INTO current_word.
*        ENDIF.
*        CASE character.
*          WHEN '!'.
*            DATA(exclamation_mark_position) = lines( words ).
*          WHEN ':'.
*            colon_position = lines( words ).
*        ENDCASE.
*        IF store_dedicated_word = abap_true.
*          current_word->* = character.
*          INSERT INITIAL LINE INTO TABLE words REFERENCE INTO current_word.
*        ENDIF.
*        start_a_new_word = abap_false.
*        store_dedicated_word = abap_false.
*      ENDIF.
*      offset = offset + 1.
*    ENDWHILE.
*
*    IF square_bracket_position = 1.
*      result-workbook_name = words[ 1 ].
*    ENDIF.
*
*    IF    exclamation_mark_position = 3
*       OR (     exclamation_mark_position = 2
*            AND square_bracket_position   = 0 ).
*      result-worksheet_name = words[ exclamation_mark_position - 1 ].
*    ENDIF.
*
*    IF colon_position = 0.
*      IF exclamation_mark_position + 1 = lines( words ).
*        " A1 or NAME
*        result-top_left = decode_range_coords( words = words
*                                               from  = lines( words )
*                                               to    = lines( words ) ).
*        IF result-top_left IS INITIAL.
*          " NAME
*          result-range_name = words[ lines( words ) ].
*        ELSE.
*          result-bottom_right = result-top_left.
*        ENDIF.
*      ELSE.
*        " $A$1, A$1, $A1
*        result-top_left = decode_range_coords( words = words
*                                               from  = exclamation_mark_position + 1
*                                               to    = lines( words ) ).
*        result-bottom_right = result-top_left.
*      ENDIF.
*    ELSE.
*      " A1:A2, $A$1:$B$2, A1:$B$2, etc.
*      result-top_left = decode_range_coords( words = words
*                                             from  = exclamation_mark_position + 1
*                                             to    = colon_position - 1 ).
*      result-bottom_right = decode_range_coords( words = words
*                                                 from  = colon_position + 1
*                                                 to    = lines( words ) ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD decode_range_coords.
*    " Remove $ if any
*    DATA(coords_without_dollar) = REDUCE #( INIT t = ``
*                       FOR <word> IN words FROM from TO to
*                       WHERE ( table_line <> '$' )
*                       NEXT t = t && <word> ).
*
*    DATA(coords) = decode_range_name_or_coords( coords_without_dollar ).
*
*    IF coords IS NOT INITIAL.
*      result = VALUE #( column = coords-column
*                        row    = coords-row ).
*
*      IF words[ from ] = '$'.
*        result-column_fixed = abap_true.
*        IF     from + 2          <= lines( words )
*           AND words[ from + 2 ]  = '$'.
*          result-row_fixed = abap_true.
*        ENDIF.
*      ELSEIF     from              < lines( words )
*             AND words[ from + 1 ] = '$'.
*        result-row_fixed = abap_true.
*      ENDIF.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD decode_range_name_or_coords.
*    DATA(offset) = 0.
*    WHILE     offset < strlen( range_name_or_coords )
*          AND range_name_or_coords+offset(1) NA '123456789'.
*      offset = offset + 1.
*    ENDWHILE.
*    IF     offset <= 3
*       AND range_name_or_coords(offset) CO 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
*       AND (    offset < 3
*             OR (     offset                   = 3
*                  AND range_name_or_coords(3) <= 'XFD' ) ).
*      IF     range_name_or_coords+offset           CO '1234567890'
*         AND CONV i( range_name_or_coords+offset ) <= lcl_xlom_worksheet=>max_rows.
*        result-column = convert_column_a_xfd_to_number( substring( val = range_name_or_coords
*                                                                   len = offset ) ).
*        result-row    = range_name_or_coords+offset.
*      ENDIF.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD formula2.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD offset.
*    result = _offset_resize( row_offset    = row_offset
*                             column_offset = column_offset
*                             row_size      = _address-bottom_right-row - _address-top_left-row + 1
*                             column_size   = _address-bottom_right-column - _address-top_left-column + 1 ).
*  ENDMETHOD.
*
*  METHOD _offset_resize.
*    result = create_from_row_column( worksheet   = parent
*                                     row         = _address-top_left-row + row_offset
*                                     column      = _address-top_left-column + column_offset
*                                     row_size    = row_size
*                                     column_size = column_size ).
*  ENDMETHOD.
*
*  METHOD optimize_array_if_range.
**    DATA(row_count) = 0.
**    DATA(column_count) = 0.
*    IF array->lif_xlom__va~type = array->lif_xlom__va~c_type-range.
*      DATA(range) = CAST lcl_xlom_range( array ).
*      result = range->application->_intersect_2_basis(
*                   arg1 = VALUE #( top_left-column     = range->_address-top_left-column
*                                   top_left-row        = range->_address-top_left-row
*                                   bottom_right-column = range->_address-bottom_right-column
*                                   bottom_right-row    = range->_address-bottom_right-row )
*                   arg2 = range->parent->_array->used_range ).
**        IF optimized_lookup_array is not initial.
**        row_count = optimized_lookup_array-bottom_right-row - optimized_lookup_array-bottom_right-row + 1.
**        column_count = optimized_lookup_array-bottom_right-column - optimized_lookup_array-bottom_right-column + 1.
**        endif.
**        DATA(worksheet_used_range) = CAST lcl_xlom_range( lookup_array_result )->parent->used_range( ).
**        row_count = nmin( val1 = lookup_array_result->row_count
**                          val2 = worksheet_used_range->rows( )->count( ) ).
**        column_count = nmin( val1 = lookup_array_result->column_count
**                             val2 = worksheet_used_range->columns( )->count( ) ).
*    ELSE.
*      result = VALUE #( top_left-column     = 1
*                        top_left-row        = 1
*                        bottom_right-column = array->column_count
*                        bottom_right-row    = array->row_count ).
**        row_count = lookup_array_result->row_count.
**        column_count = lookup_array_result->column_count.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD resize.
*    result = _offset_resize( row_offset    = 0
*                             column_offset = 0
*                             row_size      = row_size
*                             column_size   = column_size ).
*  ENDMETHOD.
*
*  METHOD row.
*    result = _address-top_left-row.
*  ENDMETHOD.
*
*  METHOD rows.
*    result = lcl_xlom_range=>create_from_top_left_bottom_ri(
*                 worksheet             = parent
*                 top_left              = _address-top_left
*                 bottom_right          = _address-bottom_right
*                 column_row_collection = c_column_row_collection-rows ).
*  ENDMETHOD.
*
*  METHOD set_formula2.
*    DATA(formula_buffer_line) = REF #( _formula_buffer[ formula = value ] OPTIONAL ).
*    IF formula_buffer_line IS NOT BOUND.
*      DATA(lexer) = lcl_xlom__ex_ut_lexer=>create( ).
*      DATA(lexer_tokens) = lexer->lexe( value ).
*      DATA(parser) = lcl_xlom__ex_ut_parser=>create( ).
*      INSERT VALUE #( formula = value
*                      object  = parser->parse( lexer_tokens ) )
*             INTO TABLE _formula_buffer
*             REFERENCE INTO formula_buffer_line.
*    ENDIF.
*    DATA(formula_expression) = formula_buffer_line->object.
*
*    IF application->calculation = lcl_xlom=>c_calculation-automatic.
*      parent->_array->lif_xlom__va_array~set_cell_value( row        = _address-top_left-row
*                                                               column     = _address-top_left-column
*                                                               formula    = formula_expression
*                                                               calculated = abap_true
*                                                               value      = formula_expression->evaluate(
*                                                                   context = lcl_xlom__ut_eval_context=>create(
*                                                                       worksheet       = parent
*                                                                       containing_cell = VALUE #(
*                                                                           row    = _address-top_left-row
*                                                                           column = _address-top_left-column ) ) ) ).
*    ELSE.
*      parent->_array->lif_xlom__va_array~set_cell_value( row        = _address-top_left-row
*                                                               column     = _address-top_left-column
*                                                               value      = lcl_xlom__va_number=>get( 0 )
*                                                               formula    = formula_expression
*                                                               calculated = abap_false ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD set_value.
*    parent->_array->lif_xlom__va_array~set_cell_value( row    = _address-top_left-row
*                                                             column = _address-top_left-column
*                                                             value  = value ).
*  ENDMETHOD.
*
*  METHOD value.
*    IF _address-top_left = _address-bottom_right.
*      result = parent->_array->lif_xlom__va_array~get_cell_value( column = _address-top_left-column
*                                                                     row    = _address-top_left-row ).
*    ELSE.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
**      IF     lif_xlom_result_array~row_count    > 1
**         AND lif_xlom_result_array~column_count > 1.
**        result = lcl_xlom__va_itab_2=>create( VALUE #( ( row = 1 column  )  ) ).
**      ELSEIF lif_xlom_result_array~row_count > 1.
**        result = lcl_xlom__va_itab_1=>create( ).
**      ELSE.
**        result = lcl_xlom__va_itab_1=>create( ).
**      ENDIF.
**      result = parent->_array->lif_xlom_result_array~get_array_value( top_left     = _address-top_left
**                                                                      bottom_right = _address-bottom_right ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~get_value.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_array.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_boolean.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_equal.
*    IF     input_result->type = lif_xlom__va=>c_type-range.
*      DATA(input_range) = CAST lcl_xlom_range( input_result ).
*      IF me->_address = input_range->_address.
*        result = abap_true.
*      ELSE.
*        result = abap_false.
*      ENDIF.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_error.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_number.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_string.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~get_array_value.
*    result = _offset_resize( row_offset    = top_left-row - 1
*                             column_offset = top_left-column - 1
*                             row_size      = bottom_right-row - top_left-row + 1
*                             column_size   = bottom_right-column - top_left-column + 1 ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~get_cell_value.
*    " if the current range starts from row 2 column 2 and the requested cell is row 2 column 2
*    " then get the worksheet cell from row 3 column 3.
*    result = parent->_array->lif_xlom__va_array~get_cell_value( column = _address-top_left-column + column - 1
*                                                                      row    = _address-top_left-row + row - 1 ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~set_array_value.
*    parent->_array->lif_xlom__va_array~set_array_value( rows = rows ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~set_cell_value.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_array IMPLEMENTATION.
*  METHOD class_constructor.
*    initial_used_range = VALUE #( top_left     = VALUE #( row    = 1
*                                                          column = 1 )
*                                  bottom_right = VALUE #( row    = 1
*                                                          column = 1 ) ).
*  ENDMETHOD.
*
*  METHOD create_from_range.
*    result = range->parent->_array.
*  ENDMETHOD.
*
*  METHOD create_initial.
*    result = NEW lcl_xlom__va_array( ).
*    result->lif_xlom__va~type = lif_xlom__va=>c_type-array.
*    result->lif_xlom__va_array~row_count    = row_count.
*    result->lif_xlom__va_array~column_count = column_count.
*    result->used_range = initial_used_range.
*    result->lif_xlom__va_array~set_array_value( rows ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~get_array_value.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*    DATA(row_count)    = bottom_right-row - top_left-row + 1.
*    DATA(column_count) = bottom_right-column - top_left-column + 1.
*    DATA(target_array) = create_initial( row_count    = row_count
*                                         column_count = column_count ).
*    DATA(row) = 1.
*    WHILE row <= row_count.
*      DATA(column) = 1.
*      WHILE column <= column_count.
*        target_array->lif_xlom__va_array~set_cell_value(
*            column = column
*            row    = row
*            value  = lif_xlom__va_array~get_cell_value( column = top_left-column + column - 1
*                                                              row    = top_left-row + row - 1 ) ).
*        column = column + 1.
*      ENDWHILE.
*      row = row + 1.
*    ENDWHILE.
*    result = target_array.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~get_cell_value.
*    IF    row    < used_range-top_left-row
*       OR row    > used_range-bottom_right-row
*       OR column < used_range-top_left-column
*       OR column > used_range-bottom_right-column.
*      result = lcl_xlom__va_empty=>get_singleton( ).
*    ELSE.
*      DATA(cell) = REF #( _cells[ row    = row
*                                  column = column ] OPTIONAL ).
*      IF cell IS NOT BOUND.
*        " Empty/Blank - Its evaluation depends on its usage (zero or empty string)
*        " =1+Empty gives 1, ="a"&Empty gives "a"
*        result = lcl_xlom__va_empty=>get_singleton( ).
*      ELSE.
*        result = cell->value.
*      ENDIF.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~set_array_value.
*    DATA(row) = 1.
*    LOOP AT rows REFERENCE INTO DATA(row2).
*
*      DATA(column) = 1.
*      LOOP AT row2->columns_of_row INTO DATA(column_value).
*        lif_xlom__va_array~set_cell_value( column = column
*                                                 row    = row
*                                                 value  = column_value ).
*        column = column + 1.
*      ENDLOOP.
*
*      row = row + 1.
*    ENDLOOP.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va_array~set_cell_value.
*    IF    row    > lif_xlom__va_array~row_count
*       OR row    < 1
*       OR column > lif_xlom__va_array~column_count
*       OR column < 1.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*    IF value IS NOT BOUND.
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*
*    CASE value->type.
*
*      WHEN value->c_type-array
*        OR value->c_type-range.
*
*        DATA(source_array) = CAST lif_xlom__va_array( value ).
*        DATA(source_array_row) = 1.
*        WHILE source_array_row <= source_array->row_count.
*          DATA(source_array_column) = 1.
*          WHILE source_array_column <= source_array->column_count.
*            DATA(source_array_cell) = source_array->get_cell_value( column = source_array_column
*                                                                    row    = source_array_row ).
*            set_cell_value_single( row        = row + source_array_row - 1
*                                   column     = column + source_array_column - 1
*                                   value      = source_array_cell
*                                   formula    = formula
*                                   calculated = calculated ).
*
*            source_array_column = source_array_column + 1.
*          ENDWHILE.
*          source_array_row = source_array_row + 1.
*        ENDWHILE.
*
*      WHEN OTHERS.
*
*        set_cell_value_single( row        = row
*                               column     = column
*                               value      = value
*                               formula    = formula
*                               calculated = calculated ).
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD set_cell_value_single.
*    DATA(cell) = REF #( _cells[ row    = row
*                                column = column ] OPTIONAL ).
*    IF cell IS NOT BOUND.
*      INSERT VALUE #( row    = row
*                      column = column )
*             INTO TABLE _cells
*             REFERENCE INTO cell.
*    ENDIF.
*    cell->value      = value.
*    cell->formula    = formula.
*    cell->calculated = calculated.
*
*    IF lines( _cells ) = 1.
*      used_range = VALUE #( top_left     = VALUE #( row    = row
*                                                    column = column )
*                            bottom_right = VALUE #( row    = row
*                                                    column = column ) ).
*    ELSE.
*      IF row < used_range-top_left-row.
*        used_range-top_left-row = row.
*      ENDIF.
*      IF column < used_range-top_left-column.
*        used_range-top_left-column = column.
*      ENDIF.
*      IF row > used_range-bottom_right-row.
*        used_range-bottom_right-row = row.
*      ENDIF.
*      IF column > used_range-bottom_right-column.
*        used_range-bottom_right-column = column.
*      ENDIF.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~get_value.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_array.
*    result = abap_true.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_boolean.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_equal.
*    IF     input_result->type = lif_xlom__va=>c_type-array.
*      DATA(input_array) = CAST lcl_xlom__va_array( input_result ).
*      IF     lif_xlom__va_array~column_count = input_array->lif_xlom__va_array~column_count
*         AND lif_xlom__va_array~row_count    = input_array->lif_xlom__va_array~row_count
*         AND me->_cells            = input_array->_cells.
*        result = abap_true.
*      ELSE.
*        result = abap_false.
*      ENDIF.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_error.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_number.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_string.
*    result = abap_false.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_boolean IMPLEMENTATION.
*  METHOD class_constructor.
*    false = create( abap_false ).
*    true  = create( abap_true ).
*  ENDMETHOD.
*
*  METHOD create.
*    result = NEW lcl_xlom__va_boolean( ).
*    result->lif_xlom__va~type = lif_xlom__va=>c_type-boolean.
*    result->boolean_value = boolean_value.
*    result->number = COND #( when boolean_value = abap_true then -1 ).
*  ENDMETHOD.
*
*  METHOD get.
*    result = SWITCH #( boolean_value WHEN abap_true
*                                     THEN true
*                                     ELSE false ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~get_value.
*    result = REF #( boolean_value ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_array.
*    result = abap_true.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_boolean.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_error.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_number.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_string.
*    result = abap_false.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va IMPLEMENTATION.
*  METHOD to_array.
*    CASE input->type.
*      WHEN input->c_type-error.
*        " TODO I didn't check whether it should be #N/A, #REF! or #VALUE!
*        RAISE EXCEPTION TYPE lcx_xlom__va EXPORTING result_error = lcl_xlom__va_error=>value_cannot_be_calculated.
*      WHEN input->c_type-array
*        OR input->c_type-range.
*        result = CAST #( input ).
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD to_boolean.
*    " If source is Number:
*    "       FALSE if 0
*    "       TRUE if not 0
*    "
*    "  If source is String:
*    "       Language-dependent.
*    "       In English:
*    "       TRUE if "TRUE"
*    "       FALSE if "FALSE"
*    CASE input->type.
*      WHEN input->c_type-array.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      WHEN input->c_type-boolean.
*        result = CAST #( input ).
*      WHEN input->c_type-error.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      WHEN input->c_type-number.
*        CASE CAST lcl_xlom__va_number( input )->get_number( ).
*          WHEN 0.
*            result = lcl_xlom__va_boolean=>false.
*          WHEN OTHERS.
*            result = lcl_xlom__va_boolean=>true.
*        ENDCASE.
*      WHEN input->c_type-range.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      WHEN input->c_type-string.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD to_number.
*    " If source is Boolean:
*    "      0 if FALSE
*    "      1 if TRUE
*    "
*    " If source is String:
*    "      Language-dependent for decimal separator.
*    "      "." if English, "," if French, etc.
*    "      Accepted: "-1", "+1", ".5", "1E1", "-.5"
*    "      "-1E-1", "1e1", "1e307", "1e05", "1e-309"
*    "      Invalid: "", "E1", "1e308", "1e-310"
*    "      #VALUE! if invalid decimal separator
*    "      #VALUE! if invalid number
*    IF input->type <> input->c_type-array
*        and input->type <> input->c_type-range.
*      DATA(cell) = input.
*    ELSE.
**      DATA(range) = CAST lcl_xlom_range( input ).
*      DATA(range) = CAST lif_xlom__va_array( input ).
*      IF    range->column_count <> 1
*         OR range->row_count    <> 1.
**      IF range->top_left <> range->_address-bottom_right.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      ENDIF.
*      cell = range->get_cell_value( column = 1
*                                    row    = 1 ).
*    ENDIF.
*
*    CASE cell->type.
*      WHEN cell->c_type-array.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      WHEN cell->c_type-boolean.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      WHEN cell->c_type-empty.
*        result = lcl_xlom__va_number=>get( 0 ).
*      WHEN cell->c_type-error.
*        RAISE EXCEPTION TYPE lcx_xlom__va EXPORTING result_error = CAST #( cell ).
*      WHEN cell->c_type-number.
*        result = CAST #( cell ).
*      WHEN cell->c_type-range.
*        " impossible because processed in the previous block.
*        RAISE EXCEPTION TYPE lcx_xlom_unexpected.
*      WHEN cell->c_type-string.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD to_range.
*    CASE input->type.
*      WHEN input->c_type-error.
*        " TODO I didn't check whether it should be #N/A, #REF! or #VALUE!
*        RAISE EXCEPTION TYPE lcx_xlom__va EXPORTING result_error = lcl_xlom__va_error=>value_cannot_be_calculated.
*      WHEN input->c_type-range.
*        result = CAST #( input ).
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD to_string.
*    CASE input->type.
*      WHEN input->c_type-empty.
*        result = lcl_xlom__va_string=>create( '' ).
*      WHEN input->c_type-error.
*        RAISE EXCEPTION TYPE lcx_xlom__va EXPORTING result_error = CAST #( input ).
*      WHEN input->c_type-number.
*        result = lcl_xlom__va_string=>create( |{ CAST lcl_xlom__va_number( input )->get_number( ) }| ).
*      WHEN input->c_type-range.
*        DATA(range) = CAST lcl_xlom_range( input ).
*        IF range->_address-top_left <> range->_address-bottom_right.
*          RAISE EXCEPTION TYPE lcx_xlom_todo.
*        ENDIF.
*        DATA(cell) = ref #( range->parent->_array->_cells[ row    = range->_address-top_left-row
*                                                           column = range->_address-top_left-column ] OPTIONAL ).
*        DATA(string) = COND string( WHEN cell IS BOUND
*                                    THEN SWITCH #( cell->value->type
*                                                   WHEN lif_xlom__va=>c_type-number THEN
*                                                     |{ CAST lcl_xlom__va_number( cell->value )->get_number( ) }|
*                                                   WHEN lif_xlom__va=>c_type-string THEN
*                                                     CAST lcl_xlom__va_string( cell->value )->get_string( )
*                                                   ELSE
*                                                     THROW lcx_xlom_todo( ) ) ).
*        result = lcl_xlom__va_string=>create( string ).
*      WHEN input->c_type-string.
*        result = CAST #( input ).
*      WHEN OTHERS.
*        RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDCASE.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_empty IMPLEMENTATION.
*  METHOD get_singleton.
*    IF singleton IS NOT BOUND.
*      singleton = NEW lcl_xlom__va_empty( ).
*      singleton->lif_xlom__va~type = lif_xlom__va=>c_type-empty.
*    ENDIF.
*    result = singleton.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~get_value.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_array.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_boolean.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_error.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_number.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_string.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_error IMPLEMENTATION.
*  METHOD class_constructor.
*    blocked                    = lcl_xlom__va_error=>create( english_error_name    = '#BLOCKED!     '
*                                                                   internal_error_number = 2047
*                                                                   formula_error_number  = 11 ).
*    calc                       = lcl_xlom__va_error=>create( english_error_name    = '#CALC!        '
*                                                                   internal_error_number = 2050
*                                                                   formula_error_number  = 14 ).
*    connect                    = lcl_xlom__va_error=>create( english_error_name    = '#CONNECT!     '
*                                                                   internal_error_number = 2046
*                                                                   formula_error_number  = 10 ).
*    division_by_zero           = lcl_xlom__va_error=>create( english_error_name    = '#DIV/0!       '
*                                                                   internal_error_number = 2007
*                                                                   formula_error_number  = 2
*                                                                   description           = 'Is produced by =1/0' ).
*    field                      = lcl_xlom__va_error=>create( english_error_name    = '#FIELD!       '
*                                                                   internal_error_number = 2049
*                                                                   formula_error_number  = 13 ).
*    getting_data               = lcl_xlom__va_error=>create( english_error_name    = '#GETTING_DATA!'
*                                                                   internal_error_number = 2043
*                                                                   formula_error_number  = 8 ).
*    na_not_applicable          = lcl_xlom__va_error=>create( english_error_name    = '#N/A          '
*                                                                   internal_error_number = 2042
*                                                                   formula_error_number  = 7
*                                                                   description           = 'Is produced by =ERROR.TYPE(1) or if C1 contains =A1:A2+B1:B3 -> C3=#N/A' ).
*    name                       = lcl_xlom__va_error=>create( english_error_name    = '#NAME?        '
*                                                                   internal_error_number = 2029
*                                                                   formula_error_number  = 5
*                                                                   description           = 'Is produced by =XXXX if XXXX is not an existing range name' ).
*    null                       = lcl_xlom__va_error=>create( english_error_name    = '#NULL!        '
*                                                                   internal_error_number = 2000
*                                                                   formula_error_number  = 1 ).
*    num                        = lcl_xlom__va_error=>create( english_error_name    = '#NUM!         '
*                                                                   internal_error_number = 2036
*                                                                   formula_error_number  = 6
*                                                                   description           = 'Is produced by =1E+240*1E+240' ).
*    python                     = lcl_xlom__va_error=>create( english_error_name    = '#PYTHON!      '
*                                                                   internal_error_number = 2222
*                                                                   formula_error_number  = 19 ).
*    ref                        = lcl_xlom__va_error=>create( english_error_name    = '#REF!         '
*                                                                   internal_error_number = 2023
*                                                                   formula_error_number  = 4
*                                                                   description           = 'Is produced by =INDEX(A1,2,1)' ).
*    spill                      = lcl_xlom__va_error=>create( english_error_name    = '#SPILL!       '
*                                                                   internal_error_number = 2045
*                                                                   formula_error_number  = 9
*                                                                   description           = 'Is produced by A1 containing ={1,2} and B1 containing a value -> A1=#SPILL!' ).
*    unknown                    = lcl_xlom__va_error=>create( english_error_name    = '#UNKNOWN!     '
*                                                                   internal_error_number = 2048
*                                                                   formula_error_number  = 12 ).
*    value_cannot_be_calculated = lcl_xlom__va_error=>create( english_error_name    = '#VALUE!       '
*                                                                   internal_error_number = 2015
*                                                                   formula_error_number  = 3
*                                                                   description           = 'Is produced by =1+"a"' ).
*  ENDMETHOD.
*
*  METHOD create.
*    result = NEW lcl_xlom__va_error( ).
*    result->lif_xlom__va~type = lif_xlom__va=>c_type-error.
*    result->english_error_name      = english_error_name.
*    result->internal_error_number   = internal_error_number.
*    result->formula_error_number    = formula_error_number.
*    result->description             = description.
*    INSERT VALUE #( english_error_name    = english_error_name
*                    internal_error_number = internal_error_number
*                    formula_error_number  = formula_error_number
*                    object                = result )
*           INTO TABLE errors.
*  ENDMETHOD.
*
*  METHOD get_by_error_number.
*    result = errors[ internal_error_number = type ]-object.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~get_value.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_array.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_boolean.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_equal.
*    RAISE EXCEPTION TYPE lcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_error.
*    result = abap_true.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_number.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_string.
*    result = abap_false.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_number IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__va_number( ).
*    result->lif_xlom__va~type = lif_xlom__va=>c_type-number.
*    result->number                  = number.
*  ENDMETHOD.
*
*  METHOD get.
*    DATA(buffer_line) = REF #( buffer[ number = number ] OPTIONAL ).
*    IF buffer_line IS NOT BOUND.
*      result = create( number ).
*      INSERT VALUE #( number = number
*                      object = result )
*             INTO TABLE buffer
*             REFERENCE INTO buffer_line.
*    ENDIF.
*    result = buffer_line->object.
*  ENDMETHOD.
*
*  METHOD get_integer.
*    " Excel rounding (1.99 -> 1, -1.99 -> -1)
*    result = floor( number ).
*  ENDMETHOD.
*
*  METHOD get_number.
*    result = number.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~get_value.
*    result = REF #( number ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_array.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_boolean.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_equal.
*    IF input_result->type = lif_xlom__va=>c_type-number.
*      DATA(input_number) = CAST lcl_xlom__va_number( input_result ).
*      IF number = input_number->number.
*        result = abap_true.
*      ELSE.
*        result = abap_false.
*      ENDIF.
*    ELSE.
*      result = abap_false.
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_error.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_number.
*    result = abap_true.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_string.
*    result = abap_false.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom__va_string IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom__va_string( ).
*    result->lif_xlom__va~type = lif_xlom__va=>c_type-string.
*    result->string                  = string.
*  ENDMETHOD.
*
*  METHOD get.
*    DATA(buffer_line) = REF #( buffer[ string = string ] OPTIONAL ).
*    IF buffer_line IS NOT BOUND.
*      INSERT VALUE #( string = string
*                      object = create( string ) )
*             INTO TABLE buffer
*             REFERENCE INTO buffer_line.
*    ENDIF.
*    result = buffer_line->object.
*  ENDMETHOD.
*
*  METHOD get_string.
*    result = string.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~get_value.
*    result = REF #( string ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_array.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_boolean.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_equal.
*    result = xsdbool( string = CAST lcl_xlom__va_string( input_result )->get_string( ) ).
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_error.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_number.
*    result = abap_false.
*  ENDMETHOD.
*
*  METHOD lif_xlom__va~is_string.
*    result = abap_true.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom_rows IMPLEMENTATION.
*  METHOD count.
*    result = lif_xlom__va_array~row_count.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom_workbook IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom_workbook( ).
*    result->application = application.
*    result->worksheets = lcl_xlom_worksheets=>create( workbook = result ).
*    result->worksheets->add( name = 'Sheet1' ).
*  ENDMETHOD.
*
*  METHOD save_as.
*    RAISE EVENT saved.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom_workbooks IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom_workbooks( ).
*    result->application = application.
*  ENDMETHOD.
*
*  METHOD add.
*    DATA workbook TYPE ty_workbook.
*
*    workbook-object = lcl_xlom_workbook=>create( application ).
*    INSERT workbook INTO TABLE workbooks.
*    count = count + 1.
*
*    set handler on_saved for workbook-object.
*    result = workbook-object.
*  ENDMETHOD.
*
*  METHOD item.
*    CASE lcl_xlom_application=>type( index ).
*      WHEN cl_abap_typedescr=>typekind_string.
*        result = workbooks[ name = index ]-object.
*      WHEN cl_abap_typedescr=>typekind_int.
*        result = workbooks[ index ]-object.
*      WHEN OTHERS.
*        " TODO
*    ENDCASE.
*  ENDMETHOD.
*
*  METHOD on_saved.
*    workbooks[ KEY by_object COMPONENTS object = sender ]-name = sender->name.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom_worksheet IMPLEMENTATION.
*  METHOD calculate.
**    RAISE EXCEPTION TYPE lcx_xlom_to_do.
*    " TODO the containing cell shouldn't be A1, it should vary for
*    "      each cell where the formula is to be calculated.
*    "      Solution: maybe store the range object within each cell
*    "                to not recalculate it (performance).
*    DATA(context) = lcl_xlom__ut_eval_context=>create( worksheet       = me
*                                                            containing_cell = VALUE #( row    = 1
*                                                                                       column = 1 ) ).
*    LOOP AT _array->_cells REFERENCE INTO DATA(cell)
*        WHERE formula    IS BOUND
*          and calculated  = abap_false.
*      context->set_containing_cell( VALUE #( row    = cell->row
*                                             column = cell->column ) ).
*      cell->value = cell->formula->evaluate( context ).
*      if sy-datum = '20241016' and cell->row = 4.
*        ASSERT 1 = 1. " Debug helper to set a break-point
*      endif.
*    ENDLOOP.
*  ENDMETHOD.
*
*  METHOD cells.
*    " This will change Z20:
*    " Range("Z20:AA25").Cells(1, 1) = "C"
*    result = lcl_xlom_range=>create_from_row_column( worksheet = me
*                                                     row       = row
*                                                     column    = column ).
*  ENDMETHOD.
*
*  METHOD create.
*    result = NEW lcl_xlom_worksheet( ).
*    result->name = name.
*    result->parent = workbook.
*    result->application = workbook->application.
*    result->_array = lcl_xlom__va_array=>create_initial( row_count    = max_rows
*                                                               column_count = max_columns ).
*  ENDMETHOD.
*
*  METHOD range.
*    IF    (     cell1_string IS NOT INITIAL
*            AND cell1_range  IS BOUND )
*       OR (     cell1_string IS INITIAL
*            AND cell1_range  IS NOT BOUND )
*       OR (     cell1_string IS INITIAL
*            AND cell2_string IS NOT INITIAL )
*       OR (     cell1_range  IS NOT BOUND
*            AND cell2_range  IS BOUND ).
*      RAISE EXCEPTION TYPE lcx_xlom_todo.
*    ENDIF.
*
*    IF cell1_string IS NOT INITIAL.
*      result = range_from_address( cell1 = cell1_string
*                                   cell2 = cell2_string ).
*    ELSE.
*      result = range_from_two_ranges( cell1 = cell1_range
*                                      cell2 = cell2_range ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD range_from_address.
*    DATA(range_1) = lcl_xlom_range=>create_from_address_or_name( address     = cell1
*                                                                 relative_to = me ).
*    IF cell2 IS INITIAL.
*      result = range_1.
*    ELSE.
*      DATA(range_2) = lcl_xlom_range=>create_from_address_or_name( address     = cell2
*                                                                      relative_to = me ).
*      result = range_from_two_ranges( cell1 = range_1
*                                      cell2 = range_2 ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD range_from_two_ranges.
*    result = lcl_xlom_range=>create( cell1 = cell1
*                                        cell2 = cell2 ).
*  ENDMETHOD.
*
*  METHOD used_range.
*    result = lcl_xlom_range=>create_from_row_column(
*                 worksheet   = me
*                 row         = _array->used_range-top_left-row
*                 column      = _array->used_range-top_left-column
*                 row_size    = _array->used_range-bottom_right-row - _array->used_range-top_left-row + 1
*                 column_size = _array->used_range-bottom_right-column - _array->used_range-top_left-column + 1 ).
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcl_xlom_worksheets IMPLEMENTATION.
*  METHOD create.
*    result = NEW lcl_xlom_worksheets( ).
*    result->application = workbook->application.
*    result->workbook = workbook.
*  ENDMETHOD.
*
*  METHOD add.
*    DATA worksheet TYPE ty_worksheet.
*
*    worksheet-name   = name.
*    worksheet-object = lcl_xlom_worksheet=>create( workbook = workbook
*                                                      name     = name ).
*    INSERT worksheet INTO TABLE worksheets.
*    count = count + 1.
*
*    application->active_sheet = worksheet-object.
*
*    result = worksheet-object.
*  ENDMETHOD.
*
*  METHOD item.
*    TRY.
*        CASE lcl_xlom_application=>type( index ).
*          WHEN cl_abap_typedescr=>typekind_string
*            OR cl_abap_typedescr=>typekind_char.
*            result = worksheets[ name = index ]-object.
*          WHEN cl_abap_typedescr=>typekind_int.
*            result = worksheets[ index ]-object.
*          WHEN OTHERS.
*            RAISE EXCEPTION TYPE lcx_xlom_todo.
*        ENDCASE.
*      CATCH cx_sy_itab_line_not_found.
*        RAISE EXCEPTION TYPE lcx_xlom__va
*          EXPORTING result_error = lcl_xlom__va_error=>ref.
*    ENDTRY.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcx_xlom__ex_ut_parser IMPLEMENTATION.
*  METHOD constructor ##ADT_SUPPRESS_GENERATION.
*    super->constructor( previous = previous
*                        textid   = textid ).
*    me->text  = text.
*    me->msgv1 = msgv1.
*    me->msgv2 = msgv2.
*    me->msgv3 = msgv3.
*    me->msgv4 = msgv4.
*  ENDMETHOD.
*
*  METHOD get_longtext.
*    IF text IS NOT INITIAL.
*      result = get_text( ).
*    ELSE.
*      result = super->get_longtext( ).
*    ENDIF.
*  ENDMETHOD.
*
*  METHOD get_text.
*    IF text IS NOT INITIAL.
*      result = text.
*      REPLACE '&1' IN result WITH msgv1.
*      REPLACE '&2' IN result WITH msgv2.
*      REPLACE '&3' IN result WITH msgv3.
*      REPLACE '&4' IN result WITH msgv4.
*    ELSE.
*      result = super->get_text( ).
*    ENDIF.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS lcx_xlom__va IMPLEMENTATION.
*  METHOD constructor.
*    super->constructor( textid = textid previous = previous ).
*    me->result_error = result_error.
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS ltc_evaluate DEFINITION FINAL
*  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*  PRIVATE SECTION.
*    METHODS ampersand                  FOR TESTING RAISING cx_static_check.
*    METHODS array                      FOR TESTING RAISING cx_static_check.
*    METHODS cell                       FOR TESTING RAISING cx_static_check.
*    METHODS colon                      FOR TESTING RAISING cx_static_check.
*    METHODS complex_1                  FOR TESTING RAISING cx_static_check.
*    METHODS complex_2                  FOR TESTING RAISING cx_static_check.
*    METHODS countif                    FOR TESTING RAISING cx_static_check.
*    METHODS equal                      FOR TESTING RAISING cx_static_check.
*    METHODS error                      FOR TESTING RAISING cx_static_check.
*    METHODS find                       FOR TESTING RAISING cx_static_check.
*    METHODS function_optional_argument FOR TESTING RAISING cx_static_check.
*    METHODS if                         FOR TESTING RAISING cx_static_check.
*    METHODS iferror                    FOR TESTING RAISING cx_static_check.
*    METHODS index                      FOR TESTING RAISING cx_static_check.
*    "! Array evaluation, e.g. INDEX(A1:B2,{2,1},{2,1})
*    METHODS index_ae                   FOR TESTING RAISING cx_static_check.
*    METHODS indirect                   FOR TESTING RAISING cx_static_check.
*    METHODS len                        FOR TESTING RAISING cx_static_check.
*    METHODS len_a1_a2                  FOR TESTING RAISING cx_static_check.
*    METHODS match                      FOR TESTING RAISING cx_static_check.
*    METHODS match_2                    FOR TESTING RAISING cx_static_check.
*    METHODS minus                      FOR TESTING RAISING cx_static_check.
*    METHODS mult                       FOR TESTING RAISING cx_static_check.
*    METHODS number                     FOR TESTING RAISING cx_static_check.
*    METHODS offset                     FOR TESTING RAISING cx_static_check.
*    METHODS plus                       FOR TESTING RAISING cx_static_check.
*    METHODS range_a1_plus_one          FOR TESTING RAISING cx_static_check.
*    METHODS range_two_sheets           FOR TESTING RAISING cx_static_check.
*    METHODS right                      FOR TESTING RAISING cx_static_check.
*    METHODS right_2                    FOR TESTING RAISING cx_static_check.
*    METHODS row                        FOR TESTING RAISING cx_static_check.
*    METHODS string                     FOR TESTING RAISING cx_static_check.
*    METHODS t                          FOR TESTING RAISING cx_static_check.
*
*    TYPES tt_parenthesis_group TYPE lcl_xlom__ex_ut_lexer=>tt_parenthesis_group.
*    TYPES tt_token             TYPE lcl_xlom__ex_ut_lexer=>tt_token.
*    TYPES ts_result_lexe       TYPE lcl_xlom__ex_ut_lexer=>ts_result_lexe.
*
*    DATA worksheet TYPE REF TO lcl_xlom_worksheet.
*    DATA range_a1  TYPE REF TO lcl_xlom_range.
*    DATA: range_a2 TYPE REF TO lcl_xlom_range,
*          range_b1 TYPE REF TO lcl_xlom_range,
*          range_b2 TYPE REF TO lcl_xlom_range,
*          range_c1 TYPE REF TO lcl_xlom_range,
*          range_d1 TYPE REF TO lcl_xlom_range,
*          application TYPE REF TO lcl_xlom_application,
*          workbook TYPE REF TO lcl_xlom_workbook.
*
*    METHODS assert_equals
*      IMPORTING act            TYPE REF TO lif_xlom__va
*                exp            TYPE REF TO lif_xlom__va
*      RETURNING VALUE(result) TYPE abap_bool.
*
*    METHODS setup.
*ENDCLASS.
*
*
*CLASS ltc_lexer DEFINITION FINAL
*  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.
*
*  PRIVATE SECTION.
*    METHODS arithmetic                     FOR TESTING RAISING cx_static_check.
*    METHODS array                          FOR TESTING RAISING cx_static_check.
*    METHODS array_two_rows                 FOR TESTING RAISING cx_static_check.
*    METHODS error_name                     FOR TESTING RAISING cx_static_check.
*    METHODS function                       FOR TESTING RAISING cx_static_check.
*    METHODS function_function              FOR TESTING RAISING cx_static_check.
*    METHODS function_optional_argument     FOR TESTING RAISING cx_static_check.
*    METHODS number                         FOR TESTING RAISING cx_static_check.
*    METHODS operator_function              FOR TESTING RAISING cx_static_check.
*    METHODS range                          FOR TESTING RAISING cx_static_check.
*    METHODS smart_table                    FOR TESTING RAISING cx_static_check.
*    METHODS smart_table_all                FOR TESTING RAISING cx_static_check.
*    METHODS smart_table_column             FOR TESTING RAISING cx_static_check.
*    METHODS smart_table_no_space           FOR TESTING RAISING cx_static_check.
*    METHODS smart_table_space_separator    FOR TESTING RAISING cx_static_check.
*    METHODS smart_table_space_boundaries   FOR TESTING RAISING cx_static_check.
*    METHODS smart_table_space_all          FOR TESTING RAISING cx_static_check.
*    METHODS text_literal                   FOR TESTING RAISING cx_static_check.
*    METHODS text_literal_with_double_quote FOR TESTING RAISING cx_static_check.
*    METHODS very_long                      FOR TESTING RAISING cx_static_check.
*
*    TYPES tt_parenthesis_group TYPE lcl_xlom__ex_ut_lexer=>tt_parenthesis_group.
*    TYPES tt_token             TYPE lcl_xlom__ex_ut_lexer=>tt_token.
*    TYPES ts_result_lexe       TYPE lcl_xlom__ex_ut_lexer=>ts_result_lexe.
*
*    CONSTANTS c_type LIKE lcl_xlom__ex_ut_lexer=>c_type VALUE lcl_xlom__ex_ut_lexer=>c_type.
*
*    METHODS lexe
*      IMPORTING !text         TYPE csequence
*      RETURNING VALUE(result) TYPE tt_token."ts_result_lexe.
*
*ENDCLASS.
*
*
*CLASS ltc_parser DEFINITION FINAL
*  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.
*
*  PRIVATE SECTION.
*    METHODS array                          FOR TESTING RAISING cx_static_check.
*    METHODS array_two_rows                 FOR TESTING RAISING cx_static_check.
*    METHODS function_argument_minus_unary  FOR TESTING RAISING cx_static_check.
*    METHODS function_function              FOR TESTING RAISING cx_static_check.
*    METHODS function_optional_argument     FOR TESTING RAISING cx_static_check.
*    METHODS if                             FOR TESTING RAISING cx_static_check.
*    METHODS number                         FOR TESTING RAISING cx_static_check.
*    METHODS one_plus_one                   FOR TESTING RAISING cx_static_check.
*    METHODS operator_function              FOR TESTING RAISING cx_static_check.
*    METHODS operator_function_operator     FOR TESTING RAISING cx_static_check.
*    METHODS parentheses_arithmetic         FOR TESTING RAISING cx_static_check.
*    METHODS parentheses_arithmetic_complex FOR TESTING RAISING cx_static_check.
*    METHODS priority                       FOR TESTING RAISING cx_static_check.
*    METHODS very_long                      FOR TESTING RAISING cx_static_check.
*
*    TYPES tt_token       TYPE lcl_xlom__ex_ut_lexer=>tt_token.
*    TYPES ts_result_lexe TYPE lcl_xlom__ex_ut_lexer=>ts_result_lexe.
*
*    CONSTANTS c_type LIKE lcl_xlom__ex_ut_lexer=>c_type VALUE lcl_xlom__ex_ut_lexer=>c_type.
*
*    METHODS assert_equals
*      IMPORTING act            TYPE REF TO lif_xlom__ex
*                exp            TYPE REF TO lif_xlom__ex
*      RETURNING VALUE(result)  TYPE REF TO lif_xlom__va.
*
*    METHODS lexe
*      IMPORTING !text         TYPE csequence
*      RETURNING VALUE(result) TYPE tt_token. "ts_result_lexe.
*
*    METHODS parse
*      IMPORTING !tokens            TYPE lcl_xlom__ex_ut_lexer=>tt_token
*      RETURNING VALUE(result)      TYPE REF TO lif_xlom__ex
*      RAISING   lcx_xlom__ex_ut_parser.
*ENDCLASS.
*
*
*CLASS ltc_range DEFINITION FINAL
*  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.
*
*  PUBLIC SECTION.
*    INTERFACES lif_xlom__ut_all_friends.
*
*  PRIVATE SECTION.
*    METHODS convert_column_a_xfd_to_number FOR TESTING RAISING cx_static_check.
*    METHODS decode_range_address_a1_invali FOR TESTING RAISING cx_static_check.
*    METHODS decode_range_address_a1_valid  FOR TESTING RAISING cx_static_check.
*    METHODS decode_range_address_sh_invali FOR TESTING RAISING cx_static_check.
*    METHODS decode_range_address_sh_valid FOR TESTING RAISING cx_static_check.
*    METHODS convert_column_number_to_a_xfd FOR TESTING RAISING cx_static_check.
*
*    TYPES ty_address TYPE lif_xlom__va_array=>ts_address.
*ENDCLASS.
*
*
*CLASS ltc_evaluate IMPLEMENTATION.
*  METHOD ampersand.
*    range_a1->set_formula2( value = `"hello "&"world"` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1 )->get_string( )
*                                        exp = `hello world` ).
*    range_a1->set_formula2( value = `"hello "&"new "&"world"` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = `hello new world` ).
*  ENDMETHOD.
*
*  METHOD array.
*    range_a1->set_formula2( value = `{1,2}` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1 )->get_number( )
*                                        exp = 1 ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_b1 )->get_number( )
*                                        exp = 2 ).
*  ENDMETHOD.
*
*  METHOD assert_equals.
*    cl_abap_unit_assert=>assert_true( xsdbool( exp->is_equal( act ) ) ).
*  ENDMETHOD.
*
*  METHOD cell.
** TODO not very clear what it should do without the Reference argument...
**    range_a1->set_formula2( value = `CELL("filename")` ).
**    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va_converter=>to_string( range_a1->value( ) )->get_string( )
**                                        exp = `\[]Sheet1` ).
*    range_a2->set_value( lcl_xlom__va_string=>create( '' ) ).
*    range_a1->set_formula2( value = `CELL("filename",A2)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = `\[]Sheet1` ).
*  ENDMETHOD.
*
*  METHOD colon.
*    DATA dummy_ref_to_offset TYPE REF TO lcl_xlom__ex_op_colon ##NEEDED.
*    DATA result              TYPE REF TO lif_xlom__va.
*
*    result = application->evaluate( `3:3` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$3:$3' ).
*
*    result = application->evaluate( `C:C` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$C:$C' ).
*  ENDMETHOD.
*
*  METHOD complex_1.
*    range_a2->set_formula2( value = `"'"&RIGHT(CELL("filename",A1),LEN(CELL("filename",A1))-FIND("]",CELL("filename",A1)))&" (2)'!$1:$1"` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a2->value( ) )->get_string( )
*                                        exp = `'Sheet1 (2)'!$1:$1` ).
*  ENDMETHOD.
*
*  METHOD complex_2.
*    DATA dummy_ref_to_iferror   TYPE REF TO lcl_xlom__ex_fu_iferror ##NEEDED.
*    DATA dummy_ref_to_t         TYPE REF TO lcl_xlom__ex_fu_t ##NEEDED.
*    DATA dummy_ref_to_ampersand TYPE REF TO lcl_xlom__ex_op_ampersand ##NEEDED.
*    DATA dummy_ref_to_index     TYPE REF TO lcl_xlom__ex_fu_index ##NEEDED.
*    DATA dummy_ref_to_offset    TYPE REF TO lcl_xlom__ex_fu_offset ##NEEDED.
*    DATA dummy_ref_to_indirect  TYPE REF TO lcl_xlom__ex_fu_indirect ##NEEDED.
*    DATA dummy_ref_to_minus     TYPE REF TO lcl_xlom__ex_op_minus ##NEEDED.
*    DATA dummy_ref_to_row       TYPE REF TO lcl_xlom__ex_fu_row ##NEEDED.
*    DATA dummy_ref_to_match     TYPE REF TO lcl_xlom__ex_fu_match ##NEEDED.
*
*    DATA(worksheet_bkpf) = workbook->worksheets->add( 'BKPF' ).
*    worksheet_bkpf->range_from_address( 'A1' )->set_value( lcl_xlom__va_string=>create( 'ID_REF_TEST' ) ).
*    worksheet_bkpf->range_from_address( 'B2' )->set_value( lcl_xlom__va_string=>create( `'BKPF (2)'!$1:$1` ) ).
*
*    DATA(worksheet_bkpf_2) = workbook->worksheets->add( 'BKPF (2)' ).
*    worksheet_bkpf_2->range_from_address( 'A1' )->set_value( lcl_xlom__va_string=>create( 'ID_REF_TEST' ) ).
*    worksheet_bkpf_2->range_from_address( 'A3' )->set_value( lcl_xlom__va_string=>create( 'MY_TEST' ) ).
*
*    DATA(range_bkpf_a3) = worksheet_bkpf->range_from_address( 'A3' ).
*    range_bkpf_a3->set_formula2( value = `IFERROR(T(""&INDEX(OFFSET(INDIRECT(BKPF!$B$2),ROW()-1,0),1,MATCH(A$1,INDIRECT(BKPF!$B$2),0))),"")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_bkpf_a3->value( ) )->get_string( )
*                                        exp = `MY_TEST` ).
*  ENDMETHOD.
*
*  METHOD countif.
*    range_a1->set_value( lcl_xlom__va_string=>create( `Hello` ) ).
*    range_a2->set_value( lcl_xlom__va_string=>create( `world` ) ).
*    range_b1->set_value( lcl_xlom__va_string=>create( `peace` ) ).
*    range_b2->set_value( lcl_xlom__va_string=>create( `love` ) ).
*    range_c1->set_formula2( value = `COUNTIF(A1:B2,"*e*")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_c1->value( ) )->get_integer( )
*                                        exp = 3 ).
*  ENDMETHOD.
*
*  METHOD equal.
*    range_a1->set_formula2( value = `1=1` ).
*    cl_abap_unit_assert=>assert_true( lcl_xlom__va=>to_boolean( range_a1->value( ) )->boolean_value ).
*  ENDMETHOD.
*
*  METHOD error.
*    range_a1->set_formula2( value = `#N/A` ).
*    cl_abap_unit_assert=>assert_equals( act = range_a1->value( )->type
*                                        exp = lif_xlom__va=>c_type-error ).
*  ENDMETHOD.
*
*  METHOD find.
*    range_a1->set_formula2( value = `FIND("b","abc")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 2 ).
*  ENDMETHOD.
*
*  METHOD function_optional_argument.
*    range_a1->set_formula2( value = `RIGHT("hello",0)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = '' ).
*    range_a1->set_formula2( value = `RIGHT("hello",)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = '' ).
*  ENDMETHOD.
*
*  METHOD if.
*    range_a1->set_formula2( value = `IF(0=1,2,4)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 4 ).
*  ENDMETHOD.
*
*  METHOD iferror.
*    range_a1->set_formula2( value = `IFERROR(#N/A,1)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 1 ).
*    range_a1->set_formula2( value = `IFERROR(2,1)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 2 ).
*  ENDMETHOD.
*
*  METHOD index.
*    range_b2->set_value( lcl_xlom__va_string=>create( `Hello` ) ).
*    range_a1->set_formula2( value = `INDEX(A1:C3,2,2)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = `Hello` ).
*
*  ENDMETHOD.
*
*  METHOD index_ae.
*    range_A1->set_value( lcl_xlom__va_string=>create( `Hello ` ) ).
*    range_b2->set_value( lcl_xlom__va_string=>create( `world` ) ).
*    range_c1->set_formula2( value = `INDEX(A1:B2,{2,1},{2,1})` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_c1->value( ) )->get_string( )
*                                        exp = `world` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_d1->value( ) )->get_string( )
*                                        exp = `Hello ` ).
*
*    range_a1->set_formula2( value = `INDEX({"a","b";"c","d"},{2,1},{2,1})` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = `d` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_b1->value( ) )->get_string( )
*                                        exp = `a` ).
*  ENDMETHOD.
*
*  METHOD indirect.
*    range_a1->set_value( lcl_xlom__va_string=>create( `Hello` ) ).
*    range_a2->set_formula2( value = `INDIRECT("A1")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = `Hello` ).
*  ENDMETHOD.
*
*  METHOD len.
*    range_a1->set_formula2( value = `LEN("ABC")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 3 ).
*    range_a1->set_formula2( value = `LEN("ABC ")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 4 ).
*    range_a1->set_formula2( value = `LEN("")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 0 ).
*  ENDMETHOD.
*
*  METHOD len_a1_a2.
*    range_a1->set_value( lcl_xlom__va_string=>create( `Hello ` ) ).
*    range_a2->set_value( lcl_xlom__va_string=>create( `world` ) ).
*    range_b1->set_formula2( value = `LEN(A1:A2)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_b1->value( ) )->get_number( )
*                                        exp = 6 ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_b2->value( ) )->get_number( )
*                                        exp = 5 ).
*  ENDMETHOD.
*
*  METHOD match.
*    range_a1->set_value( lcl_xlom__va_string=>create( `Hello ` ) ).
*    range_a2->set_value( lcl_xlom__va_string=>create( `world` ) ).
*    range_b1->set_formula2( value = `MATCH("world",A1:A2,0)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_b1->value( ) )->get_number( )
*                                        exp = 2 ).
*  ENDMETHOD.
*
*  METHOD match_2.
*    range_a1->set_value( lcl_xlom__va_string=>create( `Hello ` ) ).
*    range_a2->set_value( lcl_xlom__va_string=>create( `world` ) ).
*    range_b1->set_formula2( value = `MATCH("world",A:A,0)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_b1->value( ) )->get_number( )
*                                        exp = 2 ).
*  ENDMETHOD.
*
*  METHOD minus.
*    range_a1->set_formula2( value = `5-3` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 2 ).
*  ENDMETHOD.
*
*  METHOD mult.
*    range_a1->set_formula2( value = `2*3*4` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 24 ).
*  ENDMETHOD.
*
*  METHOD number.
*    range_a1->set_formula2( value = `1` ).
*    assert_equals( act = range_a1->value( )
*                   exp = lcl_xlom__va_number=>get( 1 ) ).
*
*    range_a1->set_formula2( value = `-1` ).
*    assert_equals( act = range_a1->value( )
*                   exp = lcl_xlom__va_number=>get( -1 ) ).
*  ENDMETHOD.
*
*  METHOD offset.
*    DATA dummy_ref_to_offset TYPE REF TO lcl_xlom__ex_fu_offset ##NEEDED.
*    DATA result              TYPE REF TO lif_xlom__va.
*
*    result = application->evaluate( `OFFSET(A1,1,1)` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$B$2' ).
*
*    result = application->evaluate( `OFFSET(A1,2,0)` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$A$3' ).
*
*    result = application->evaluate( `OFFSET(A1,2,2)` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$C$3' ).
*
*    result = application->evaluate( `OFFSET(C2,-1,-2)` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$A$1' ).
*
*    result = application->evaluate( `OFFSET(A1,1,1,,)` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$B$2' ).
*
*    result = application->evaluate( `OFFSET(A1,1,1,2,2)` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$B$2:$C$3' ).
*
*    result = application->evaluate( `OFFSET(1:1,1,0)` ).
*    cl_abap_unit_assert=>assert_equals( act = CAST lcl_xlom_range( result )->address( )
*                                        exp = '$2:$2' ).
*  ENDMETHOD.
*
*  METHOD plus.
*    range_a1->set_formula2( value = `1+1` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a1->value( ) )->get_number( )
*                                        exp = 2 ).
*  ENDMETHOD.
*
*  METHOD range_a1_plus_one.
*    range_a1->set_value( lcl_xlom__va_number=>create( 10 ) ).
*    DATA(range_a2) = worksheet->range_from_address( 'A2' ).
*    range_a2->set_formula2( 'A1+1' ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_number( range_a2->value( ) )->get_number( )
*                                        exp = 11 ).
*  ENDMETHOD.
*
*  METHOD range_two_sheets.
*    DATA dummy_ref_to_offset TYPE REF TO lcl_xlom__ex_fu_offset ##NEEDED.
*
*    range_a1->set_value( lcl_xlom__va_string=>create( `Hello` ) ).
*
*    DATA(worksheet_2) = workbook->worksheets->add( 'Sheet2' ).
*    DATA(range_sheet2_b2) = worksheet_2->range_from_address( 'B2' ).
*    range_sheet2_b2->set_formula2( |"C"&Sheet1!A1| ).
*
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_sheet2_b2->value( ) )->get_string( )
*                                        exp = `CHello` ).
*  ENDMETHOD.
*
*  METHOD right.
*    range_a1->set_formula2( value = `RIGHT("Hello",2)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = 'lo' ).
*    range_a1->set_formula2( value = `RIGHT(25,1)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = '5' ).
*    range_a1->set_formula2( value = `RIGHT("hello")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = 'o' ).
*  ENDMETHOD.
*
*  METHOD right_2.
*    range_a1->set_formula2( value = `RIGHT("hello")` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = 'o' ).
*    range_a1->set_formula2( value = `RIGHT("hello",0)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = '' ).
*    range_a1->set_formula2( value = `RIGHT("hello",)` ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom__va=>to_string( range_a1->value( ) )->get_string( )
*                                        exp = '' ).
*  ENDMETHOD.
*
*  METHOD row.
*    range_a1->set_formula2( value = `ROW(B2)` ).
*    assert_equals( act = range_a1->value( )
*                   exp = lcl_xlom__va_number=>create( 2 ) ).
*    range_a1->set_formula2( value = `ROW()` ).
*    assert_equals( act = range_a1->value( )
*                   exp = lcl_xlom__va_number=>create( 1 ) ).
*  ENDMETHOD.
*
*  METHOD setup.
*    application = lcl_xlom_application=>create( ).
*    workbook = application->workbooks->add( ).
*    TRY.
*    worksheet = workbook->worksheets->item( 'Sheet1' ).
*    range_a1 = worksheet->range_from_address( 'A1' ).
*    range_a2 = worksheet->range_from_address( 'A2' ).
*    range_b1 = worksheet->range_from_address( 'B1' ).
*    range_b2 = worksheet->range_from_address( 'B2' ).
*    range_c1 = worksheet->range_from_address( 'C1' ).
*    range_d1 = worksheet->range_from_address( 'D1' ).
*    CATCH lcx_xlom__va INTO DATA(error).
*      cl_abap_unit_assert=>fail( 'unexpected' ).
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD string.
*    range_a1->set_formula2( value = `"1"` ).
*    assert_equals( act = range_a1->value( )
*                   exp = lcl_xlom__va_string=>create( '1' ) ).
*  ENDMETHOD.
*
*  METHOD t.
*    range_a1->set_formula2( value = `T("1")` ).
*    assert_equals( act = range_a1->value( )
*                   exp = lcl_xlom__va_string=>create( '1' ) ).
*
*    range_a1->set_formula2( value = `T(1)` ).
*    assert_equals( act = range_a1->value( )
*                   exp = lcl_xlom__va_string=>create( '' ) ).
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS ltc_lexer IMPLEMENTATION.
*  METHOD arithmetic.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '2*(1+3*(5+1))' )
*                                        exp = VALUE tt_token( ( value = `2`  type = c_type-number )
*                                                              ( value = `*`  type = c_type-operator )
*                                                              ( value = `(`  type = c_type-parenthesis_open )
*                                                              ( value = `1`  type = c_type-number )
*                                                              ( value = `+`  type = c_type-operator )
*                                                              ( value = `3`  type = c_type-number )
*                                                              ( value = `*`  type = c_type-operator )
*                                                              ( value = `(`  type = c_type-parenthesis_open )
*                                                              ( value = `5`  type = c_type-number )
*                                                              ( value = `+`  type = c_type-operator )
*                                                              ( value = `1`  type = c_type-number )
*                                                              ( value = `)`  type = c_type-parenthesis_close )
*                                                              ( value = `)`  type = c_type-parenthesis_close ) ) ).
*  ENDMETHOD.
*
*  METHOD array.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '{1,2}' )
*                                        exp = VALUE tt_token( ( value = `{` type = c_type-curly_bracket_open )
*                                                              ( value = `1` type = c_type-number )
*                                                              ( value = `,` type = c_type-comma )
*                                                              ( value = `2` type = c_type-number )
*                                                              ( value = `}` type = c_type-curly_bracket_close ) ) ).
*  ENDMETHOD.
*
*  METHOD array_two_rows.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '{1,2;3,4}' )
*                                        exp = VALUE tt_token( ( value = `{` type = c_type-curly_bracket_open )
*                                                              ( value = `1` type = c_type-number )
*                                                              ( value = `,` type = c_type-comma )
*                                                              ( value = `2` type = c_type-number )
*                                                              ( value = `;` type = c_type-semicolon )
*                                                              ( value = `3` type = c_type-number )
*                                                              ( value = `,` type = c_type-comma )
*                                                              ( value = `4` type = c_type-number )
*                                                              ( value = `}` type = c_type-curly_bracket_close ) ) ).
*  ENDMETHOD.
*
*  METHOD error_name.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '#N/A!' )
*                                        exp = VALUE tt_token( ( value = `#N/A!` type = c_type-error_name ) ) ).
*  ENDMETHOD.
*
*  METHOD function.
*    cl_abap_unit_assert=>assert_equals( act = lexe( 'IF(1=1,0,1)' )
*                                        exp = VALUE tt_token( ( value = `IF` type = c_type-function_name )
*                                                              ( value = `1`  type = c_type-number )
*                                                              ( value = `=`  type = c_type-operator )
*                                                              ( value = `1`  type = c_type-number )
*                                                              ( value = `,`  type = ',' )
*                                                              ( value = `0`  type = c_type-number )
*                                                              ( value = `,`  type = ',' )
*                                                              ( value = `1`  type = c_type-number )
*                                                              ( value = `)`  type = ')' ) ) ).
*  ENDMETHOD.
*
*  METHOD function_function.
*    cl_abap_unit_assert=>assert_equals( act = lexe( 'LEN(RIGHT("text",2))' )
*                                        exp = VALUE tt_token( ( value = `LEN`   type = c_type-function_name )
*                                                              ( value = `RIGHT` type = c_type-function_name )
*                                                              ( value = `text`  type = c_type-text_literal )
*                                                              ( value = `,`     type = c_type-comma )
*                                                              ( value = `2`     type = c_type-number )
*                                                              ( value = `)`     type = c_type-parenthesis_close )
*                                                              ( value = `)`     type = c_type-parenthesis_close ) ) ).
*  ENDMETHOD.
*
*  METHOD function_optional_argument.
*    cl_abap_unit_assert=>assert_equals( act = lexe( 'RIGHT("text",)' )
*                                        exp = VALUE tt_token( ( value = `RIGHT` type = c_type-function_name )
*                                                              ( value = `text`  type = c_type-text_literal )
*                                                              ( value = `,`     type = c_type-comma )
*                                                              ( value = `)`     type = c_type-parenthesis_close ) ) ).
*  ENDMETHOD.
*
*  METHOD lexe.
*    result = lcl_xlom__ex_ut_lexer=>create( )->lexe( text ).
*  ENDMETHOD.
*
*  METHOD number.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '25' )
*                                        exp = VALUE tt_token( ( value = `25` type = c_type-number ) ) ).
*
*    cl_abap_unit_assert=>assert_equals( act = lexe( '-1' )
*                                        exp = VALUE tt_token( ( value = `-` type = c_type-operator )
*                                                              ( value = `1` type = c_type-number ) ) ).
*  ENDMETHOD.
*
*  METHOD operator_function.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '1+LEN("text")' )
*                                        exp = VALUE tt_token( ( value = `1`    type = c_type-number )
*                                                              ( value = `+`    type = c_type-operator )
*                                                              ( value = `LEN`  type = c_type-function_name )
*                                                              ( value = `text` type = c_type-text_literal )
*                                                              ( value = `)`    type = c_type-parenthesis_close ) ) ).
*  ENDMETHOD.
*
*  METHOD range.
*    cl_abap_unit_assert=>assert_equals( act = lexe( 'Sheet1!$A$1' )
*                                        exp = VALUE tt_token( ( value = `Sheet1!$A$1` type = 'W' ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lexe( `'Sheet 1'!$A$1` )
*                                        exp = VALUE tt_token( ( value = `'Sheet 1'!$A$1` type = 'W' ) ) ).
*  ENDMETHOD.
*
*  METHOD smart_table.
*    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[]' )
*                                        exp = VALUE tt_token( ( value = `Table1` type = c_type-table_name )
*                                                              ( value = `]`      type = `]` ) ) ).
*  ENDMETHOD.
*
*  METHOD smart_table_all.
*    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[[#All]]' )
*                                        exp = VALUE tt_token( ( value = `Table1` type = c_type-table_name )
*                                                              ( value = `[#All]` type = c_type-square_bracket_open )
*                                                              ( value = `]`      type = c_type-square_bracket_close ) ) ).
*  ENDMETHOD.
*
*  METHOD smart_table_column.
*    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[Column1]' )
*                                        exp = VALUE tt_token( ( value = `Table1`  type = c_type-table_name )
*                                                              ( value = `[Column1]` type = c_type-square_bracket_open )
*                                                              ( value = `]`      type = c_type-square_bracket_close ) ) ).
*  ENDMETHOD.
*
*  METHOD smart_table_no_space.
*    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
*    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[[#Headers],[#Data],[% Commission]]` )
*                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
*                                                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[#Data]`        type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[% Commission]` type = c_type-square_bracket_open )
*                                                              ( value = `]`              type = `]` ) ) ).
*  ENDMETHOD.
*
*  METHOD smart_table_space_all.
*    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
*    DATA(act) = lexe( `DeptSales[ [#Headers], [#Data], [% Commission] ]` ).
*    cl_abap_unit_assert=>assert_equals( act = act
*                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
*                                                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[#Data]`        type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[% Commission]` type = c_type-square_bracket_open )
*                                                              ( value = `]`              type = `]` ) ) ).
*  ENDMETHOD.
*
*  METHOD smart_table_space_boundaries.
*    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
*    DATA(act) = lexe( `DeptSales[ [#Headers],[#Data],[% Commission] ]` ).
*    cl_abap_unit_assert=>assert_equals( act = act
*                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
*                                                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[#Data]`        type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[% Commission]` type = c_type-square_bracket_open )
*                                                              ( value = `]`              type = `]` ) ) ).
*  ENDMETHOD.
*
*  METHOD smart_table_space_separator.
*    " https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
*    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[[#Headers], [#Data], [% Commission]]` )
*                                        exp = VALUE tt_token( ( value = `DeptSales`      type = c_type-table_name )
*                                                              ( value = `[#Headers]`     type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[#Data]`        type = c_type-square_bracket_open )
*                                                              ( value = `,`              type = `,` )
*                                                              ( value = `[% Commission]` type = c_type-square_bracket_open )
*                                                              ( value = `]`              type = `]` ) ) ).
*  ENDMETHOD.
*
*  METHOD text_literal.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '"IF(1=1,0,1)"' )
*                                        exp = VALUE tt_token( ( value = `IF(1=1,0,1)` type = c_type-text_literal ) ) ).
*  ENDMETHOD.
*
*  METHOD text_literal_with_double_quote.
*    cl_abap_unit_assert=>assert_equals( act = lexe( '"IF(A1=""X"",0,1)"' )
*                                        exp = VALUE tt_token( ( value = `IF(A1="X",0,1)` type = c_type-text_literal ) ) ).
*  ENDMETHOD.
*
*  METHOD very_long.
*    cl_abap_unit_assert=>assert_equals( act = lexe( |(a{ repeat( val = ',a'
*                                                                 occ = 5000 )
*                                                    })| )
*                                        exp = VALUE tt_token( ( value = `(` type = '(' )
*                                                              ( value = `a` type = 'W' )
*                                                              ( LINES OF VALUE
*                                                                tt_token( FOR i = 1 WHILE i <= 5000
*                                                                          ( value = `,` type = ',' )
*                                                                          ( value = `a` type = 'W' ) ) )
*                                                              ( value = `)` type = ')' ) ) ).
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS ltc_parser IMPLEMENTATION.
*  METHOD array.
*    assert_equals(
*        act = parse( tokens = VALUE #( ( value = `{` type = c_type-curly_bracket_open )
*                                       ( value = `1` type = c_type-number )
*                                       ( value = `,` type = c_type-comma )
*                                       ( value = `2` type = c_type-number )
*                                       ( value = `}` type = c_type-curly_bracket_close ) ) )
*        exp = lcl_xlom__ex_el_array=>create(
*                  rows = VALUE #( ( columns_of_row = VALUE #( ( lcl_xlom__ex_el_number=>create( 1 ) )
*                                                              ( lcl_xlom__ex_el_number=>create( 2 ) ) ) ) ) ) ).
*  ENDMETHOD.
*
*  METHOD array_two_rows.
*    assert_equals(
*        act = parse( tokens = VALUE #( ( value = `{` type = c_type-curly_bracket_open )
*                                       ( value = `1` type = c_type-number )
*                                       ( value = `,` type = c_type-comma )
*                                       ( value = `2` type = c_type-number )
*                                       ( value = `;` type = c_type-semicolon )
*                                       ( value = `3` type = c_type-number )
*                                       ( value = `,` type = c_type-comma )
*                                       ( value = `4` type = c_type-number )
*                                       ( value = `}` type = c_type-curly_bracket_close ) ) )
*        exp = lcl_xlom__ex_el_array=>create(
*                  rows = VALUE #( ( columns_of_row = VALUE #( ( lcl_xlom__ex_el_number=>create( 1 ) )
*                                                              ( lcl_xlom__ex_el_number=>create( 2 ) ) ) )
*                                  ( columns_of_row = VALUE #( ( lcl_xlom__ex_el_number=>create( 3 ) )
*                                                              ( lcl_xlom__ex_el_number=>create( 4 ) ) ) ) ) ) ).
*  ENDMETHOD.
*
*  METHOD assert_equals.
*    cl_abap_unit_assert=>assert_true( xsdbool( exp->is_equal( act ) ) ).
*  ENDMETHOD.
*
*  METHOD function_argument_minus_unary.
*    DATA(act) = parse( tokens = VALUE #( ( value = `OFFSET` type = c_type-function_name )
*                                         ( value = `B2`     type = c_type-text_literal )
*                                         ( value = `,`      type = c_type-comma )
*                                         ( value = `-`      type = c_type-operator )
*                                         ( value = `1`      type = c_type-number )
*                                         ( value = `,`      type = c_type-comma )
*                                         ( value = `-`      type = c_type-operator )
*                                         ( value = `1`      type = c_type-number )
*                                         ( value = `)`      type = c_type-parenthesis_close ) ) ).
*    assert_equals(
*        act = act
*        exp = lcl_xlom__ex_fu_offset=>create(
*                  reference = lcl_xlom__ex_el_string=>create( text = 'B2' )
*                  rows      = lcl_xlom__ex_op_minus_unry=>create( operand = lcl_xlom__ex_el_number=>create( 1 ) )
*                  cols      = lcl_xlom__ex_op_minus_unry=>create(
*                                  operand = lcl_xlom__ex_el_number=>create( 1 ) ) ) ).
*  ENDMETHOD.
*
*  METHOD function_function.
*    DATA(act) = parse( tokens = VALUE #( ( value = `LEN`   type = c_type-function_name )
*                                         ( value = `RIGHT` type = c_type-function_name )
*                                         ( value = `text`  type = c_type-text_literal )
*                                         ( value = `,`     type = c_type-comma )
*                                         ( value = `2`     type = c_type-number )
*                                         ( value = `)`     type = c_type-parenthesis_close )
*                                         ( value = `)`     type = c_type-parenthesis_close ) ) ).
*    assert_equals(
*        act = act
*        exp = lcl_xlom__ex_fu_len=>create( text = lcl_xlom__ex_fu_right=>create(
*                                                            text      = lcl_xlom__ex_el_string=>create( 'text' )
*                                                            num_chars = lcl_xlom__ex_el_number=>create( 2 ) ) ) ).
*  ENDMETHOD.
*
*  METHOD function_optional_argument.
*    DATA(act) = parse( tokens = VALUE #( ( value = `RIGHT` type = c_type-function_name )
*                                         ( value = `text`  type = c_type-text_literal )
*                                         ( value = `,`     type = c_type-comma )
*                                         ( value = `)`     type = c_type-parenthesis_close ) ) ).
*    assert_equals( act = act
*                   exp = lcl_xlom__ex_fu_right=>create( text      = lcl_xlom__ex_el_string=>create( 'text' )
*                                                              num_chars = lcl_xlom__ex_el_empty_arg=>create( ) ) ).
*  ENDMETHOD.
*
*  METHOD if.
*    assert_equals( act = parse( tokens = VALUE #( ( value = `IF` type = c_type-function_name )
*                                                  ( value = `1`  type = c_type-number )
*                                                  ( value = `=`  type = c_type-operator )
*                                                  ( value = `1`  type = c_type-number )
*                                                  ( value = `,`  type = ',' )
*                                                  ( value = `0`  type = c_type-number )
*                                                  ( value = `,`  type = ',' )
*                                                  ( value = `1`  type = c_type-number )
*                                                  ( value = `)`  type = ')' ) ) )
*                   exp = lcl_xlom__ex_fu_if=>create(
*                             condition     = lcl_xlom__ex_op_equal=>create(
*                                                 left_operand  = lcl_xlom__ex_el_number=>create( 1 )
*                                                 right_operand = lcl_xlom__ex_el_number=>create( 1 ) )
*                             expr_if_true  = lcl_xlom__ex_el_number=>create( 0 )
*                             expr_if_false = lcl_xlom__ex_el_number=>create( 1 ) ) ).
*  ENDMETHOD.
*
*  METHOD lexe.
*    DATA(lexer) = lcl_xlom__ex_ut_lexer=>create( ).
*    result = lexer->lexe( text ).
*  ENDMETHOD.
*
*  METHOD number.
*    assert_equals( act = parse( tokens = VALUE #( ( value = `25` type = c_type-number ) ) )
*                   exp = lcl_xlom__ex_el_number=>create( 25 ) ).
*
*    assert_equals( act = parse( tokens = VALUE #( ( value = `-` type = c_type-operator )
*                                                  ( value = `1` type = c_type-number ) ) )
*                   exp = lcl_xlom__ex_op_minus_unry=>create( lcl_xlom__ex_el_number=>create( 1 ) ) ).
*  ENDMETHOD.
*
*  METHOD one_plus_one.
*    assert_equals( act = parse( tokens = VALUE #( ( value = `1`  type = c_type-number )
*                                                  ( value = `+`  type = c_type-operator )
*                                                  ( value = `1`  type = c_type-number ) ) )
*                   exp = lcl_xlom__ex_op_plus=>create( left_operand  = lcl_xlom__ex_el_number=>create( 1 )
*                                                           right_operand = lcl_xlom__ex_el_number=>create( 1 ) ) ).
*  ENDMETHOD.
*
*  METHOD operator_function.
*    DATA(act) = parse( tokens = VALUE #( ( value = `1`    type = c_type-number )
*                                         ( value = `+`    type = c_type-operator )
*                                         ( value = `LEN`  type = c_type-function_name )
*                                         ( value = `text` type = c_type-text_literal )
*                                         ( value = `)`    type = c_type-parenthesis_close ) ) ).
*    assert_equals( act = act
*                   exp = lcl_xlom__ex_op_plus=>create(
*                             left_operand  = lcl_xlom__ex_el_number=>create( 1 )
*                             right_operand = lcl_xlom__ex_fu_len=>create(
*                                                 text = lcl_xlom__ex_el_string=>create( 'text' ) ) ) ).
*  ENDMETHOD.
*
*  METHOD operator_function_operator.
*    DATA(act) = parse( tokens = VALUE #( ( value = `1`    type = c_type-number )
*                                         ( value = `+`    type = c_type-operator )
*                                         ( value = `LEN`  type = c_type-function_name )
*                                         ( value = `text` type = c_type-text_literal )
*                                         ( value = `)`    type = c_type-parenthesis_close )
*                                         ( value = `+`    type = c_type-operator )
*                                         ( value = `1`    type = c_type-number ) ) ).
*    assert_equals( act = act
*                   exp = lcl_xlom__ex_op_plus=>create(
*                             left_operand  = lcl_xlom__ex_el_number=>create( 1 )
*                             right_operand = lcl_xlom__ex_op_plus=>create(
*                                                 left_operand  = lcl_xlom__ex_fu_len=>create(
*                                                                     text = lcl_xlom__ex_el_string=>create( 'text' ) )
*                                                 right_operand = lcl_xlom__ex_el_number=>create( 1 ) ) ) ).
*  ENDMETHOD.
*
*  METHOD parentheses_arithmetic.
*    " lexe( '2*(1+3)' )
*    DATA(act) = parse( VALUE #( ( value = `2`  type = c_type-number )
*                                ( value = `*`  type = c_type-operator )
*                                ( value = `(`  type = c_type-parenthesis_open )
*                                ( value = `1`  type = c_type-number )
*                                ( value = `+`  type = c_type-operator )
*                                ( value = `3`  type = c_type-number )
*                                ( value = `)`  type = c_type-parenthesis_close ) ) ).
*    DATA(exp) = lcl_xlom__ex_op_mult=>create( left_operand  = lcl_xlom__ex_el_number=>create( 2 )
*                                                  right_operand = lcl_xlom__ex_op_plus=>create(
*                                                      left_operand  = lcl_xlom__ex_el_number=>create( 1 )
*                                                      right_operand = lcl_xlom__ex_el_number=>create( 3 ) ) ).
*    assert_equals( act = act
*                   exp = exp ).
*  ENDMETHOD.
*
*  METHOD parentheses_arithmetic_complex.
*    " lexe( '2*(1+3*(5+1))' )
*    DATA(act) = parse( tokens = VALUE #( ( value = `2`  type = c_type-number )
*                                         ( value = `*`  type = c_type-operator )
*                                         ( value = `(`  type = c_type-parenthesis_open )
*                                         ( value = `1`  type = c_type-number )
*                                         ( value = `+`  type = c_type-operator )
*                                         ( value = `3`  type = c_type-number )
*                                         ( value = `*`  type = c_type-operator )
*                                         ( value = `(`  type = c_type-parenthesis_open )
*                                         ( value = `5`  type = c_type-number )
*                                         ( value = `+`  type = c_type-operator )
*                                         ( value = `1`  type = c_type-number )
*                                         ( value = `)`  type = c_type-parenthesis_close )
*                                         ( value = `)`  type = c_type-parenthesis_close ) ) ).
*    DATA(exp) = lcl_xlom__ex_op_mult=>create(
*                    left_operand  = lcl_xlom__ex_el_number=>create( 2 )
*                    right_operand = lcl_xlom__ex_op_plus=>create(
*                        left_operand  = lcl_xlom__ex_el_number=>create( 1 )
*                        right_operand = lcl_xlom__ex_op_mult=>create(
*                                            left_operand  = lcl_xlom__ex_el_number=>create( 3 )
*                                            right_operand = lcl_xlom__ex_op_plus=>create(
*                                                left_operand  = lcl_xlom__ex_el_number=>create( 5 )
*                                                right_operand = lcl_xlom__ex_el_number=>create( 1 ) ) ) ) ).
*    assert_equals( act = act
*                   exp = exp ).
*  ENDMETHOD.
*
*  METHOD parse.
*    result = lcl_xlom__ex_ut_parser=>create( )->parse( tokens ).
*  ENDMETHOD.
*
*  METHOD priority.
*    " lexe( '1+2*3' )
*    DATA(act) = parse( VALUE #( ( value = `1`  type = c_type-number )
*                                ( value = `+`  type = c_type-operator )
*                                ( value = `2`  type = c_type-number )
*                                ( value = `*`  type = c_type-operator )
*                                ( value = `3`  type = c_type-number ) ) ).
*    DATA(exp) = lcl_xlom__ex_op_plus=>create( left_operand  = lcl_xlom__ex_el_number=>create( 1 )
*                                                  right_operand = lcl_xlom__ex_op_mult=>create(
*                                                      left_operand  = lcl_xlom__ex_el_number=>create( 2 )
*                                                      right_operand = lcl_xlom__ex_el_number=>create( 3 ) ) ).
*    assert_equals( act = act
*                   exp = exp ).
*  ENDMETHOD.
*
*  METHOD very_long.
*    cl_abap_unit_assert=>fail( msg = 'TO DO' ).
**    DATA(a) = parse(
**        lexe(
**            `IFERROR(IF(C2<>"",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assig` &&
**`ned Attorney, or Sales Team",B2<>"Jimmy Edwards",B2<>"Kathleen McCarthy"),B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assigned Attorney, or Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(VL` &&
**`OOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(C2<>"",VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1` &&
**            `!$A:$B,2,FALSE),"INTAKE TEAM")))))), VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE),"")` ) ).
*  ENDMETHOD.
*ENDCLASS.
*
*
*CLASS ltc_range IMPLEMENTATION.
*  METHOD convert_column_a_xfd_to_number.
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = 'XFD' )
*                                        exp = 16384 ).
*
*    TRY.
*        lcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = 'XFE' ).
*        cl_abap_unit_assert=>fail( msg = 'Exception expected for XFE - Column does not exist' ).
*      CATCH cx_root ##NO_HANDLER.
*    ENDTRY.
*
*    TRY.
*        lcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = 'ZZZZ' ).
*        cl_abap_unit_assert=>fail( msg = 'Exception expected for ZZZZ - Column does not exist' ).
*      CATCH cx_root ##NO_HANDLER.
*    ENDTRY.
*
*    TRY.
*        lcl_xlom_range=>convert_column_a_xfd_to_number( roman_letters = '1' ).
*        cl_abap_unit_assert=>fail( msg = 'Exception expected for 1 - Invalid column ID' ).
*      CATCH cx_root ##NO_HANDLER.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD convert_column_number_to_a_xfd.
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>convert_column_number_to_a_xfd( 16384 )
*                                        exp = 'XFD' ).
*
*    TRY.
*        lcl_xlom_range=>convert_column_number_to_a_xfd( 16385 ).
*        cl_abap_unit_assert=>fail( msg = 'Exception expected for 16385 - Column does not exist' ).
*      CATCH cx_root ##NO_HANDLER.
*    ENDTRY.
*
*    TRY.
*        lcl_xlom_range=>convert_column_number_to_a_xfd( -1 ).
*        cl_abap_unit_assert=>fail( msg = 'Exception expected for -1 - Column does not exist' ).
*      CATCH cx_root ##NO_HANDLER.
*    ENDTRY.
*  ENDMETHOD.
*
*  METHOD decode_range_address_a1_invali.
*    LOOP AT VALUE string_table( ( `:` ) ( `` ) ( `$` ) ( `A` ) ( `A:` ) ( `$$A1` ) ( `A:A1` ) ( `B2:A1` ) ) INTO DATA(address).
*      TRY.
*          lcl_xlom_range=>decode_range_address_a1( address ).
*          cl_abap_unit_assert=>fail( msg = |Exception expected for address "{ address }"| ).
*        CATCH cx_root ##NO_HANDLER.
*      ENDTRY.
*    ENDLOOP.
*  ENDMETHOD.
*
*  METHOD decode_range_address_a1_valid.
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( 'A1' )
*                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1
*                                                                                        row    = 1 )
*                                                                bottom_right = VALUE #( column = 1
*                                                                                        row    = 1 ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( 'A$1' )
*                                        exp = VALUE ty_address( top_left     = VALUE #( column    = 1
*                                                                                        row       = 1
*                                                                                        row_fixed = abap_true )
*                                                                bottom_right = VALUE #( column    = 1
*                                                                                        row       = 1
*                                                                                        row_fixed = abap_true ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( '$A1' )
*                                        exp = VALUE ty_address( top_left     = VALUE #( column       = 1
*                                                                                        column_fixed = abap_true
*                                                                                        row          = 1 )
*                                                                bottom_right = VALUE #( column       = 1
*                                                                                        column_fixed = abap_true
*                                                                                        row          = 1 ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( '$A$1' )
*                                        exp = VALUE ty_address( top_left     = VALUE #( column       = 1
*                                                                                        column_fixed = abap_true
*                                                                                        row          = 1
*                                                                                        row_fixed    = abap_true )
*                                                                bottom_right = VALUE #( column       = 1
*                                                                                        column_fixed = abap_true
*                                                                                        row          = 1
*                                                                                        row_fixed    = abap_true ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( 'A1:B1' )
*                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1
*                                                                                        row    = 1 )
*                                                                bottom_right = VALUE #( column = 2
*                                                                                        row    = 1 ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( 'A:A' )
*                                        exp = VALUE ty_address( top_left     = VALUE #( column = 1 )
*                                                                bottom_right = VALUE #( column = 1 ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( '1:1' )
*                                        exp = VALUE ty_address( top_left     = VALUE #( row = 1 )
*                                                                bottom_right = VALUE #( row = 1 ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( 'Sheet1!A1' )
*                                        exp = VALUE ty_address( worksheet_name = 'Sheet1'
*                                                                top_left       = VALUE #( column = 1
*                                                                                          row    = 1 )
*                                                                bottom_right   = VALUE #( column = 1
*                                                                                          row    = 1 ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( `'Sheet1 (2)'!A1` )
*                                        exp = VALUE ty_address( worksheet_name = 'Sheet1 (2)'
*                                                                top_left       = VALUE #( column = 1
*                                                                                          row    = 1 )
*                                                                bottom_right   = VALUE #( column = 1
*                                                                                          row    = 1 ) ) ).
*  ENDMETHOD.
*
*  METHOD decode_range_address_sh_invali.
*    LOOP AT VALUE string_table( ( `:` ) ( `` ) ( `$` ) ( `A` ) ( `A:` ) ( `$$A1` ) ( `A:A1` ) ( `B2:A1` ) ) INTO DATA(address).
*      TRY.
*          lcl_xlom_range=>decode_range_address_a1( address ).
*          cl_abap_unit_assert=>fail( msg = |Exception expected for address "{ address }"| ).
*        CATCH cx_root ##NO_HANDLER.
*      ENDTRY.
*    ENDLOOP.
*  ENDMETHOD.
*
*  METHOD decode_range_address_sh_valid.
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( 'BKPF!A:A' )
*                                        exp = VALUE ty_address( worksheet_name = 'BKPF'
*                                                                top_left       = VALUE #( column = 1 )
*                                                                bottom_right   = VALUE #( column = 1 ) ) ).
*    cl_abap_unit_assert=>assert_equals( act = lcl_xlom_range=>decode_range_address_a1( `'BKPF (2)'!A:A` )
*                                        exp = VALUE ty_address( worksheet_name = 'BKPF (2)'
*                                                                top_left       = VALUE #( column = 1 )
*                                                                bottom_right   = VALUE #( column = 1 ) ) ).
*  ENDMETHOD.
*ENDCLASS.
