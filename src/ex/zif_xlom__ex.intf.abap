"! Any kind of expression
INTERFACE zif_xlom__ex
  PUBLIC.


  TYPES ty_expression_type TYPE i.
  TYPES:
    BEGIN OF ts_operand_result,
      name                     TYPE string,
      object                   TYPE REF TO zif_xlom__va,
      "! <ul>
      "! <li>'X': the argument isn't changed when the formula is expanded for Array Evaluation
      "! e.g. the argument Array of the function INDEX: if A1 contains =INDEX(C1:D2,{1,2},{1,2}),
      "! A2 and A2 values are the same as if they contain =INDEX(C1:D2,1,1) and =INDEX(C1:D2,2,2).</li>
      "! <li>' ': the argument is changed when the formula is expanded for Array Evaluation
      "! e.g.the argument Text of the function RIGHT: if A1 contains =RIGHT(A1:A2,{1;2}),
      "! A1 and A2 values are the same as if they contain =RIGHT(A1,1) and =RIGHT(A2,2).</li>
      "! </ul>
      not_part_of_result_array TYPE abap_bool,
    END OF ts_operand_result.
  TYPES tt_operand_result TYPE SORTED TABLE OF ts_operand_result WITH UNIQUE KEY name.
  TYPES:
    BEGIN OF ts_operand_expr,
      name                     TYPE string,
      object                   TYPE REF TO zif_xlom__ex,
      not_part_of_result_array TYPE abap_bool,
    END OF ts_operand_expr.
  TYPES tt_operand_expr TYPE SORTED TABLE OF ts_operand_expr WITH UNIQUE KEY name.
  TYPES:
    BEGIN OF ts_evaluate_array_operands,
      result          TYPE REF TO zif_xlom__va,
      operand_results TYPE tt_operand_result,
    END OF ts_evaluate_array_operands.

  CONSTANTS:
    "! Used to replace IS INSTANCE OF, in case one day the code is backported to ABAP before 7.50.
    BEGIN OF c_type,
      array          TYPE ty_expression_type VALUE 1,
      boolean        TYPE ty_expression_type VALUE 2,
      empty_argument TYPE ty_expression_type VALUE 3,
      error          TYPE ty_expression_type VALUE 4,
      number         TYPE ty_expression_type VALUE 5,
      range          TYPE ty_expression_type VALUE 6,
      string         TYPE ty_expression_type VALUE 7,
      BEGIN OF function,
        address  TYPE ty_expression_type VALUE 100,
        cell     TYPE ty_expression_type VALUE 101,
        countif  TYPE ty_expression_type VALUE 102,
        find     TYPE ty_expression_type VALUE 103,
        if       TYPE ty_expression_type VALUE 104,
        iferror  TYPE ty_expression_type VALUE 105,
        index    TYPE ty_expression_type VALUE 106,
        indirect TYPE ty_expression_type VALUE 107,
        len      TYPE ty_expression_type VALUE 108,
        match    TYPE ty_expression_type VALUE 109,
        offset   TYPE ty_expression_type VALUE 110,
        right    TYPE ty_expression_type VALUE 111,
        row      TYPE ty_expression_type VALUE 112,
        t        TYPE ty_expression_type VALUE 113,
      END OF function,
      BEGIN OF operation,
        ampersand   TYPE ty_expression_type VALUE 10,
        equal       TYPE ty_expression_type VALUE 11,
        minus       TYPE ty_expression_type VALUE 12,
        minus_unary TYPE ty_expression_type VALUE 13,
        mult        TYPE ty_expression_type VALUE 14,
        plus        TYPE ty_expression_type VALUE 15,
      END OF operation,
    END OF c_type.
  DATA type                 TYPE ty_expression_type  READ-ONLY.
  DATA result_of_evaluation TYPE REF TO zif_xlom__va READ-ONLY.

  METHODS evaluate
    IMPORTING !context      TYPE REF TO zcl_xlom__ex_ut_eval_context
    RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

  METHODS evaluate_single
    IMPORTING arguments     TYPE tt_operand_result
              !context      TYPE REF TO zcl_xlom__ex_ut_eval_context
    RETURNING VALUE(result) TYPE REF TO zif_xlom__va.

  METHODS is_equal
    IMPORTING expression    TYPE REF TO zif_xlom__ex
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS set_result
    IMPORTING !value        TYPE REF TO zif_xlom__va
    RETURNING VALUE(result) TYPE REF TO zif_xlom__va.
ENDINTERFACE.
