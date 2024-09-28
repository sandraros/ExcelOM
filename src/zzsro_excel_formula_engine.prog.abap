REPORT zzsro_excel_formula_engine.

* formula with operation on arrays        result
*
* -> if the left operand is one line high, the line is replicated till max lines of the right operand.
* -> if the left operand is one column wide, the column is replicated till max columns of the right operand.
* -> if the right operand is one line high, the line is replicated till max lines of the left operand.
* -> if the right operand is one column wide, the column is replicated till max columns of the left operand.
*
* -> if the left operand has less lines than the right operand, additional lines are added with #N/A.
* -> if the left operand has less columns than the right operand, additional columns are added with #N/A.
* -> if the right operand has less lines than the left operand, additional lines are added with #N/A.
* -> if the right operand has less columns than the left operand, additional columns are added with #N/A.
*
* -> target array size = max lines of both operands + max columns of both operands.
* -> each target cell of the target array is calculated like this:
*    T(1,1) = L(1,1) op R(1,1)
*    T(2,1) = L(2,1) op R(2,1)
*    etc.
*    If the left cell or right cell is #N/A, the target cell is also #N/A.
*
* Examples where one of the two operands has 1 cell, 1 line or 1 column
*
* a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
*
* a | b | c   op   k                      a op k | b op k | c op k
* d | e | f                               d op k | e op k | f op k
* g | h | i                               g op k | h op k | i op k
*
* a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
* d | e | f                               d op k | e op l | f op m | #N/A
* g | h | i                               g op k | h op l | i op m | #N/A
*
* a | b | c   op   k                      a op k | b op k | c op k
* d | e | f        l                      d op l | e op l | f op l
* g | h | i        m                      g op m | h op m | i op m
*                  n                      #N/A   | #N/A   | #N/A
*
* a | b | c   op   k                      a op k | b op k | c op k
* d | e | f        l                      d op l | e op l | f op l
* g | h | i                               #N/A   | #N/A   | #N/A
*
* a | b | c   op   k                      a op k | b op k | c op k
*                  l                      a op l | b op l | c op l
*                  m                      a op m | b op m | c op m
*
* Both operands have more than 1 line and more than 1 column
*
* a | b | c   op   k | n                  a op k | b op n | #N/A
* d | e | f        l | o                  d op l | e op o | #N/A
* g | h | i                               #N/A   | #N/A   | #N/A
*
* a | b | c   op   k | n                  a op k | b op n | #N/A
* d | e | f        l | o                  d op l | e op o | #N/A
*                  m | p                  #N/A   | #N/A   | #N/A
*
* a | b       op   k | n | q              a op k | b op n | #N/A
* d | e            l | o | r              d op l | e op o | #N/A
* g | h                                   #N/A   | #N/A   | #N/A

CLASS lcx_excelom_expr_parser DEFINITION DEFERRED.
CLASS lcx_excelom_to_do DEFINITION DEFERRED.
CLASS lcx_excelom_unexpected DEFINITION DEFERRED.
CLASS lcl_excelom DEFINITION DEFERRED.
CLASS lcl_excelom_error_value DEFINITION DEFERRED.
CLASS lcl_excelom_expr_array DEFINITION DEFERRED.
CLASS lcl_excelom_expr_expressions DEFINITION DEFERRED.
CLASS lcl_excelom_expr_function_call DEFINITION DEFERRED.
CLASS lcl_excelom_expr_number DEFINITION DEFERRED.
CLASS lcl_excelom_expr_plus DEFINITION DEFERRED.
CLASS lcl_excelom_expr_string DEFINITION DEFERRED.
CLASS lcl_excelom_expr_sub_expr DEFINITION DEFERRED.
CLASS lcl_excelom_expr_table DEFINITION DEFERRED.
CLASS lcl_excelom_formula2 DEFINITION DEFERRED.
CLASS lcl_excelom_range DEFINITION DEFERRED.
CLASS lcl_excelom_range_value DEFINITION DEFERRED.
CLASS lcl_excelom_result_array DEFINITION DEFERRED.
CLASS lcl_excelom_result_error DEFINITION DEFERRED.
CLASS lcl_excelom_result_number DEFINITION DEFERRED.
CLASS lcl_excelom_result_string DEFINITION DEFERRED.
CLASS lcl_excelom_tools DEFINITION DEFERRED.
CLASS lcl_excelom_workbook DEFINITION DEFERRED.
CLASS lcl_excelom_workbooks DEFINITION DEFERRED.
CLASS lcl_excelom_worksheet DEFINITION DEFERRED.
CLASS lcl_excelom_worksheets DEFINITION DEFERRED.
INTERFACE lif_excelom_expr_expression DEFERRED.
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


INTERFACE lif_excelom_result.
  TYPES ty_type TYPE i.
  CONSTANTS:
    BEGIN OF c_type,
      number TYPE ty_type VALUE 1,
      string TYPE ty_type VALUE 2,
      array  TYPE ty_type VALUE 3,
      error  TYPE ty_type VALUE 4,
    END OF c_type.
  DATA type TYPE ty_type READ-ONLY.
  data row_count type i READ-ONLY.
  data column_count type i READ-ONLY.
  methods get_cell_value
  IMPORTING column_offset TYPE I
            row_offset    TYPE I
            RETURNING VALUE(result) type ref to lif_excelom_result.
  methods set_cell_value
  IMPORTING column_offset TYPE I
            row_offset    TYPE I
            value type ref to lif_excelom_result.
ENDINTERFACE.


CLASS lcl_excelom_result_array DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    CLASS-METHODS create_from_range
      IMPORTING range       TYPE ref to lcl_excelom_range
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_result_array.
    CLASS-METHODS create_initial
      IMPORTING
        number_of_rows    TYPE i
        number_of_columns TYPE i
      RETURNING
        value(result)     TYPE REF TO lcl_excelom_result_array.

  PRIVATE SECTION.
    DATA number_of_rows    TYPE i.
    DATA number_of_columns TYPE i.
*    DATA range TYPE ref to lcl_excelom_range.
ENDCLASS.


CLASS lcl_excelom_result_error DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_result.

    TYPES ty_error_number TYPE i.

*    TYPES:
*      BEGIN OF ty_errors,
*        "! #BLOCKED!
*        blocked                    TYPE ty_error_number,
*        "! #CALC!
*        calc                       TYPE ty_error_number,
*        "! #CONNECT!
*        connect                    TYPE ty_error_number,
*        "! #DIV/0!
*        "! Est produit par exemple par: <ul>
*        "! <li>=1/0</li>
*        "! </ul>
*        division_by_zero           TYPE ty_error_number,
*        "! #FIELD!
*        field                      TYPE ty_error_number,
*        "! #GETTING_DATA!
*        getting_data               TYPE ty_error_number,
*        "! #N/A
*        "! Est produit par exemple par: <ul>
*        "! <li>=ERROR.TYPE(1)</li>
*        "! <li>C1 contains =A1:A2+B1:B3 -> C3=#N/A</li>
*        "! </ul>
*        na_not_applicable          TYPE ty_error_number,
*        "! #NAME?
*        "! Est produit par exemple par: <ul>
*        "! <li>=abc</li>
*        "! </ul>
*        name                       TYPE ty_error_number,
*        "! #NULL!
*        null                     TYPE ty_error_number,
*        "! #NUM!
*        "! Est produit par exemple par: <ul>
*        "! <li>=1E+240*1E+240</li>
*        "! </ul>
*        num                        TYPE ty_error_number,
*        "! #PYTHON!
*        python                     TYPE ty_error_number,
*        "! #REF!
*        ref                        TYPE ty_error_number,
*        "! #SPILL!
*        spill                      TYPE ty_error_number,
*        "! #UNKNOWN!
*        unknown                    TYPE ty_error_number,
*        "! #VALUE!
*        "! Est produit par exemple par: <ul>
*        "! <li>=1+"a"</li>
*        "! </ul>
*        value_cannot_be_calculated TYPE ty_error_number,
*      END OF ty_errors.

*    "! https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/cell-error-values
*    "! You can insert a cell error value into a cell or test the value of a cell for an error value by
*    "! using the CVErr function. The cell error values can be one of the following xlCVError constants.
*    "! NB: many errors are missing, the list of the other errors can be found in xlCVError.
*    "! <ul>
*    "! <li>Constant . .Error number . .Cell error value</li>
*    "! <li>xlErrDiv0 . 2007 . . . . . .#DIV/0!         </li>
*    "! <li>xlErrNA . . 2042 . . . . . .#N/A            </li>
*    "! <li>xlErrName . 2029 . . . . . .#NAME?          </li>
*    "! <li>xlErrNull . 2000 . . . . . .#NULL!          </li>
*    "! <li>xlErrNum . .2036 . . . . . .#NUM!           </li>
*    "! <li>xlErrRef . .2023 . . . . . .#REF!           </li>
*    "! <li>xlErrValue .2015 . . . . . .#VALUE!         </li>
*    "! </ul>
*    "! VB example:
*    "! <ul>
*    "! <li>If IsError(ActiveCell.Value) Then            </li>
*    "! <li>. If ActiveCell.Value = CVErr(xlErrDiv0) Then</li>
*    "! <li>. End If                                     </li>
*    "! <li>End If                                       </li>
*    "! </ul>
*    "! NB:
*    "! <ul>
*    "! <li>CVErr(xlErrDiv0) is of type Variant/Error and Locals/Watches shows: Error 2007</li>
*    "! <li>There is no Error data type, only Variant can be used.                        </li>
*    "! </ul>
*    CLASS-DATA internal_error_numbers type ty_errors READ-ONLY.
*    "! =ERROR.TYPE(#DIV/0!) -> 2
*    CLASS-DATA formula_error_numbers TYPE ty_errors READ-ONLY.

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
    DATA error_name TYPE string.
    DATA internal_error_number TYPE ty_error_number.
    DATA formula_error_number  TYPE ty_error_number.

    CLASS-METHODS create
      IMPORTING error_name type string
                internal_error_number TYPE ty_error_number
                formula_error_number TYPE ty_error_number
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


INTERFACE lif_excelom_expr_expression.
  METHODS evaluate RETURNING VALUE(result) TYPE REF TO lif_excelom_result.
ENDINTERFACE.


CLASS lcl_excelom_expr_array DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_array.
ENDCLASS.


CLASS lcl_excelom_expr_expressions DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    METHODS append IMPORTING expression TYPE REF TO lif_excelom_expr_expression.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_expressions.
ENDCLASS.


CLASS lcl_excelom_expr_function_call DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.

    CLASS-METHODS create
      IMPORTING !name         TYPE csequence
                arguments     TYPE REF TO lcl_excelom_expr_expressions
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_function_call.
ENDCLASS.


CLASS lcl_excelom_expr_helper DEFINITION.
  PUBLIC SECTION.
    TYPES:
      BEGIN OF ts_evaluate_array_operands,
        result        TYPE REF TO lif_excelom_result,
        left_operand  TYPE REF TO lif_excelom_result,
        right_operand TYPE REF TO lif_excelom_result,
      END OF ts_evaluate_array_operands.
    CLASS-METHODS evaluate_array_operands
      IMPORTING expression2    TYPE REF TO lif_excelom_expr_expression
                left_operand  TYPE REF TO lif_excelom_expr_expression
                right_operand TYPE REF TO lif_excelom_expr_expression
      RETURNING VALUE(result) TYPE ts_evaluate_array_operands.
ENDCLASS.


CLASS lcl_excelom_expr_lexer DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    TYPES:
      BEGIN OF ts_token,
        value TYPE string,
        type  TYPE c LENGTH 1,
      END OF ts_token.
    TYPES tt_token TYPE STANDARD TABLE OF ts_token WITH EMPTY KEY.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_lexer.

    METHODS lexe IMPORTING !text         TYPE csequence
                 RETURNING VALUE(result) TYPE tt_token.

  PRIVATE SECTION.
    METHODS complete_with_non_matches
      IMPORTING i_string  TYPE string
      CHANGING  c_matches TYPE match_result_tab.
ENDCLASS.


CLASS lcl_excelom_expr_number DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.

    CLASS-METHODS create
      IMPORTING !number       TYPE f
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_number.

  PRIVATE SECTION.
    DATA number TYPE f.
ENDCLASS.


INTERFACE lif_excelom_expr_operator.
  METHODS get_priority RETURNING VALUE(result) TYPE i.
  "! 1 : predecessor operand only (% e.g. 10%)
  "! 2 : before and after operand only (+ - * / ^ & e.g. 1+1)
  "! 3 : successor operand only (unary + and - e.g. +5)
  "!
  "! @parameter result | /
  METHODS get_operand_positions RETURNING VALUE(result) TYPE i.

  METHODS set_operands IMPORTING predecessor TYPE REF TO lif_excelom_expr_expression OPTIONAL
                                 successor   TYPE REF TO lif_excelom_expr_expression OPTIONAL.

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


CLASS lcl_excelom_expr_parser DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS create
      IMPORTING formula_cell  TYPE REF TO lcl_excelom_range
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_parser.

    METHODS parse IMPORTING lexer_tokens  TYPE lcl_excelom_expr_lexer=>tt_token
                  RETURNING VALUE(result) TYPE REF TO lif_excelom_expr_expression
                  RAISING   lcx_excelom_expr_parser.

  PRIVATE SECTION.
    DATA formula_cell        TYPE REF TO lcl_excelom_range.
    DATA formula_offset      TYPE i.
    DATA current_token_index TYPE sytabix.
    DATA tokens              TYPE lcl_excelom_expr_lexer=>tt_token.

    METHODS get_token
      RETURNING VALUE(result) TYPE string.

    METHODS parse_expression         RETURNING VALUE(result) TYPE REF TO lif_excelom_expr_expression.

    METHODS parse_function_arguments RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_expressions.

    METHODS parse_tokens_up_to IMPORTING stop_at_token TYPE csequence
                               RETURNING VALUE(result) TYPE string_table.

    METHODS skip_spaces.

ENDCLASS.


CLASS lcl_excelom_expr_plus DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.
    INTERFACES lif_excelom_expr_operator.

    CLASS-METHODS create
      IMPORTING left_operand  TYPE REF TO lif_excelom_expr_expression
                right_operand TYPE REF TO lif_excelom_expr_expression
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_plus.

  PRIVATE SECTION.
    DATA left_operand  TYPE REF TO lif_excelom_expr_expression.
    DATA right_operand TYPE REF TO lif_excelom_expr_expression.
ENDCLASS.


CLASS lcl_excelom_expr_string DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.

    CLASS-METHODS create
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_string.
ENDCLASS.


CLASS lcl_excelom_expr_sub_expr DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_expr_sub_expr.
ENDCLASS.


CLASS lcl_excelom_expr_table DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.

    CLASS-METHODS create
      IMPORTING table_name            TYPE csequence
                row_column_specifiers TYPE string_table
      RETURNING VALUE(result)         TYPE REF TO lcl_excelom_expr_table.
ENDCLASS.


CLASS lcl_excelom_tools DEFINITION FINAL.

  PUBLIC SECTION.
    CLASS-METHODS type
      IMPORTING any_data_object TYPE any
      RETURNING VALUE(result)   TYPE abap_typekind.

ENDCLASS.


CLASS lcl_excelom_cell_format DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_cell_format.

  PRIVATE SECTION.
    CLASS-DATA singleton TYPE REF TO lcl_excelom_error_value.
ENDCLASS.


CLASS lcl_excelom_error_value DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CLASS-METHODS get_singleton
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_error_value.

  PRIVATE SECTION.
    CLASS-DATA singleton TYPE REF TO lcl_excelom_error_value.
ENDCLASS.


CLASS lcl_excelom_formula2 DEFINITION FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    METHODS calculate.

    CLASS-METHODS create IMPORTING !range        TYPE REF TO lcl_excelom_range
                         RETURNING VALUE(result) TYPE REF TO lcl_excelom_formula2.

    METHODS set IMPORTING !value TYPE string.

  PRIVATE SECTION.
    DATA range       TYPE REF TO lcl_excelom_range.
    DATA _expression TYPE REF TO lif_excelom_expr_expression.
ENDCLASS.


CLASS lcl_excelom_range DEFINITION FINAL
  CREATE PRIVATE
  FRIENDS lcl_excelom_range_value
          lcl_excelom_result_array
          lcl_excelom_formula2.

  PUBLIC SECTION.
    INTERFACES lif_excelom_expr_expression.

    METHODS calculate.

    "! Called by the Worksheet.Range property.
    "! @parameter cell1  | Required    Variant A String that is a range reference when one argument is used. Either a String that is a range reference or a Range object when two arguments are used.
    "! @parameter cell2  | Optional    Variant Either a String that is a range reference or a Range object. Cell2 defines another extremity of the range returned by the property.
    "! @parameter result | .
    CLASS-METHODS create
      IMPORTING cell1         TYPE REF TO lcl_excelom_range
                cell2         TYPE REF TO lcl_excelom_range OPTIONAL
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.

    CLASS-METHODS create_from_address
      IMPORTING address       TYPE clike
                relative_to   TYPE REF TO lcl_excelom_worksheet
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range.

    METHODS formula2
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_formula2.

    METHODS parent
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

    METHODS value
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_range_value.

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

    CLASS-METHODS decode_range_address
      IMPORTING address       TYPE string
      RETURNING VALUE(result) TYPE ty_address.

    DATA _value    TYPE REF TO lcl_excelom_range_value.
    DATA _formula2 TYPE REF TO lcl_excelom_formula2.
    DATA _address  TYPE ty_address.
    DATA _parent   TYPE REF TO lcl_excelom_worksheet.
ENDCLASS.


CLASS lcl_excelom_range_value DEFINITION FINAL
  CREATE PRIVATE
  FRIENDS lcl_excelom_worksheet.

  PUBLIC SECTION.
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


CLASS lcl_excelom_worksheet DEFINITION FINAL
  CREATE PRIVATE
  FRIENDS lcl_excelom_range_value
          lcl_excelom_range
          lcl_excelom_formula2.

  PUBLIC SECTION.
    TYPES ty_name TYPE string.

    "! Worksheet.Calculate method (Excel).
    "! Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.
    "! <p>expression.Calculate</p>
    "! expression A variable that represents a Worksheet object.
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.calculate(method)
    METHODS calculate.

    CLASS-METHODS create RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheet.

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
    TYPES tt_formula TYPE STANDARD TABLE OF REF TO lcl_excelom_formula2 WITH EMPTY KEY.

    DATA formulas TYPE tt_formula.
    DATA _cells   TYPE tt_cell.
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

    CLASS-METHODS create
      RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbook.

    METHODS worksheets RETURNING VALUE(result) TYPE REF TO lcl_excelom_worksheets.

  PRIVATE SECTION.
    DATA _worksheets TYPE REF TO lcl_excelom_worksheets.
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


CLASS lcl_excelom DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS create RETURNING VALUE(result) TYPE REF TO lcl_excelom.

    METHODS workbooks RETURNING VALUE(result) TYPE REF TO lcl_excelom_workbooks.
    METHODS calculate.

  PRIVATE SECTION.
    DATA _workbooks TYPE REF TO lcl_excelom_workbooks.
ENDCLASS.


CLASS lcl_excelom_cell_format IMPLEMENTATION.
  METHOD create.
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


CLASS lcl_excelom_expr_array IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.
  METHOD lif_excelom_expr_expression~evaluate.

  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_expressions IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_expressions( ).
  ENDMETHOD.

  METHOD append.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_function_call IMPLEMENTATION.
  METHOD create.
    " TODO: parameter NAME is never used (ABAP cleaner)
    " TODO: parameter ARGUMENTS is never used (ABAP cleaner)

    result = NEW lcl_excelom_expr_function_call( ).
  ENDMETHOD.
  METHOD lif_excelom_expr_expression~evaluate.

  ENDMETHOD.
ENDCLASS.


CLASS LCL_excelom_expr_helper IMPLEMENTATION.
  METHOD evaluate_array_operands.
    result-left_operand = left_operand->evaluate( ).
    result-right_operand = right_operand->evaluate( ).
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
        TRY.
        DATA(expression) = lcl_excelom_expr_plus=>create(
            left_operand  = lcl_excelom_expr_number=>create(
                                number = CAST lcl_excelom_result_number( left_operand_result_one_cell )->get_number( ) )
            right_operand = lcl_excelom_expr_number=>create(
                number = CAST lcl_excelom_result_number( right_operand_result_one_cell )->get_number( ) ) ).
        catch cx_sy_move_cast_error ##NO_HANDLER.
        endtry.

        DATA(target_array_result_one_cell) = cond #( when expression is bound
        then expression->lif_excelom_expr_expression~evaluate( )
        ELSE lcl_excelom_result_error=>na_not_applicable ).

        target_array->lif_excelom_result~set_cell_value(
            row_offset    = row_offset
            column_offset = column_offset
            value         = target_array_result_one_cell ).

        column_offset = column_offset + 1.
      ENDDO.
      row_offset = row_offset + 1.
    ENDDO.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_lexer IMPLEMENTATION.
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
    result = NEW lcl_excelom_expr_lexer( ).
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


CLASS lcl_excelom_expr_number IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_number( ).
    result->number = number.
  ENDMETHOD.

  METHOD lif_excelom_expr_expression~evaluate.
    result = lcl_excelom_result_number=>create( number ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_parser IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_parser( ).
    result->formula_cell = formula_cell.
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
        s_token      TYPE REF TO lcl_excelom_expr_lexer=>ts_token,
        o_expression TYPE REF TO lif_excelom_expr_expression,
      END OF ts_expression_part.
    TYPES tt_expression_part TYPE STANDARD TABLE OF ts_expression_part WITH EMPTY KEY.
    TYPES to_expression      TYPE REF TO lif_excelom_expr_expression.

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
          RAISE EXCEPTION TYPE lcx_excelom_unexpected.
        WHEN '{'.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        WHEN ')'.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        WHEN ']'.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        WHEN '}'.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
*          result = lcl_array=>create( ).
        WHEN ` `.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        WHEN ':'.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
        WHEN ','.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
          " end of expression
          EXIT.
        WHEN ';'.
          RAISE EXCEPTION TYPE lcx_excelom_to_do.
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
            expression = lcl_excelom_expr_function_call=>create( name      = token->value
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
            expression = lcl_excelom_expr_table=>create( table_name            = token->value
                                                         row_column_specifiers = row_column_specifiers ).
          ELSEIF substring( val = token->value
                            len = 1 ) CA '0123456789'.
            " The word is a number.
            " NB: the number 0.5 is represented in the formulas with the leading 0 (i.e. "0.5")
            current_token_index = current_token_index + 1.
            expression = lcl_excelom_expr_number=>create( EXACT #( token->value ) ).
          ELSE.
            " The word is a cell reference, name of named range, constant (TRUE, FALSE)
            " TODO
            current_token_index = current_token_index + 1.
            expression = lcl_excelom_range=>create_from_address( address     = token->value
                                                                 relative_to = formula_cell->parent( ) ).
          ENDIF.
        WHEN '"'.
          " Remove double quotes e.g. "say ""hello""" -> say "hello"
          current_token_index = current_token_index + 1.
          expression = lcl_excelom_expr_string=>create( replace( val   = token->value
                                                                 regex = '^"|(")"|"$'
                                                                 with  = '$1'
                                                                 occ   = 0 ) ).
        WHEN OTHERS.
          RAISE EXCEPTION TYPE lcx_excelom_unexpected.
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
          position            TYPE sytabix,
          operator            TYPE string,
          operator_expression TYPE REF TO lif_excelom_expr_expression,
          preceding_operand   TYPE ts_expression_part,
          succeding_operand   TYPE ts_expression_part,
          priority            TYPE i,
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
        DATA(buffer_expression) = VALUE to_expression( ).
        LOOP AT GROUP group_by_priority REFERENCE INTO work_line.
          IF buffer_expression IS NOT BOUND.
            buffer_expression = work_line->preceding_operand-o_expression.
          ENDIF.
          CASE work_line->operator.
            WHEN '+'.
              work_line->operator_expression = lcl_excelom_expr_plus=>create(
                                                   left_operand  = buffer_expression
                                                   right_operand = work_line->succeding_operand-o_expression ).
            WHEN OTHERS.
              RAISE EXCEPTION TYPE lcx_excelom_expr_parser
                EXPORTING text = |Invalid operator "{ work_line->operator }"|.
          ENDCASE.
          buffer_expression = work_line->operator_expression.
        ENDLOOP.
      ENDLOOP.
      result = buffer_expression.
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
  ENDMETHOD.

  METHOD skip_spaces.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_plus IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_expr_plus( ).
    result->left_operand  = left_operand.
    result->right_operand = right_operand.
  ENDMETHOD.

  METHOD lif_excelom_expr_expression~evaluate.
    data(array_evaluation) = LCL_excelom_expr_helper=>evaluate_array_operands( expression2   = me
                                                                               left_operand  = left_operand
                                                                               right_operand = right_operand ).
    if array_evaluation-result is bound.
      result = array_evaluation-result.
    else.
      IF     array_evaluation-left_operand->type  = lif_excelom_result=>c_type-number
         AND array_evaluation-right_operand->type = lif_excelom_result=>c_type-number.
        result = lcl_excelom_result_number=>create(
                     CAST lcl_excelom_result_number( array_evaluation-left_operand )->get_number( )
                      + CAST lcl_excelom_result_number( array_evaluation-left_operand )->get_number( ) ).
      ELSEIF array_evaluation-left_operand->type = lif_excelom_result=>c_type-array.
        DATA(left_array) = CAST lcl_excelom_result_array( array_evaluation-left_operand ).
      ELSEIF array_evaluation-right_operand->type = lif_excelom_result=>c_type-array.
        DATA(right_array) = CAST lcl_excelom_result_array( array_evaluation-right_operand ).
      ELSEIF    array_evaluation-left_operand->type  = lif_excelom_result=>c_type-string
             OR array_evaluation-right_operand->type = lif_excelom_result=>c_type-string.
        result = lcl_excelom_result_error=>value_cannot_be_calculated.
      ENDIF.
    endif.
  ENDMETHOD.

  METHOD lif_excelom_expr_operator~get_priority.
    result = 6.
  ENDMETHOD.

  METHOD lif_excelom_expr_operator~get_operand_positions.
  ENDMETHOD.

  METHOD lif_excelom_expr_operator~set_operands.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_string IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.

  METHOD lif_excelom_expr_expression~evaluate.

  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_sub_expr IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.

  METHOD lif_excelom_expr_expression~evaluate.

  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_expr_table IMPLEMENTATION.
  METHOD create.
  ENDMETHOD.

  METHOD lif_excelom_expr_expression~evaluate.

  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_formula2 IMPLEMENTATION.
  METHOD calculate.
*    _expression->evaluate( ).
  ENDMETHOD.

  METHOD create.
    result = NEW lcl_excelom_formula2( ).
    result->range = range.
    DATA(worksheet) = range->_parent.
    INSERT result INTO TABLE worksheet->formulas.
  ENDMETHOD.

  METHOD set.
    DATA(lexer) = lcl_excelom_expr_lexer=>create( ).
    DATA(lexer_tokens) = lexer->lexe( value ).
    IF lexer_tokens IS INITIAL.
      RETURN.
    ENDIF.
    IF lexer_tokens[ 1 ]-value = '='.
      DELETE lexer_tokens INDEX 1.
    ENDIF.
    DATA(parser) = lcl_excelom_expr_parser=>create( range ).
    _expression = parser->parse( lexer_tokens ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_range IMPLEMENTATION.
  METHOD create.
    IF cell2 IS INITIAL.
      result = cell1.
    ELSE.
      result = NEW lcl_excelom_range( ).
      result->_address = VALUE #( top_left     = VALUE #( column = nmin( val1 = cell1->_address-top_left-column
                                                                         val2 = cell2->_address-top_left-column )
                                                          row    = nmin( val1 = cell1->_address-top_left-column
                                                                         val2 = cell2->_address-top_left-column ) )
                                  bottom_right = VALUE #( column = nmax( val1 = cell1->_address-top_left-column
                                                                         val2 = cell2->_address-top_left-column )
                                                          row    = nmax( val1 = cell1->_address-top_left-column
                                                                         val2 = cell2->_address-top_left-column ) ) ).
    ENDIF.
  ENDMETHOD.

  METHOD create_from_address.
    result = NEW lcl_excelom_range( ).
    result->_parent  = relative_to.
    result->_address = decode_range_address( address ).
  ENDMETHOD.

  METHOD decode_range_address.
    " A1 (relative column and row)
    " $A1 (absolute column, relative column)
    " A$1
    " $A$1
    " A1:A2
    " $A$A
    " A:A
    " 1:1
    " Sheet1!A1
    " 'Sheet 1'!A1
    " '[C:\workbook.xlsx]'!NAME' (workbook absolute path / name global scope)
    " '[workbook.xlsx]Sheet 1'!$A$1' (workbook relative path)
    " [1]!NAME (XLSX internal notation for workbooks)
    " [1]Sheet1!$A$3
    IF address = 'A1'.
      result = VALUE #( top_left     = VALUE #( column       = 1
                                                column_fixed = abap_false
                                                row          = 1
                                                row_fixed    = abap_false )
                        bottom_right = VALUE #( column       = 1
                                                column_fixed = abap_false
                                                row          = 1
                                                row_fixed    = abap_false      ) ).
    ELSEIF address = 'A2'.
      result = VALUE #( top_left     = VALUE #( column       = 1
                                                column_fixed = abap_false
                                                row          = 2
                                                row_fixed    = abap_false )
                        bottom_right = VALUE #( column       = 1
                                                column_fixed = abap_false
                                                row          = 2
                                                row_fixed    = abap_false      ) ).
    ELSEIF address = 'B$1'.
      result = VALUE #( top_left     = VALUE #( column       = 2
                                                column_fixed = abap_false
                                                row          = 1
                                                row_fixed    = abap_false )
                        bottom_right = VALUE #( column       = 1
                                                column_fixed = abap_false
                                                row          = 1
                                                row_fixed    = abap_false      ) ).
    ENDIF.
  ENDMETHOD.

  METHOD calculate.
    formula2( )->calculate( ).
  ENDMETHOD.

  METHOD formula2.
    IF _formula2 IS NOT BOUND.
      _formula2 = lcl_excelom_formula2=>create( me ).
    ENDIF.
    result = _formula2.
  ENDMETHOD.

  METHOD parent.
    result = _parent.
  ENDMETHOD.

  METHOD value.
    IF _value IS NOT BOUND.
      _value = lcl_excelom_range_value=>create( me ).
    ENDIF.
    result = _value.
  ENDMETHOD.

  METHOD lif_excelom_expr_expression~evaluate.
    result = lcl_excelom_result_array=>create_from_range( me ).
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_range_value IMPLEMENTATION.
  METHOD create.
    result = NEW lcl_excelom_range_value( ).
    result->range = range.
  ENDMETHOD.

  METHOD set.
    DATA(row) = range->_address-top_left-row.
    WHILE row <= range->_address-bottom_right-row.
      DATA(column) = range->_address-top_left-column.
      WHILE column <= range->_address-bottom_right-column.
        DATA(cell) = REF #( range->_parent->_cells[ column = column
                                                    row    = row ] OPTIONAL ).
        IF cell IS NOT BOUND.
          INSERT VALUE #( column = column
                          row    = row )
                 INTO TABLE range->_parent->_cells
                 REFERENCE INTO cell.
        ENDIF.
        cell->value_type = type.
        CASE type.
          WHEN cl_abap_typedescr=>typekind_float.
            cell->value2-double = value.
          WHEN cl_abap_typedescr=>typekind_string.
            cell->value2-string = value.
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
*    result->number_of_columns = range->rows->get_count( ).
*    result->range = range.
  ENDMETHOD.
  METHOD lif_excelom_result~get_cell_value.
*    if range is bound.
*      range->_address-bottom_right-column = .
*      range->_parent->_
*    endif.
  ENDMETHOD.
  METHOD create_initial.
    result = NEW lcl_excelom_result_array( ).
    result->number_of_rows    = number_of_rows.
    result->number_of_columns = number_of_columns.
  ENDMETHOD.
  METHOD lif_excelom_result~set_cell_value.

  ENDMETHOD.

ENDCLASS.


CLASS lcl_excelom_result_error IMPLEMENTATION.
  METHOD class_constructor.
    " Error instances
    " TODO #PYTHON! internal error number is not 2222, what is it?
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
                                                                   formula_error_number  = 2 ).
    field                      = lcl_excelom_result_error=>create( error_name            = '#FIELD!       '
                                                                   internal_error_number = 2049
                                                                   formula_error_number  = 13 ).
    getting_data               = lcl_excelom_result_error=>create( error_name            = '#GETTING_DATA!'
                                                                   internal_error_number = 2043
                                                                   formula_error_number  = 8 ).
    na_not_applicable          = lcl_excelom_result_error=>create( error_name            = '#N/A          '
                                                                   internal_error_number = 2042
                                                                   formula_error_number  = 7 ).
    name                       = lcl_excelom_result_error=>create( error_name            = '#NAME?        '
                                                                   internal_error_number = 2029
                                                                   formula_error_number  = 5 ).
    null                       = lcl_excelom_result_error=>create( error_name            = '#NULL!        '
                                                                   internal_error_number = 2000
                                                                   formula_error_number  = 1 ).
    num                        = lcl_excelom_result_error=>create( error_name            = '#NUM!         '
                                                                   internal_error_number = 2036
                                                                   formula_error_number  = 6 ).
    python                     = lcl_excelom_result_error=>create( error_name            = '#PYTHON!      '
                                                                   internal_error_number = 2222
                                                                   formula_error_number  = 19 ).
    ref                        = lcl_excelom_result_error=>create( error_name            = '#REF!         '
                                                                   internal_error_number = 2023
                                                                   formula_error_number  = 4 ).
    spill                      = lcl_excelom_result_error=>create( error_name            = '#SPILL!       '
                                                                   internal_error_number = 2045
                                                                   formula_error_number  = 9 ).
    unknown                    = lcl_excelom_result_error=>create( error_name            = '#UNKNOWN!     '
                                                                   internal_error_number = 2048
                                                                   formula_error_number  = 12 ).
    value_cannot_be_calculated = lcl_excelom_result_error=>create( error_name            = '#VALUE!       '
                                                                   internal_error_number = 2015
                                                                   formula_error_number  = 3 ).
*    internal_error_numbers = VALUE ty_errors( blocked                    = 2047
*                                              calc                       = 2050
*                                              connect                    = 2046
*                                              division_by_zero           = 2007
*                                              field                      = 2049
*                                              getting_data               = 2043
*                                              na_not_applicable          = 2042
*                                              name                       = 2029
*                                              null                       = 2000
*                                              num                        = 2036
*                                              " python                     = ????
*                                              ref                        = 2023
*                                              spill                      = 2045
*                                              unknown                    = 2048
*                                              value_cannot_be_calculated = 2015 ).
*    " Formula function "ERROR.TYPE"
*    formula_error_numbers = VALUE ty_errors( blocked                    = 11
*                                             calc                       = 14
*                                             connect                    = 10
*                                             division_by_zero           = 2
*                                             field                      = 13
*                                             getting_data               = 8
*                                             na_not_applicable          = 7
*                                             name                       = 5
*                                             null                       = 1
*                                             num                        = 6
*                                             python                     = 19
*                                             ref                        = 4
*                                             spill                      = 9
*                                             unknown                    = 12
*                                             value_cannot_be_calculated = 3 ).
  ENDMETHOD.
  METHOD create.
    result = NEW lcl_excelom_result_error( ).
    result->lif_excelom_result~type = lif_excelom_result=>c_type-error.
    result->error_name = error_name.
    result->internal_error_number = internal_error_number.
    result->formula_error_number  = formula_error_number.
  ENDMETHOD.
  METHOD get_by_error_number.
  ENDMETHOD.
  METHOD lif_excelom_result~get_cell_value.
  ENDMETHOD.
  METHOD lif_excelom_result~set_cell_value.

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
  METHOD lif_excelom_result~get_cell_value.
  ENDMETHOD.
  METHOD lif_excelom_result~set_cell_value.

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
  METHOD lif_excelom_result~set_cell_value.

  ENDMETHOD.

ENDCLASS.


CLASS lcl_excelom_tools IMPLEMENTATION.
  METHOD type.
    DESCRIBE FIELD any_data_object TYPE result.
  ENDMETHOD.
ENDCLASS.


CLASS lcl_excelom_worksheet IMPLEMENTATION.
  METHOD calculate.
    LOOP AT formulas INTO DATA(formula).
      formula->calculate( ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create.
    result = NEW lcl_excelom_worksheet( ).
  ENDMETHOD.

  METHOD range_from_address.
    DATA(range_1) = lcl_excelom_range=>create_from_address( address     = cell1
                                                            relative_to = me ).
    IF cell2 IS INITIAL.
      result = range_1.
    ELSE.
      DATA(range_2) = lcl_excelom_range=>create_from_address( address     = cell2
                                                              relative_to = me ).
      result = range_from_two_ranges( cell1 = range_1
                                      cell2 = range_2 ).
    ENDIF.
  ENDMETHOD.

  METHOD range_from_two_ranges.
    result = lcl_excelom_range=>create( cell1 = cell1
                                        cell2 = cell2 ).
  ENDMETHOD.
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
    result = _workbooks.
  ENDMETHOD.
ENDCLASS.


CLASS ltc_parser DEFINITION FINAL
  FOR TESTING RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS test   FOR TESTING RAISING cx_static_check.
    METHODS test2  FOR TESTING RAISING cx_static_check.
    METHODS test3  FOR TESTING RAISING cx_static_check.
    METHODS test4  FOR TESTING RAISING cx_static_check.
    METHODS test31 FOR TESTING RAISING cx_static_check.

    TYPES tt_token TYPE lcl_excelom_expr_lexer=>tt_token.

    METHODS lexe
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE lcl_excelom_expr_lexer=>tt_token.

    METHODS parse IMPORTING lexer_tokens  TYPE lcl_excelom_expr_lexer=>tt_token
                  RETURNING VALUE(result) TYPE REF TO lif_excelom_expr_expression.

    METHODS evaluate IMPORTING expression    TYPE REF TO lif_excelom_expr_expression
                     RETURNING VALUE(result) TYPE REF TO lif_excelom_result.

    METHODS get_texts_from_matches IMPORTING i_string      TYPE string
                                             i_matches     TYPE match_result_tab
                                   RETURNING VALUE(result) TYPE string_table.

ENDCLASS.


CLASS ltc_parser IMPLEMENTATION.
  METHOD evaluate.
*    result = expression->evaluate( ).
  ENDMETHOD.

  METHOD get_texts_from_matches.
    LOOP AT i_matches REFERENCE INTO DATA(match).
      APPEND substring( val = i_string
                        off = match->offset
                        len = match->length ) TO result.
    ENDLOOP.
  ENDMETHOD.

  METHOD lexe.
    DATA(lexer) = lcl_excelom_expr_lexer=>create( ).
    result = lexer->lexe( text ).
  ENDMETHOD.

  METHOD parse.
    DATA(parser) = lcl_excelom_expr_parser=>create( lcl_excelom_range=>create_from_address(
                                                        address     = 'A1'
                                                        relative_to = lcl_excelom_worksheet=>create( ) ) ).
    result = parser->parse( lexer_tokens ).
  ENDMETHOD.

  METHOD test.
    cl_abap_unit_assert=>assert_equals( act = lexe( 'IF(1=1,0,1)' )
                                        exp = VALUE tt_token( ( value = `IF` type = 'W' )
                                                              ( value = `(`  type = '(' )
                                                              ( value = `1`  type = 'W' )
                                                              ( value = `=`  type = 'O' )
                                                              ( value = `1`  type = 'W' )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `0`  type = 'W' )
                                                              ( value = `,`  type = ',' )
                                                              ( value = `1`  type = 'W' )
                                                              ( value = `)`  type = ')' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Sheet1!$A$1' )
                                        exp = VALUE tt_token( ( value = `Sheet1!$A$1` type = 'W' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( '"IF(1=1,0,1)"' )
                                        exp = VALUE tt_token( ( value = `"IF(1=1,0,1)"` type = '"' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( '"IF(A1=""X"",0,1)"' )
                                        exp = VALUE tt_token( ( value = `"IF(A1=""X"",0,1)"` type = '"' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( |(a{ repeat( val = ',a'
                                                                 occ = 5000 ) })| )
                                        exp = VALUE tt_token( ( value = `(` type = '(' )
                                                              ( value = `a` type = 'W' )
                                                              ( LINES OF VALUE
                                                                tt_token( FOR i = 1 WHILE i <= 5000
                                                                          ( value = `,` type = ',' )
                                                                          ( value = `a` type = 'W' ) ) )
                                                              ( value = `)` type = ')' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[]' )
                                        exp = VALUE tt_token( ( value = `Table1` type = 'W' )
                                                              ( value = `[`     type = `[` )
                                                              ( value = `]`     type = `]` )  ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( 'Table1[Column1]' )
                                        exp = VALUE tt_token( ( value = `Table1`    type = 'W' )
                                                              ( value = `[Column1]` type = `[` ) ) ).
  ENDMETHOD.

  METHOD test2.
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[[#Headers],[#Data],[% Commission]]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[ [#Headers],[#Data],[% Commission] ]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[[#Headers], [#Data], [% Commission]]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lexe( `DeptSales[ [#Headers], [#Data], [% Commission] ]` )
                                        exp = VALUE tt_token( ( value = `DeptSales`      type = 'W' )
                                                              ( value = `[`              type = `[` )
                                                              ( value = `[#Headers]`     type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[#Data]`        type = `[` )
                                                              ( value = `,`              type = `,` )
                                                              ( value = `[% Commission]` type = `[` )
                                                              ( value = `]`              type = `]` ) ) ).
  ENDMETHOD.

  METHOD test3.
    DATA(result) = evaluate( parse( lexe( `1+1` ) ) ).
    cl_abap_unit_assert=>assert_equals( act = result->type
                                        exp = result->c_type-number ).
    cl_abap_unit_assert=>assert_equals( act = CAST lcl_excelom_result_number( result )->get_number( )
                                        exp = 2 ).
  ENDMETHOD.

  METHOD test31.
    DATA(app) = lcl_excelom=>create( ).
    DATA(workbook) = app->workbooks( )->add( 'name' ).
    DATA(worksheet) = workbook->worksheets( )->item( 'Sheet1' ).
    worksheet->range_from_address( 'A1' )->value( )->set_double( 10 ).
    DATA(range) = worksheet->range_from_address( 'A2' ).
    range->formula2( )->set( '=A1+1' ).
    app->calculate( ).
    cl_abap_unit_assert=>assert_equals( act = range->value( )
                                        exp = 11 ).
  ENDMETHOD.

  METHOD test4.
*    data(a) = parse( lexe(
*`IFERROR(IF(C2<>"",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assig` &&
*`ned Attorney, or Sales Team",B2<>"Jimmy Edwards",B2<>"Kathleen McCarthy"),B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Intake Team, Assigned Attorney, or Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(VL` &&
*`OOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Assigned Attorney",B2,IF(AND(VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE)="Sales Team",OR(B2="Jimmy Edwards",B2="Kathleen McCarthy")),"Sales Team",IF(C2<>"",VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1` &&
*`!$A:$B,2,FALSE),"INTAKE TEAM")))))), VLOOKUP(A2&"",[LPSMatch.xlsx]Sheet1!$A:$B,2,FALSE),"")` ) ).
  ENDMETHOD.
ENDCLASS.
