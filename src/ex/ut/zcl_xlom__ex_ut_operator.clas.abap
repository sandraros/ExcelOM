"! Calculation operators and precedence in Excel
"! https://support.microsoft.com/en-us/office/calculation-operators-and-precedence-in-excel-48be406d-4975-4d31-b2b8-7af9e0e2878a
CLASS zcl_xlom__ex_ut_operator DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    TYPES tt_operand_position_offset TYPE STANDARD TABLE OF i WITH EMPTY KEY.
    TYPES tt_expression              TYPE STANDARD TABLE OF REF TO zif_xlom__ex WITH EMPTY KEY.

    CLASS-METHODS class_constructor.

    CLASS-METHODS create
      IMPORTING !name                    TYPE string
                unary                    TYPE abap_bool
                operand_position_offsets TYPE tt_operand_position_offset
                !priority                TYPE i
                !description             TYPE csequence
      RETURNING VALUE(result)            TYPE REF TO zcl_xlom__ex_ut_operator.

    CLASS-METHODS get
      IMPORTING operator      TYPE string
                unary         TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_ut_operator.

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
    METHODS get_priority
      RETURNING VALUE(result) TYPE i.

    "! 1 : predecessor operand only (% e.g. 10%)
    "! 2 : before and after operand only (+ - * / ^ & e.g. 1+1)
    "! 3 : successor operand only (unary + and - e.g. +5)
    METHODS get_operand_position_offsets
      RETURNING VALUE(result) TYPE tt_operand_position_offset.

  PRIVATE SECTION.
    TYPES:
      "! operator precedence
      "! Get operator priorities
      BEGIN OF ts_operator,
        name                     TYPE string,
        "! +1 for unary operators (e.g. -1)
        "! -1 and +1 for binary operators (e.g. 1*2)
        "! -1 for postfix operators (e.g. 10%)
        operand_position_offsets TYPE tt_operand_position_offset,
        "! To distinguish unary from binary operators + and -
        unary                    TYPE abap_bool,
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
        priority                 TYPE i,
        desc                     TYPE string,
        handler                  TYPE REF TO zcl_xlom__ex_ut_operator,
      END OF ts_operator.
    TYPES tt_operator TYPE SORTED TABLE OF ts_operator WITH UNIQUE KEY name unary.

    CLASS-DATA operators TYPE tt_operator.

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
    "! Unary operators are + and - (like in --A1 or +5)
    DATA unary                    TYPE abap_bool.
ENDCLASS.


CLASS zcl_xlom__ex_ut_operator IMPLEMENTATION.
  METHOD class_constructor.
    " Calculation operators and precedence in Excel
    " https://support.microsoft.com/en-us/office/calculation-operators-and-precedence-in-excel-48be406d-4975-4d31-b2b8-7af9e0e2878a
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
              unary                    = operator->unary
              operand_position_offsets = operator->operand_position_offsets
              priority                 = operator->priority
              description              = operator->desc ).
    ENDLOOP.
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_ut_operator( ).
    result->name                     = name.
    result->operand_position_offsets = operand_position_offsets.
    result->priority                 = priority.
    result->unary                    = unary.
    INSERT VALUE #( name     = name
                    unary    = unary
                    priority = priority
                    desc     = description
                    handler  = result )
           INTO TABLE operators.
  ENDMETHOD.

  METHOD get.
    result = VALUE #( operators[ name  = operator
                                 unary = unary ]-handler OPTIONAL ).
    IF result IS NOT BOUND.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
  ENDMETHOD.

  METHOD get_operand_position_offsets.
    result = operand_position_offsets.
  ENDMETHOD.

  METHOD get_priority.
    result = priority.
  ENDMETHOD.
ENDCLASS.
