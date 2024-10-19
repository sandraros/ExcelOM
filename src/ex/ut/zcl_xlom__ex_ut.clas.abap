CLASS zcl_xlom__ex_ut DEFINITION
  PUBLIC
  CREATE PUBLIC.

  PUBLIC SECTION.
    CLASS-METHODS are_equal
      IMPORTING expression_1  TYPE REF TO zif_xlom__ex
                expression_2  TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE abap_bool.

  PROTECTED SECTION.

  PRIVATE SECTION.
ENDCLASS.


CLASS zcl_xlom__ex_ut IMPLEMENTATION.
  METHOD are_equal.
    result = xsdbool(    (     expression_1 IS NOT BOUND
                           AND expression_2 IS NOT BOUND )
                      OR (     expression_1 IS BOUND
                           AND expression_2 IS BOUND
                           AND expression_1->is_equal( expression_2 ) ) ).
  ENDMETHOD.
ENDCLASS.
