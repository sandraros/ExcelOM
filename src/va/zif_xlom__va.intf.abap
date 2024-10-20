INTERFACE zif_xlom__va
  PUBLIC.

  TYPES ty_type TYPE i.

  CONSTANTS:
    BEGIN OF c_type,
      boolean TYPE ty_type VALUE 1,
      array   TYPE ty_type VALUE 2,
      empty   TYPE ty_type VALUE 3,
      error   TYPE ty_type VALUE 4,
      number  TYPE ty_type VALUE 5,
      range   TYPE ty_type VALUE 6,
      string  TYPE ty_type VALUE 7,
    END OF c_type.
  DATA type TYPE ty_type READ-ONLY.

  METHODS get_value
    RETURNING VALUE(result) TYPE REF TO data.

  METHODS is_array
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_boolean
    RETURNING VALUE(result) TYPE abap_bool.

  "! Checks whether the current result has the exact same type as the input result,
  "! and the same values. For instance, lcl_xlom__va_string=>create( '1'
  "! )->is_equal( lcl_xlom__va_string=>create( '1' ) ) is true, but
  "! lcl_xlom__va_number=>create( 1
  "! )->is_equal( lcl_xlom__va_string=>create( '1' ) ) is false.
  METHODS is_equal
    IMPORTING input_result  TYPE REF TO zif_xlom__va
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_error
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_number
    RETURNING VALUE(result) TYPE abap_bool.

  METHODS is_string
    RETURNING VALUE(result) TYPE abap_bool.
ENDINTERFACE.
