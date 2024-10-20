CLASS zcl_xlom__va_string DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__va.

    CLASS-METHODS create
      IMPORTING !string       TYPE csequence
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_string.

    CLASS-METHODS get
      IMPORTING !string       TYPE csequence
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_string.

    METHODS get_string
      RETURNING VALUE(result) TYPE string.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ts_buffer_line,
        string TYPE string,
        object TYPE REF TO zcl_xlom__va_string,
      END OF ts_buffer_line.
    TYPES tt_buffer TYPE HASHED TABLE OF ts_buffer_line WITH UNIQUE KEY string.

    CLASS-DATA buffer TYPE tt_buffer.

    DATA string TYPE string.
ENDCLASS.


CLASS zcl_xlom__va_string IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__va_string( ).
    result->zif_xlom__va~type = zif_xlom__va=>c_type-string.
    result->string            = string.
  ENDMETHOD.

  METHOD get.
    DATA(buffer_line) = REF #( buffer[ string = string ] OPTIONAL ).
    IF buffer_line IS NOT BOUND.
      INSERT VALUE #( string = string
                      object = create( string ) )
             INTO TABLE buffer
             REFERENCE INTO buffer_line.
    ENDIF.
    result = buffer_line->object.
  ENDMETHOD.

  METHOD get_string.
    result = string.
  ENDMETHOD.

  METHOD zif_xlom__va~get_value.
    result = REF #( string ).
  ENDMETHOD.

  METHOD zif_xlom__va~is_array.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_equal.
    result = xsdbool( string = CAST zcl_xlom__va_string( input_result )->get_string( ) ).
  ENDMETHOD.

  METHOD zif_xlom__va~is_error.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_string.
    result = abap_true.
  ENDMETHOD.
ENDCLASS.
