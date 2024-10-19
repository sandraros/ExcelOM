class ZCL_XLOM__VA_STRING definition
  public
  final
  create private
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

public section.

  interfaces ZIF_XLOM__VA .

  class-methods CREATE
    importing
      !STRING type CSEQUENCE
    returning
      value(RESULT) type ref to ZCL_XLOM__VA_STRING .
  class-methods GET
    importing
      !STRING type CSEQUENCE
    returning
      value(RESULT) type ref to ZCL_XLOM__VA_STRING .
  methods GET_STRING
    returning
      value(RESULT) type STRING .
protected section.
private section.

  types:
    BEGIN OF ts_buffer_line,
        string TYPE string,
        object TYPE REF TO zcl_xlom__va_string,
      END OF ts_buffer_line .
  types:
    tt_buffer TYPE HASHED TABLE OF ts_buffer_line WITH UNIQUE KEY string .

  class-data BUFFER type TT_BUFFER .
  data STRING type STRING .
ENDCLASS.



CLASS ZCL_XLOM__VA_STRING IMPLEMENTATION.


  method CREATE.

    result = NEW ZCL_xlom__va_string( ).
    result->ZIF_xlom__va~type = ZIF_xlom__va=>c_type-string.
    result->string                  = string.

  endmethod.


  method GET.

    DATA(buffer_line) = REF #( buffer[ string = string ] OPTIONAL ).
    IF buffer_line IS NOT BOUND.
      INSERT VALUE #( string = string
                      object = create( string ) )
             INTO TABLE buffer
             REFERENCE INTO buffer_line.
    ENDIF.
    result = buffer_line->object.

  endmethod.


  method GET_STRING.

    result = string.

  endmethod.


  method ZIF_XLOM__VA~GET_VALUE.

    result = REF #( string ).

  endmethod.


  method ZIF_XLOM__VA~IS_ARRAY.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_BOOLEAN.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_EQUAL.

    result = xsdbool( string = CAST ZCL_xlom__va_string( input_result )->get_string( ) ).

  endmethod.


  method ZIF_XLOM__VA~IS_ERROR.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_NUMBER.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_STRING.

    result = abap_true.

  endmethod.
ENDCLASS.
