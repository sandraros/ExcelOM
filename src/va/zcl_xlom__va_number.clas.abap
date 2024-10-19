class ZCL_XLOM__VA_NUMBER definition
  public
  final
  create private .

public section.

  interfaces ZIF_XLOM__VA .

  class-methods CREATE
    importing
      !NUMBER type F
    returning
      value(RESULT) type ref to ZCL_XLOM__VA_NUMBER .
  class-methods GET
    importing
      !NUMBER type F
    returning
      value(RESULT) type ref to ZCL_XLOM__VA_NUMBER .
  methods GET_INTEGER
    returning
      value(RESULT) type I .
  methods GET_NUMBER
    returning
      value(RESULT) type F .
protected section.
private section.

  types:
    BEGIN OF ts_buffer_line,
        number TYPE f,
        object TYPE REF TO zcl_xlom__va_number,
      END OF ts_buffer_line .
  types:
    tt_buffer TYPE SORTED TABLE OF ts_buffer_line WITH UNIQUE KEY number .

  class-data BUFFER type TT_BUFFER .
  data NUMBER type F .
ENDCLASS.



CLASS ZCL_XLOM__VA_NUMBER IMPLEMENTATION.


  method CREATE.

    result = NEW ZCL_xlom__va_number( ).
    result->ZIF_xlom__va~type = ZIF_xlom__va=>c_type-number.
    result->number                  = number.

  endmethod.


  method GET.

    DATA(buffer_line) = REF #( buffer[ number = number ] OPTIONAL ).
    IF buffer_line IS NOT BOUND.
      result = create( number ).
      INSERT VALUE #( number = number
                      object = result )
             INTO TABLE buffer
             REFERENCE INTO buffer_line.
    ENDIF.
    result = buffer_line->object.

  endmethod.


  method GET_INTEGER.

    " Excel rounding (1.99 -> 1, -1.99 -> -1)
    result = floor( number ).

  endmethod.


  method GET_NUMBER.

    result = number.

  endmethod.


  method ZIF_XLOM__VA~GET_VALUE.

    result = REF #( number ).

  endmethod.


  method ZIF_XLOM__VA~IS_ARRAY.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_BOOLEAN.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_EQUAL.

    IF input_result->type = ZIF_xlom__va=>c_type-number.
      DATA(input_number) = CAST ZCL_xlom__va_number( input_result ).
      IF number = input_number->number.
        result = abap_true.
      ELSE.
        result = abap_false.
      ENDIF.
    ELSE.
      result = abap_false.
    ENDIF.

  endmethod.


  method ZIF_XLOM__VA~IS_ERROR.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_NUMBER.

    result = abap_true.

  endmethod.


  method ZIF_XLOM__VA~IS_STRING.

    result = abap_false.

  endmethod.
ENDCLASS.
