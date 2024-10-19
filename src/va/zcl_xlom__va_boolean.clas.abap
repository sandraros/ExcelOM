class ZCL_XLOM__VA_BOOLEAN definition
  public
  final
  create private
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

public section.

  interfaces ZIF_XLOM__VA .

  class-data FALSE type ref to ZCL_XLOM__VA_BOOLEAN read-only .
  class-data TRUE type ref to ZCL_XLOM__VA_BOOLEAN read-only .
  data BOOLEAN_VALUE type ABAP_BOOL read-only .

  class-methods CLASS_CONSTRUCTOR .
  class-methods GET
    importing
      !BOOLEAN_VALUE type ABAP_BOOL
    returning
      value(RESULT) type ref to ZCL_XLOM__VA_BOOLEAN .
protected section.
private section.

  data NUMBER type F .

  class-methods CREATE
    importing
      !BOOLEAN_VALUE type ABAP_BOOL
    returning
      value(RESULT) type ref to ZCL_XLOM__VA_BOOLEAN .
ENDCLASS.



CLASS ZCL_XLOM__VA_BOOLEAN IMPLEMENTATION.


  method CLASS_CONSTRUCTOR.

    false = create( abap_false ).
    true  = create( abap_true ).

  endmethod.


  method CREATE.

    result = NEW ZCL_xlom__va_boolean( ).
    result->ZIF_xlom__va~type = ZIF_xlom__va=>c_type-boolean.
    result->boolean_value = boolean_value.
    result->number = COND #( when boolean_value = abap_true then -1 ).

  endmethod.


  method GET.

    result = SWITCH #( boolean_value WHEN abap_true
                                     THEN true
                                     ELSE false ).

  endmethod.


  method ZIF_XLOM__VA~GET_VALUE.

    result = REF #( boolean_value ).

  endmethod.


  method ZIF_XLOM__VA~IS_ARRAY.

    result = abap_true.

  endmethod.


  method ZIF_XLOM__VA~IS_BOOLEAN.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_EQUAL.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.


  method ZIF_XLOM__VA~IS_ERROR.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_NUMBER.

    result = abap_false.

  endmethod.


  method ZIF_XLOM__VA~IS_STRING.

    result = abap_false.

  endmethod.
ENDCLASS.
