class ZCL_XLOM__VA_EMPTY definition
  public
  final
  create private .

public section.

  interfaces ZIF_XLOM__VA .

  class-methods GET_SINGLETON
    returning
      value(RESULT) type ref to ZCL_XLOM__VA_EMPTY .
protected section.
private section.

  class-data SINGLETON type ref to ZCL_XLOM__VA_EMPTY .
ENDCLASS.



CLASS ZCL_XLOM__VA_EMPTY IMPLEMENTATION.


  method GET_SINGLETON.

    IF singleton IS NOT BOUND.
      singleton = NEW ZCL_xlom__va_empty( ).
      singleton->ZIF_xlom__va~type = ZIF_xlom__va=>c_type-empty.
    ENDIF.
    result = singleton.

  endmethod.


  method ZIF_XLOM__VA~GET_VALUE.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.


  method ZIF_XLOM__VA~IS_ARRAY.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.


  method ZIF_XLOM__VA~IS_BOOLEAN.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.


  method ZIF_XLOM__VA~IS_EQUAL.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.


  method ZIF_XLOM__VA~IS_ERROR.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.


  method ZIF_XLOM__VA~IS_NUMBER.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.


  method ZIF_XLOM__VA~IS_STRING.

    RAISE EXCEPTION TYPE ZCX_xlom_todo.

  endmethod.
ENDCLASS.
