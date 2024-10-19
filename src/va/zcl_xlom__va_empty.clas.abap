CLASS zcl_xlom__va_empty DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__va.

    CLASS-METHODS get_singleton
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_empty.

  PRIVATE SECTION.
    CLASS-DATA singleton TYPE REF TO zcl_xlom__va_empty.
ENDCLASS.


CLASS zcl_xlom__va_empty IMPLEMENTATION.
  METHOD get_singleton.
    IF singleton IS NOT BOUND.
      singleton = NEW zcl_xlom__va_empty( ).
      singleton->zif_xlom__va~type = zif_xlom__va=>c_type-empty.
    ENDIF.
    result = singleton.
  ENDMETHOD.

  METHOD zif_xlom__va~get_value.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_array.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_boolean.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_error.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_number.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_string.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.
ENDCLASS.
