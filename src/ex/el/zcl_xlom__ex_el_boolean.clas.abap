CLASS zcl_xlom__ex_el_boolean DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING boolean_value TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_el_boolean.

  PRIVATE SECTION.
    DATA boolean_value TYPE abap_bool.
ENDCLASS.


CLASS zcl_xlom__ex_el_boolean IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_el_boolean( ).
    result->boolean_value     = boolean_value.
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-boolean.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    result = zif_xlom__ex~set_result( zcl_xlom__va_boolean=>get( boolean_value ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
