CLASS zcl_xlom__ex_el_boolean DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-DATA false TYPE REF TO zcl_xlom__ex_el_boolean.
    CLASS-DATA true  TYPE REF TO zcl_xlom__ex_el_boolean.

    CLASS-METHODS class_constructor.

    METHODS constructor.

    CLASS-METHODS create
      IMPORTING boolean_value TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_el_boolean.

  PRIVATE SECTION.
    DATA boolean_value TYPE abap_bool.

    CLASS-METHODS _create
      IMPORTING boolean_value TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_el_boolean.
ENDCLASS.


CLASS zcl_xlom__ex_el_boolean IMPLEMENTATION.
  METHOD class_constructor.
    false = create( abap_false ).
    true = create( abap_false ).
  ENDMETHOD.

  METHOD constructor.
    zif_xlom__ex~type = zif_xlom__ex=>c_type-boolean.
  ENDMETHOD.

  METHOD create.
    IF boolean_value = abap_false.
      if false is not bound.
        false = _create( abap_false ).
      endif.
      RESULT = false.
    else.
      if true is not bound.
        true = _create( abap_true ).
      endif.
      RESULT = true.
    endif.
  ENDMETHOD.

  METHOD _create.
    result = NEW zcl_xlom__ex_el_boolean( ).
    result->boolean_value = boolean_value.
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-boolean.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    result = zif_xlom__ex~set_result( zcl_xlom__va_boolean=>get( boolean_value ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    result = zcl_xlom__va_boolean=>get( boolean_value ).
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE zcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~set_result.
*    zif_xlom__ex~result_of_evaluation = value.
*    result = value.
  ENDMETHOD.
ENDCLASS.
