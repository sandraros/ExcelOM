CLASS zcl_xlom__ex_el_number DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !number       TYPE f
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_el_number.

  PRIVATE SECTION.
    DATA number TYPE f.
ENDCLASS.


CLASS zcl_xlom__ex_el_number IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_el_number( ).
    result->number            = number.
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-number.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    result = zif_xlom__ex~set_result( zcl_xlom__va_number=>create( number ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    IF     expression->type = zif_xlom__ex=>c_type-number
       AND number           = CAST zcl_xlom__ex_el_number( expression )->number.
      result = abap_true.
    ELSE.
      result = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
