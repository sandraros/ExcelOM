CLASS zcl_xlom__ex_fu_if DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !condition    TYPE REF TO zif_xlom__ex
                expr_if_true  TYPE REF TO zif_xlom__ex
                expr_if_false TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_if.

  PRIVATE SECTION.
    DATA condition     TYPE REF TO zif_xlom__ex.
    DATA expr_if_true  TYPE REF TO zif_xlom__ex.
    DATA expr_if_false TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_if IMPLEMENTATION.
  METHOD create.
    IF    condition     IS NOT BOUND
       OR expr_if_true  IS NOT BOUND
       OR expr_if_false IS NOT BOUND.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
    result = NEW zcl_xlom__ex_fu_if( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-if.
    result->condition         = condition.
    result->expr_if_true      = expr_if_true.
    result->expr_if_false     = expr_if_false.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(condition_evaluation) = zcl_xlom__va=>to_boolean( condition->evaluate( context ) ).
    result = zif_xlom__ex~set_result( COND #( WHEN condition_evaluation = zcl_xlom__va_boolean=>true
                                              THEN expr_if_true->evaluate( context )
                                              ELSE expr_if_false->evaluate( context ) ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    IF expression->type = zif_xlom__ex=>c_type-function-if.
      DATA(if) = CAST zcl_xlom__ex_fu_if( expression ).
      IF     condition->is_equal( if->condition )
         AND expr_if_true->is_equal( if->expr_if_true )
         AND expr_if_false->is_equal( if->expr_if_false ).
        result = abap_true.
      ELSE.
        result = abap_false.
      ENDIF.
    ELSE.
      result = abap_false.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
