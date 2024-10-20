CLASS zcl_xlom__ex_el_string DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !text         TYPE csequence
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_el_string.

  PRIVATE SECTION.
    DATA string TYPE string.
ENDCLASS.


CLASS zcl_xlom__ex_el_string IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_el_string( ).
    result->string            = text.
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-string.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    result = zif_xlom__ex~set_result( zcl_xlom__va_string=>create( string ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    result = zcl_xlom__va_string=>get( string ).
*    RAISE EXCEPTION TYPE zcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    IF expression->type <> zif_xlom__ex=>c_type-string.
*      RETURN.
*    ENDIF.
*    DATA(string_object) = CAST zcl_xlom__ex_el_string( expression ).
*    result = xsdbool( string = string_object->string ).
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~set_result.
*    zif_xlom__ex~result_of_evaluation = value.
*    result = value.
  ENDMETHOD.
ENDCLASS.
