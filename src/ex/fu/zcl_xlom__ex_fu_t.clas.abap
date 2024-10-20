"! T(value)
"! If value is or refers to text, T returns value. If value does not refer to text, T returns "" (empty text).
"! Examples: T("text") = "text", T(1) = "", T({1}) = "", T(FALSE) = "", T(XFD1024000) = "" (empty). But T(#N/A) = #N/A.
"! https://support.microsoft.com/en-us/office/t-function-fb83aeec-45e7-4924-af95-53e073541228
CLASS zcl_xlom__ex_fu_t DEFINITION
  PUBLIC FINAL
  INHERITING FROM zcl_xlom__ex_fu
*  CREATE PRIVATE
  GLOBAL FRIENDS zcl_xlom__ex_fu.

  PUBLIC SECTION.
    CLASS-METHODS create
      IMPORTING !value        TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_t.

    METHODs zif_xlom__ex~evaluate_single REDEFINITION.

  PROTECTED SECTION.
    METHODS constructor.

  PRIVATE SECTION.
    CONSTANTS:
      BEGIN OF c_arg,
        value TYPE i VALUE 1,
      END OF c_arg.

*    DATA value TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_t IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    zif_xlom__ex~type = zif_xlom__ex=>c_type-function-t.
    zif_xlom__ex~parameters = VALUE #( ( name = 'VALUE' ) ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_fu_t( ).
    result->zif_xlom__ex~arguments_or_operands = VALUE #( ( VALUE ) ).
    zcl_xlom__ex_ut=>check_arguments_or_operands(
      EXPORTING expression            = result
      CHANGING  arguments_or_operands = result->zif_xlom__ex~arguments_or_operands ).
*    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-t.
*    result->zif_xlom__ex~set_arguments_or_operands( VALUE #( ( value ) ) ).
*    result->value             = value.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~evaluate.
*    result = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
*                                 expression = me
*                                 context    = context
*                                 operands   = zif_xlom__ex~arguments_or_operands )."VALUE #( ( name = 'VALUE' object = value ) ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(value_result) = arguments[ c_arg-value ].
        result = zcl_xlom__va_string=>create( COND #( WHEN value_result->type = value_result->c_type-string
                                                 THEN zcl_xlom__va=>to_string( value_result )->get_string( )
                                                 ELSE '' ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
    zif_xlom__ex~result_of_evaluation = result.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~is_equal.
*    RAISE EXCEPTION TYPE zcx_xlom_todo.
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~set_arguments_or_operands.
*    IF lines( arguments_or_operands ) <> 1
*        OR arguments_or_operands[ c_arg-value ] IS NOT BOUND.
*      RAISE EXCEPTION TYPE zcx_xlom_todo.
*    ENDIF.
*    zif_xlom__ex~arguments_or_operands = arguments_or_operands.
**    value = arguments_or_operands[ c_arg-value ].
*  ENDMETHOD.
*
*  METHOD zif_xlom__ex~set_result.
*    zif_xlom__ex~result_of_evaluation = value.
*    result = value.
  ENDMETHOD.
ENDCLASS.
