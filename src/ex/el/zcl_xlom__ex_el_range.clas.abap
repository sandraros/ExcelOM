CLASS zcl_xlom__ex_el_range DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.
    INTERFACES zif_xlom__ex.
    INTERFACES zif_xlom__ex_array.

    CLASS-METHODS create
      IMPORTING address_or_name TYPE string
      RETURNING VALUE(result)   TYPE REF TO zcl_xlom__ex_el_range.

  PRIVATE SECTION.
    DATA _address_or_name TYPE string.
ENDCLASS.


CLASS zcl_xlom__ex_el_range IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_el_range( ).
    result->_address_or_name  = address_or_name.
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-range.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    TRY.
        result = zif_xlom__ex~set_result( zcl_xlom_range=>create_from_address_or_name(
                                              address     = _address_or_name
                                              relative_to = context->worksheet ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_unexpected.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    IF expression->type <> zif_xlom__ex=>c_type-range.
      RETURN.
    ENDIF.
    DATA(compare_range) = CAST zcl_xlom__ex_el_range( expression ).
    result = xsdbool( _address_or_name = compare_range->_address_or_name ).
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
