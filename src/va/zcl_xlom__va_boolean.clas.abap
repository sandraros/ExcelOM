CLASS zcl_xlom__va_boolean DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__va.

    CLASS-DATA false TYPE REF TO zcl_xlom__va_boolean READ-ONLY.
    CLASS-DATA true  TYPE REF TO zcl_xlom__va_boolean READ-ONLY.

    DATA boolean_value TYPE abap_bool READ-ONLY.

    CLASS-METHODS class_constructor.

    CLASS-METHODS get
      IMPORTING boolean_value TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_boolean.

  PRIVATE SECTION.
    DATA number TYPE f.

    CLASS-METHODS create
      IMPORTING boolean_value TYPE abap_bool
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_boolean.
ENDCLASS.


CLASS zcl_xlom__va_boolean IMPLEMENTATION.
  METHOD class_constructor.
    false = create( abap_false ).
    true  = create( abap_true ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__va_boolean( ).
    result->zif_xlom__va~type = zif_xlom__va=>c_type-boolean.
    result->boolean_value     = boolean_value.
    result->number            = COND #( WHEN boolean_value = abap_true THEN -1 ).
  ENDMETHOD.

  METHOD get.
    result = SWITCH #( boolean_value
                       WHEN abap_true
                       THEN true
                       ELSE false ).
  ENDMETHOD.

  METHOD zif_xlom__va~get_value.
    result = REF #( boolean_value ).
  ENDMETHOD.

  METHOD zif_xlom__va~is_array.
    result = abap_true.
  ENDMETHOD.

  METHOD zif_xlom__va~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_error.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_string.
    result = abap_false.
  ENDMETHOD.
ENDCLASS.
