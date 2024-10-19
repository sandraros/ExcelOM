CLASS zcx_xlom__va DEFINITION
  PUBLIC
  INHERITING FROM cx_static_check
  CREATE PUBLIC.

  PUBLIC SECTION.
    DATA result_error TYPE REF TO zcl_xlom__va_error READ-ONLY.

    METHODS constructor
      IMPORTING result_error TYPE REF TO zcl_xlom__va_error.

  PROTECTED SECTION.

  PRIVATE SECTION.
ENDCLASS.


CLASS zcx_xlom__va IMPLEMENTATION.
  METHOD constructor ##ADT_SUPPRESS_GENERATION.
    super->constructor( textid   = textid
                        previous = previous ).
    me->result_error = result_error.
  ENDMETHOD.
ENDCLASS.
