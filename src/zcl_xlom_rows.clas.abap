CLASS zcl_xlom_rows DEFINITION
  PUBLIC
  INHERITING FROM zcl_xlom_range FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    METHODS count REDEFINITION.

  PROTECTED SECTION.

  PRIVATE SECTION.
ENDCLASS.


CLASS zcl_xlom_rows IMPLEMENTATION.
  METHOD count.
    result = zif_xlom__va_array~row_count.
  ENDMETHOD.
ENDCLASS.
