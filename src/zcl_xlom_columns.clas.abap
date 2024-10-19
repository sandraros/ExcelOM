CLASS zcl_xlom_columns DEFINITION
  PUBLIC
  INHERITING FROM zcl_xlom_range FINAL
  CREATE PUBLIC
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    METHODS count REDEFINITION.

  PROTECTED SECTION.

  PRIVATE SECTION.
ENDCLASS.


CLASS zcl_xlom_columns IMPLEMENTATION.
  METHOD count.
    result = zif_xlom__va_array~column_count.
  ENDMETHOD.
ENDCLASS.
