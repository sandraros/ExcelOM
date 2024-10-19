"! https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets
CLASS zcl_xlom_worksheets DEFINITION
  PUBLIC
  CREATE PUBLIC
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    DATA application TYPE REF TO zcl_xlom_application READ-ONLY.
    DATA count       TYPE i                           READ-ONLY.
    DATA workbook    TYPE REF TO zcl_xlom_workbook    READ-ONLY.

    METHODS add
      IMPORTING !name         TYPE zcl_xlom_worksheet=>ty_name
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_worksheet.

    CLASS-METHODS create
      IMPORTING workbook      TYPE REF TO zcl_xlom_workbook
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_worksheets.

    "! @parameter index  | Required    Variant The name or index number of the object.
    METHODS item
      IMPORTING !index        TYPE simple
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_worksheet
      RAISING   zcx_xlom__va.

  PROTECTED SECTION.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ty_worksheet,
        name   TYPE zcl_xlom_worksheet=>ty_name,
        object TYPE REF TO zcl_xlom_worksheet,
      END OF ty_worksheet.
    TYPES ty_worksheets TYPE SORTED TABLE OF ty_worksheet WITH UNIQUE KEY name.

    DATA worksheets TYPE ty_worksheets.
ENDCLASS.


CLASS zcl_xlom_worksheets IMPLEMENTATION.
  METHOD add.
    DATA worksheet TYPE ty_worksheet.

    worksheet-name   = name.
    worksheet-object = zcl_xlom_worksheet=>create( workbook = workbook
                                                   name     = name ).
    INSERT worksheet INTO TABLE worksheets.
    count = count + 1.

    application->active_sheet = worksheet-object.

    result = worksheet-object.
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom_worksheets( ).
    result->application = workbook->application.
    result->workbook    = workbook.
  ENDMETHOD.

  METHOD item.
    TRY.
        CASE zcl_xlom_application=>type( index ).
          WHEN cl_abap_typedescr=>typekind_string
            OR cl_abap_typedescr=>typekind_char.
            result = worksheets[ name = index ]-object.
          WHEN cl_abap_typedescr=>typekind_int.
            result = worksheets[ index ]-object.
          WHEN OTHERS.
            RAISE EXCEPTION TYPE zcx_xlom_todo.
        ENDCASE.
      CATCH cx_sy_itab_line_not_found.
        RAISE EXCEPTION TYPE zcx_xlom__va
          EXPORTING result_error = zcl_xlom__va_error=>ref.
    ENDTRY.
  ENDMETHOD.
ENDCLASS.
