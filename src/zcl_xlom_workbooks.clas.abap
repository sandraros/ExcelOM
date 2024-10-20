"! https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks
CLASS zcl_xlom_workbooks DEFINITION
  PUBLIC
  CREATE PUBLIC
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    DATA application TYPE REF TO zcl_xlom_application READ-ONLY.
    DATA count       TYPE i                           READ-ONLY.

    CLASS-METHODS create
      IMPORTING !application  TYPE REF TO zcl_xlom_application
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_workbooks.

    "! Add (Template)
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.add
    METHODS add
      IMPORTING template      TYPE any OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_workbook.

    "! @parameter index  | Required    Variant The name or index number of the object.
    METHODS item
      IMPORTING !index        TYPE simple
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_workbook.

  PROTECTED SECTION.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ty_workbook,
        name   TYPE zcl_xlom_workbook=>ty_name,
        object TYPE REF TO zcl_xlom_workbook,
      END OF ty_workbook.
    TYPES ty_workbooks TYPE SORTED TABLE OF ty_workbook WITH NON-UNIQUE KEY name
                              WITH UNIQUE SORTED KEY by_object COMPONENTS object.

    DATA workbooks TYPE ty_workbooks.

    METHODS on_saved
      FOR EVENT saved OF zcl_xlom_workbook
      IMPORTING sender.
ENDCLASS.


CLASS zcl_xlom_workbooks IMPLEMENTATION.
  METHOD add.
    " TODO: parameter TEMPLATE is never used (ABAP cleaner)

    DATA workbook TYPE ty_workbook.

    workbook-object = zcl_xlom_workbook=>create( application ).
    INSERT workbook INTO TABLE workbooks.
    count = count + 1.

    SET HANDLER on_saved FOR workbook-object.
    result = workbook-object.
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom_workbooks( ).
    result->application = application.
  ENDMETHOD.

  METHOD item.
    CASE zcl_xlom_application=>type( index ).
      WHEN cl_abap_typedescr=>typekind_string.
        result = workbooks[ name = index ]-object.
      WHEN cl_abap_typedescr=>typekind_int.
        result = workbooks[ index ]-object.
      WHEN OTHERS.
        " TODO
    ENDCASE.
  ENDMETHOD.

  METHOD on_saved.
    workbooks[ KEY by_object COMPONENTS object = sender ]-name = sender->name.
  ENDMETHOD.
ENDCLASS.
