"! https://learn.microsoft.com/en-us/office/vba/api/excel.workbook
CLASS zcl_xlom_workbook DEFINITION
  PUBLIC
  CREATE PUBLIC
  GLOBAL FRIENDS zif_xlom__ut_all_friends.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.

    TYPES ty_name TYPE string.

    DATA application TYPE REF TO zcl_xlom_application READ-ONLY.
    "! workbook name
    DATA name        TYPE string                      READ-ONLY.
    "! workbook path
    DATA path        TYPE string                      READ-ONLY.
    DATA worksheets  TYPE REF TO zcl_xlom_worksheets  READ-ONLY.

    CLASS-METHODS create
      IMPORTING !application  TYPE REF TO zcl_xlom_application
      RETURNING VALUE(result) TYPE REF TO zcl_xlom_workbook.

    "! SaveAs (FileName, FileFormat, Password, WriteResPassword,
    "!         ReadOnlyRecommended, CreateBackup, AccessMode,
    "!         ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)
    "! https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
    "!
    "! @parameter file_name | A string that indicates the name of the file to be saved. You can include
    "!                        a full path; if you don't, Microsoft Excel saves the file in the current folder.
    METHODS save_as
      IMPORTING file_name TYPE csequence.

  PROTECTED SECTION.

  PRIVATE SECTION.
    EVENTS saved.
ENDCLASS.


CLASS zcl_xlom_workbook IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom_workbook( ).
    result->application = application.
    result->worksheets  = zcl_xlom_worksheets=>create( workbook = result ).
    result->worksheets->add( name = 'Sheet1' ).
  ENDMETHOD.

  METHOD save_as.
    " TODO: parameter FILE_NAME is never used (ABAP cleaner)

    RAISE EVENT saved.
  ENDMETHOD.
ENDCLASS.
