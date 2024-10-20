FUNCTION Z_XLOM_UI_DATA_VIEWER.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     REFERENCE(DATA_VIEWER) TYPE REF TO  ZCL_XLOM__VW
*"----------------------------------------------------------------------
  global_data-data_viewer = data_viewer.
  CALL SCREEN 100.
ENDFUNCTION.
