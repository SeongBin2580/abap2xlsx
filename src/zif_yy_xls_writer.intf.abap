INTERFACE zif_yy_xls_writer
  PUBLIC .


  METHODS write_file
    IMPORTING
      !io_excel      TYPE REF TO zcl_yy_xls
    RETURNING
      VALUE(ep_file) TYPE xstring
    RAISING
      zcx_yy_xls.
ENDINTERFACE.
