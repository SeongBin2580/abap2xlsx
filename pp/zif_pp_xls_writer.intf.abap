INTERFACE zif_pp_xls_writer
  PUBLIC .


  METHODS write_file
    IMPORTING
      !io_excel      TYPE REF TO zcl_pp_xls
    RETURNING
      VALUE(ep_file) TYPE xstring
    RAISING
      zcx_pp_xls.
ENDINTERFACE.
