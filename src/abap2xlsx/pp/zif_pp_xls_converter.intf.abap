INTERFACE zif_pp_xls_converter
  PUBLIC .


  METHODS can_convert_object
    IMPORTING
      !io_object TYPE REF TO object
    RAISING
      zcx_pp_xls .
  METHODS create_fieldcatalog
    IMPORTING
      !is_option       TYPE zpps_xls_converter_option
      !io_object       TYPE REF TO object
      !it_table        TYPE STANDARD TABLE
    EXPORTING
      !es_layout       TYPE zpps_xls_converter_layo
      !et_fieldcatalog TYPE zppy_xls_converter_fcat
      !eo_table        TYPE REF TO data
      !et_colors       TYPE zppy_xls_converter_col
      !et_filter       TYPE zppy_xls_converter_fil
    RAISING
      zcx_pp_xls .

   METHODS get_supported_class
     RETURNING VALUE(rv_supported_class) TYPE seoclsname.
ENDINTERFACE.
