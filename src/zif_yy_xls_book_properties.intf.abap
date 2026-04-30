INTERFACE zif_yy_xls_book_properties
  PUBLIC .

  TYPES tv_excel_appversion TYPE c LENGTH 7.

  DATA creator TYPE zyye_xls_creator .
  DATA lastmodifiedby TYPE zyye_xls_creator .
  DATA created TYPE timestampl .
  DATA modified TYPE timestampl .
  DATA title TYPE zyye_xls_title .
  DATA subject TYPE zyye_xls_subject .
  DATA description TYPE zyye_xls_description .
  DATA keywords TYPE zyye_xls_keywords .
  DATA category TYPE zyye_xls_category .
  DATA company TYPE zyye_xls_company .
  DATA application TYPE zyye_xls_application .
  DATA docsecurity TYPE zyye_xls_docsecurity .
  DATA scalecrop TYPE zyye_xls_scalecrop .
  DATA linksuptodate TYPE flag .
  DATA shareddoc TYPE flag .
  DATA hyperlinkschanged TYPE flag .
  DATA appversion TYPE tv_excel_appversion .

  METHODS initialize .
ENDINTERFACE.
