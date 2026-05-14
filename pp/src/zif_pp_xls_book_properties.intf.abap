INTERFACE zif_pp_xls_book_properties
  PUBLIC .

  TYPES tv_excel_appversion TYPE c LENGTH 7.

  DATA creator TYPE zppe_xls_creator .
  DATA lastmodifiedby TYPE zppe_xls_creator .
  DATA created TYPE timestampl .
  DATA modified TYPE timestampl .
  DATA title TYPE zppe_xls_title .
  DATA subject TYPE zppe_xls_subject .
  DATA description TYPE zppe_xls_description .
  DATA keywords TYPE zppe_xls_keywords .
  DATA category TYPE zppe_xls_category .
  DATA company TYPE zppe_xls_company .
  DATA application TYPE zppe_xls_application .
  DATA docsecurity TYPE zppe_xls_docsecurity .
  DATA scalecrop TYPE zppe_xls_scalecrop .
  DATA linksuptodate TYPE flag .
  DATA shareddoc TYPE flag .
  DATA hyperlinkschanged TYPE flag .
  DATA appversion TYPE tv_excel_appversion .

  METHODS initialize .
ENDINTERFACE.
