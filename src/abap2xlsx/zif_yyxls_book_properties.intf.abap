INTERFACE zif_yyxls_book_properties
  PUBLIC .

  TYPES tv_excel_appversion TYPE c LENGTH 7.

  DATA creator TYPE zyyxls_creator .
  DATA lastmodifiedby TYPE zyyxls_creator .
  DATA created TYPE timestampl .
  DATA modified TYPE timestampl .
  DATA title TYPE zyyxls_title .
  DATA subject TYPE zyyxls_subject .
  DATA description TYPE zyyxls_description .
  DATA keywords TYPE zyyxls_keywords .
  DATA category TYPE zyyxls_category .
  DATA company TYPE zyyxls_company .
  DATA application TYPE zyyxls_application .
  DATA docsecurity TYPE zyyxls_docsecurity .
  DATA scalecrop TYPE zyyxls_scalecrop .
  DATA linksuptodate TYPE flag .
  DATA shareddoc TYPE flag .
  DATA hyperlinkschanged TYPE flag .
  DATA appversion TYPE tv_excel_appversion .

  METHODS initialize .
ENDINTERFACE.
