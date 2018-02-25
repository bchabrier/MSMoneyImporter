@echo OFF
setlocal EnableExtensions
setlocal EnableDelayedExpansion

: Transforms a transaction
: 
: Available transaction fields are stored in the 'fields' variable, e.g. field=NAME, 
: field=MEMO. Each field value is stored in variable field_<field>, e.g. field_NAME, 
: field_MEMO.
: 
: New fields can be added in 'fields', with their values defined accordingly.
: 
: Transformed or added fields must be echo'ed with their value, in order to be taken 
: into account, e.g:
: echo field_NAME=the new value of NAME
: 


