# Excel-library
Collection of useful Excel udfs and macros

User defined functions that are included:

* __REVERSE__ - Function for reversing the content of a cell based on text and delimiter
`=REVERSE(text, delimiter)`

* __EXPLODE__ - Exploding function for excel. Returns the specified item of the exploded string. 0 based count
`=EXPLODE(text, delimiter, itemNumber)`

* __ISFONTCOLOR__ - Function for checking the font color of the cell
`=ISFONTCOLOR(red,green,blue)`

* __ASDISPLAYED__ - Useful helper function to show contents of the cell as is. Can be used for example for concatenating dates
`=ASDISPLAYED(targetCell)`

* __LISTLENGTH__ - Returns the number of items in the list based on delimiter
`=LISTLENGTH(text, delimiter)`

* __ISVALIDEMAIL__ - Helper function for checking validity of email addresses
`=ISVALIDEMAIL(email)`

* __RGBTOHEX__ - Converts given rgb value to HEX color
`=RGBTOHEX(red, green, blue)`

* __HEXTORGB__ - Converts HEX color to RGB also accepts html colors with #
`=HEXTORGB(hex)` 
