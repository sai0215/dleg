clea
clea all
new1=0
sele 1
USE c:\stock\dbf\newarr.dbf EXCLUSIVE
*
SELECT 2
USE c:\stock\dbf\protab.dbf EXCLUSIVE
*
SELECT 1
SET ORDER TO TAG Barcode OF c:\stock\dbf\protab.cdx IN Protab
SET RELATION TO Newarr.barcode INTO Protab ADDITIVE
SET ORDER TO TAG Barcode OF c:\stock\dbf\newarr.cdx
*
sele 1
go top
new1=newarr.qty
*
do while !eof()
  sele 2
  replace qte2 with new1
  sele 1
  skip 1
  wait wind "please wait" nowait
  new1=newarr.qty
enddo
clea
clea all
return