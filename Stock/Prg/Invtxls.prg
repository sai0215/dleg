PARAMETERS title,ttype

totqtyin    = 0
totqtyout   = 0
totcurqty   = 0
totcostin   = 0
totcurcost  = 0
gtotqtyin   = 0
gtotqtyout  = 0
gtotcurqty  = 0
gtotcostin  = 0
gtotcurcost = 0
IF !USED('invxls')
	SELECT 0
	USE invxls
ELSE
	SELECT invxls	
ENDIF	
= checkzapfile('invxls')
APPEND BLANK
REPLACE product WITH 'STOCK INVENTORY  ('+UPPER(title)+')'
IF ttype>0
	REPLACE barcode WITH IIF(ttype=1,'In stock prod.','All products')
ENDIF	
APPEND BLANK
IF EMPTY(m.frdate)
	REPLACE product WITH SPACE(10)+chgdate(m.todate)
ELSE
	REPLACE product WITH 'FROM '+DTOC(m.frdate)+' TO '+DTOC(m.todate)
ENDIF
IF m.sea
	APPEND BLANK
	REPLACE product WITH SPACE(10)+m.famdes
ENDIF	
APPEND BLANK
APPEND BLANK
APPEND BLANK
APPEND BLANK
SELECT invtab
SET ORDER TO fmname
GO TOP
m.famser = famtab1
SELECT invxls
REPLACE product WITH ' '+famtab1.name
SELECT invtab
DO WHILE !EOF()
	SCATTER MEMVAR
	SELECT invxls
	APPEND BLANK
	REPLACE product   WITH m.name
	REPLACE barcode   WITH m.barcode
	REPLACE size      WITH fl4tab.name
	REPLACE qty_in    WITH qtyini
	REPLACE qty_out   WITH qtyexit
	REPLACE curr_qty  WITH invqty
	REPLACE cost_unit WITH costff
	REPLACE cost_in   WITH totini
	REPLACE curr_cost WITH totsel
	totqtyin   = totqtyin+qtyini
	totqtyout  = totqtyout+qtyexit
	totcurqty  = totcurqty+invqty
	totcostin  = totcostin+totini
	totcurcost = totcurcost+totsel
	SELECT invtab
	SKIP
	IF famtab1<>m.famser
		m.famser   = famtab1
		SELECT invxls
		APPEND BLANK
		REPLACE barcode   WITH 'Total:'
		REPLACE qty_in    WITH totqtyin
		REPLACE qty_out   WITH totqtyout
		REPLACE curr_qty  WITH totcurqty
		REPLACE cost_in   WITH totcostin
		REPLACE curr_cost WITH totcurcost
		APPEND BLANK
		APPEND BLANK
		REPLACE product WITH ' '+famtab1.name
		SELECT invtab
		gtotqtyin   = gtotqtyin+totqtyin
		gtotqtyout  = gtotqtyout+totqtyout
		gtotcurqty  = gtotcurqty+totcurqty
		gtotcostin  = gtotcostin+totcostin
		gtotcurcost = gtotcurcost+totcurcost
		
		totqtyin   = 0
		totqtyout  = 0
		totcurqty  = 0
		totcostin  = 0
		totcurcost = 0		
	ENDIF
ENDDO
SELECT invxls
APPEND BLANK
APPEND BLANK
REPLACE product WITH 'Total Quantity IN'
REPLACE barcode WITH SPACE(9)+STR(gtotqtyin,6)
APPEND BLANK
REPLACE product WITH 'Total Quantity OUT'
REPLACE barcode WITH SPACE(9)+STR(gtotqtyout,6)
APPEND BLANK
REPLACE product WITH 'Total Current QUANTITY'
REPLACE barcode WITH SPACE(9)+STR(gtotcurqty,6)
APPEND BLANK
APPEND BLANK
REPLACE product WITH 'Total Cost IN'
REPLACE barcode WITH SPACE(4)+STR(gtotcostin,11,2)
REPLACE size    WITH 'F.F.'
APPEND BLANK
REPLACE product WITH 'Total Current Cost'
REPLACE barcode WITH SPACE(4)+STR(gtotcurcost,11,2)
REPLACE size    WITH 'F.F.'
filename = 'iv'+PADL(MONTH(DATE()),2,'0')+STR(YEAR(DATE()),4)+'.xls'
savefile = PUTFILE('Save as',filename,'xls')
COPY TO &savefile TYPE XLS FIELDS Invxls.product,Invxls.barcode,Invxls.size,Invxls.qty_in,Invxls.qty_out,Invxls.curr_qty,Invxls.cost_unit,Invxls.cost_in,Invxls.curr_cost