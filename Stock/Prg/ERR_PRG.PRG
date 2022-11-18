PARAMETERS mtterror, mttess, mttess1, mttprog, mttlineno

errtttt = ALIAS()
SELECT 0
USE error
APPEND BLANK
REPLACE error    WITH mtterror
REPLACE message  WITH mttess
REPLACE message1 WITH mttess1
REPLACE prog     WITH mttprog
REPLACE lineno   WITH mttlineno
REPLACE date     WITH DATE()
REPLACE time     WITH time()
SELECT error
USE
IF !EMPTY(errtttt)
   SELECT (errtttt)
ENDIF
RETURN
