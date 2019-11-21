[get the cloest value from a range](https://www.mrexcel.com/board/threads/how-to-find-the-closest-value-in-a-range-to-a-given-lookup-value.629441/)

Apr 18, 2012
#3

Single cell solution, although it is a little long.
With the Range being search in A1:A13
With the Value being looked up in B1


If you have Excel 2003 or earlier you will need to use
```vb
=INDEX(A1:A13,IF(ISERROR(MATCH(B1-MIN(ABS(B1-A1:A13)),A1:A13,0)),MATCH(B1+MIN(ABS(B1-A1:A13)),A1:A13,0),MATCH(B1-MIN(ABS(B1-A1:A13)),A1:A13,0)))
```
Confirm with CTRL+SHIFT+ENTER

If you have Excel 2007 or later you can use this slightly shorter version
```vb
=INDEX(A1:A13,IFERROR(MATCH(B1-MIN(ABS(B1-A1:A13)),A1:A13,0),MATCH(B1+MIN(ABS(B1-A1:A13)),A1:A13,0)))
```
