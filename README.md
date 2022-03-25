## spending-calc
!!!!!! NOT DONE !!!!!!!
Calculates my total income and expense and prints it into a excel sheets.
Skandiabanken provides an xlsx file with all the transactions in a choosen 
timespan and that file is what I used to model this program. 

Takes two arguments
First is a other income/expenses integer
Second is an xlsx file with all the transactions for the month.

### Example file

  A     B         C       
1
2
3
4 Date  Descript.  Amount
5       Ica        130,00
6       Coop       321,30


### Summary
It will summarise total expenses, income and the difference and then write it to 
another xlsx document at the last column of the rows.


The program will view transactions with description "Swish från" as something I 
payed for a friend and therefore subtract that amount from total expenses. 
