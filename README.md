# VBA_BillConciergeProgram

This is the Code that organizes all the transactions (over the period of the entered date range) for each user id (or household), and identifies any inconsistencies in bill provider fees, and flags them.


 Flagged cases:

 i. Certain Provider charges that appear less often than they are supposed to. (day range varies from provider to provider - code refers to a key with the specific cases)
 
 ii. Certain Provider charges that appear more often than they are supposed to. (day range varies from provider to provider - code refers to a key with the specific cases)
 
 iii. Certain provider charges that don&#39;t appear for a while (e.g., after 2 or more months)
 
 iv. Certain Provider charge amounts that are out of range from the referred excel sheet with the bill key.
 
 v. Fixed Payment charges that change from the previous payment to current payment (changes by 1 cent or more).
