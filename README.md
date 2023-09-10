# VBA_BillConciergeProgram

This is the Code that organizes all the transactions (over the period of the entered date range) for each user id (or household), and identifies any inconsistencies in bill provider fees, and flags them.


 Flagged cases:

 i. Provider charges that appear less often than they are supposed to. (day range varies from provider to provider - code refers to a key with the specific cases - ’DaysChange’ Macro)
 
 ii. Provider charges that appear more often than they are supposed to. (day range varies from provider to provider - code refers to a key with the specific cases - ’DaysChange’ Macro)
 
 iii. Charges from a certain provider that don&#39;t appear for a while (e.g., after 2 or more months)
 
 iv. Bill charges that are out of range from the referred excel sheet with the bill key. (‘Delta’ Macro Flags this one)
 
 v. Fixed Payment charges that change from the previous payment to current payment (changes by 1 cent or more).
