# VBA_BillConciergeProgram
# This is the Code that organizes all the transactions (over the period of the entered date range) for each user id (or household), and identifies any inconsistencies in bill provider fees, and flags them.


# Flagged cases:
# i. Provider charges that appear less often than they are supposed to. (day range varies from provider to provider - code refers to a key with the specific cases - ’DaysChange’ Macro)
# ii. Provider charges that appear more often than they are supposed to. (day range varies from provider to provider - code refers to a key with the specific cases - ’DaysChange’ Macro)
# iii. Charges from a certain provider that don&#39;t appear for a while (e.g., after 2 or more months)
# iv. Bill charges that are out of range from the referred excel sheet with the bill key. (‘Delta’ Macro Flags this one)
# v. Fixed Payment charges that change from the previous payment to current payment (changes by 1 cent or more).

# Provider Descriptions that Will never get highlighted (due to exception conditions):
# - "CAPITAL ONE AUTOPAY PYMT" & "ACH SETTLEMENT OFFSET"

# Red Highlights Occur When:
# - If a Transaction Occurs for the first time.
# - If a prior transaction occurred two months and zero days prior to the current transaction.
# - If a prior transaction occurred three or more months prior to the current transaction.
# - NOTE: For providers that are Quarterly, Semi-Annually, and Yearly, it would be 2+ months from expected time of appearance (e.g., Quarterly is expected to appear every 4 months, so it would be red if it appeared instead after exactly 6 or more months later)
# KEY: stday(icount) refers to expected days between a payment for a certain provider [stday is a collection that stores the expected days]. {stday(icount) - 30} to adjust for different days&#39; differences. This is ONLY for providers of Monthly frequency

# Blue Highlights Occur When:
# - Distance between prior and current transaction is over one month [distance between days is greater than 33 days]
# - For one month ex:
# - If the distance between the days of the payments are greater than 3, then it would highlight blue. (Month # not taken into account - already in if-condition)
# - For two month ex:
# - If the distance between the days (current - previous) is &lt;= -27, then it would not highlight blue (Month # not taken into account - already in if-condition)
# - Ex: If the date of the last Spectrum payment is May 28th, and the end date is July 1st…
# - Result: Since (1 - 28) = -27, then it would not be highlighted blue
# - If the distance between the days (current - previous) is &gt; -27, then it would highlight blue (Month # not taken into account - already in if-condition)
# - Ex: If the date of the last Spectrum payment is May 28th, and the end date is July 5th…
# - Result: Since (5 - 28) = -23, then it would be highlighted blue
# - NOTE: For providers that are Quarterly, Semi-Annually, and Yearly, separate conditions apply for each - but the logic is roughly the same - just adjusted for # of months between

# Green Highlights Occur When:
# - Distance between prior and current transaction is less than 28 days
# - Can be within the same month.
# - If the Current day - Previous day of a provider charge is &lt; 28, then it would highlight green
# - Ex: If the date of the last Spectrum payment is May 1st, and the current spectrum payment is May 27th
# - Result: Since (27 - 1) = 26, it would be highlighted green
# - For one month apart
# - If the distance in days — if the current day - the previous day of a provider charge is &lt; -3 then it would highlight green
# - Ex: If the date of the last Spectrum payment is May 28th, and the current spectrum payment is June 1st
# - Result: Since (1 - 28) = -27, it would be highlighted green
# - NOTE: For providers that are Quarterly, Semi-Annually, and Yearly, separate conditions apply for each - but the logic is roughly the same - just adjusted for # of months between

# Orange Highlights Occur When: (Logic is subjected to change)
# - If a cost is fixed, and if the amount of a certain provider charge changes by 1% or more between the current and previous payment
# - Ex: on May 5th, 2023, you paid $100 for Spectrum, and on June 5th, 2023, you paid $101 for Spectrum:
# - Result: It would highlight orange
# - NOTE: Same logic no matter which frequency.

# Pink Highlights Occur When:
# - A provider payment charge is out of range from the key — in the ‘Providers’ tab.
