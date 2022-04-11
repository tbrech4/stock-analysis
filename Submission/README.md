# stock-analysis

## Overview of Project

The purpose of this project was to utilize VBA to automate analysis on the total annual volume and percentage return of twelve stocks across 2017 and 2018. The daily closing price & volume of each ticker were the data points that were used in this analysis, but the dataset also included other daily pricing information - including daily highs, lows, and opening price. 

## Results

### 2017 Results

![2017 returns] (https://raw.githubusercontent.com/tbrech4/stock-analysis/main/Resources/2017_Returns.png)

In 2017, all but one of the selected stocks - $TERP - had a positive return. The best performing stock was $DQ, returning almost 200% in 2017 - this is a very impressive single year gain. $DQ also had the lowest total yearly volume of all 12 stocks, however there isn't evidence of a relationship between volume and returns as the other stocks that returned over 100% were also traded very frequently.

If you were to invest $10,000 in each of the twelve stocks at the beginning of 2017, you would have had a 67% gain. 

### 2018 Results

2018 was a rough year for the selected stocks. Only 17% of the selected stocks had a positive return. The biggest losers were $DQ and $JKS, both losing over 50% of their value in 2017. DQ proved to be a very volatile stock between 2017 and 2018 as it was the best performing stock in 2017, and one of the worst performing stocks in 2018.

![2018 returns] (https://raw.githubusercontent.com/tbrech4/stock-analysis/main/Resources/2018_Returns.png)

If you would have kept the original 10,000 you invested in each of the twelve stocks at the beginning of 2017, you would now have a 48% return over the two-year period. However, your portfolio would have dropped 11% in 2018.

### Execution times of refactored code

Refactoring the code led to a dramatic speed increase. For example, the 2017 code ran in .0078 seconds after refactoring the code. This is 620% faster than the original code took to ran (about half a second).

![2017 code time] (https://github.com/tbrech4/stock-analysis/blob/main/Resources/2017_Runtime.png)

The speed increase when analyzing 2018 data was very similar. This makes sense since both years had the same amount of data (3013 lines of data).

## Summary

### Advantages / Disadvantages of refactoring code

One of the biggest advantages of refactoring code is potential to make code optimized and run faster. Another advantage is that your code could possibly be cleaner and better structured after refactoring.

The disadvantage is that you could potentially create new bugs, or even break the code. "If it ain't broke, don't fix it" is a saying that comes to my mind here - trying to fix code to make it run better could lead to the code not working at all. Luckily with proper version control, you should be able to always go back to versions of the code where everything ran correctly.

### Pros / Cons of refactoring the original VBA script

The main advantage of refactoring the VBA script was the huge speed increase that was observed. The original code only took about half a second to run, but the new code managed to improve speeds by 600%. 

If there were 100 stocks and 200,000 lines of data, then the difference in timing would be much more tangible to the user.

I don't think there were really any cons to refactoring the code in this specific example since I was able to get the refactored code running correctly. 

