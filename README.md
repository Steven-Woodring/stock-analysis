# stock-analysis
#### Analyzing green energy stocks to make an educated financial investment

## Overview of Project
#### This analysis of stock data utilizes Excel VBA to aid in the construction of a successful, diversified investment portfolio. Our fictional client, Steve, is attempting to advise his parents as to which green energy stocks are reasonable investments. From data that gives the activity of these green stocks over two years (2017 and 2018), the output of the final automated analysis is the total daily volume and percent return from the chosen year. With these values automatically calculated in less than a second, it is easy to make an informed decision on whether to invest in each stock.

## Results
#### In general, 2017 was a great year for our selected green energy stocks with TERP being the lone stock out of the twelve that had a negative return. Additionally, this company only experienced a 7.2% drop in price over the whole year, a value that was dwarfed by some of the positive returns from the other energy stocks. One third of the stocks increased over 100% in value. DQ, while having the lowest total daily volume, was the best performer in 2017 with nearly a 200% price jump. In terms of total volume, most of the stocks saw somewhere around 100,000,000 – 300,000,000 trades throughout the year. The three stocks that greatly differed from this trend were DQ (35,796,200), FSLR (684,181,400), and SPWR (782,187,000).

<img width="225" alt="Screen Shot 2022-01-01 at 9 14 45 PM" src="https://user-images.githubusercontent.com/95303422/147895168-7d1e17aa-bd89-477f-9355-4dbcfb75e14d.png">

#### On the contrary, the same selection of green energy stocks performed poorly for the most part in 2018 with only two companies winding up in the green. Half of the stocks with negative returns lost more than 20% of their value from the start of the year, and the average return of the entire selection was -8.5%. The silver lining of 2018 was that the stocks that did increase in price, ENPH and RUN, gained 81.9% and 84.0%, respectively. In terms of total volume, it was overall slightly higher than the previous year. The lowest volume was AY with 83,079,900, and four companies (FSLR, RUN, SPWR, and ENPH) saw over 475 million trades.

<img width="225" alt="Screen Shot 2022-01-01 at 9 15 41 PM" src="https://user-images.githubusercontent.com/95303422/147895198-400c62d7-d5b3-4c5e-9e8f-e734425dd470.png">

#### The original script took over a minute to run, while the refactored script outputs the values almost instantly, clocking in consistently around 0.15 seconds. The refactored script is about 600 times faster than the original.

<img width="249" alt="Screen Shot 2022-01-01 at 10 37 27 PM" src="https://user-images.githubusercontent.com/95303422/147895217-19818120-dc67-48ee-bdf0-6eea4a8de65e.png">

## Summary
#### The first advantage of refactoring code is that it helps to find bugs. By making the code more concise, it is easier to follow the code logically and identify issues. The obvious advantage, of course, is that it speeds up the execution of the program. By altering sections of the code such as variables, loops, and if-then statements, the program requires less computer processing power and executes faster. An additional benefit is that it hopefully makes the code clean and organized by eliminating unnecessary variables and other elements.

#### The first disadvantage of refactoring code is that it takes time. If the original code is already substantially efficient, it might not end up being worth it to take the time to refactor the code. More importantly, refactoring can be risky. Though the product of a refactor may make it easier to identify bugs, it is possible to create bugs during the process, especially when working with an extensive program.

#### When looking at this VBA code specifically, it is apparent that the advantages of refactoring far outweighed the disadvantages. The process only took a moderate amount of time, and considerably reduced the amount of time the program takes to execute. On top of the faster speed, the refactored code is also organized and concise, and the process did not create any new bugs. This will help to identify any bugs that may emerge in the future if the code were to be expanded.
