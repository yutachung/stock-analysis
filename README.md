# stock-analysis
## Overview of Project
	The purpose of the stock analysis was to extract information for the client with the use of creating macros
through VBA. The purpose of the macros is to create a simple way for a user to be able to click one button and be
able to find the information that would be meaninful to them. Another purpose was for the macros to be versatile
across various worksheets that would be formatted in a similar matter such that the macro would not only be
applicable to only one worksheet in excel.
## Results
	When taking a quick glance, it can be seen that the refactored script was able to be
executed faster. When the script is observed though, there were only minor changes that have been made that have
allowed the macro to perform faster. The very first one to notice would be the tickerVolumes, tickerStartingPrices,
and tickerEndingPrices have been declared, and secondly, the arrays have been used to pull the values when formatting
the values on our stock analysis worksheet. We find that the original script takes around 0.84 second while the
refactored version of the script takes somewhere around 0.14 second.
## Summary
	When looking strictly at results, we will find that refactoring code should allow the script to be more
effcient and would allow people who could or will refer to your code in the future easier to understand. It can
be found that even though the code functions, it may not be the best way to solve the problem. On the contrary it
can be precisely for the same reason that refactoring code may be disadvantageous. Spending resources on something
that carries out its specific job may not be something a company or a person is willing to spend their time on.
	These advantages and disadvantages can be tied to the original and refactored script of the stock analysis.
The code was refactored in order to complete the task faster by 0.70 seconds. 0.70 seconds is not a significant
amount of difference in performance time. However, the context in how this code would be relevant. A couple that could
be considered would be how often is the code running, and how much data does the code need to sift through? In the
extreme case that a code needs to be run on repeat, the difference in execution speed could actually become a
significant value. The difference in execution speed could also increase at a different scale depending on the code
screens through the data. On the other hand, if one were to look to refactor a code that is not necessary that often
does not need to be refactored for the sake of decreasing the execution speed by 0.70 seconds. 