# data-hw2-VBA-Scripting

- Run the Micro: run_ticker_analyzer.
- Sub: run_ticker_analyser
  - run the sub "ticker_analyzer for each sheets
- Sub: ticker_analyzer

  - run for the selected sheet
  - go through ticker column to look for ticker name change.
  - once find the ticker change, do the follwing things
    - calculating yearly change, yearly change percentage, calculate the total volume
    - check if the current value is greater than the value store in greatest- variables. Assign the value to greatest if yes

- Time complexity:

  - Assumption:
    - 705714 records
    - 2835 stockers in total
    - 250 recorder per ticker
    - 3 years
  - linear search
    - the times of iteration
      - 705714 times per year
      - 2,100,000 times for 3 year
  - binary search
    - iteration times for a ticker in worst case:
      - LOG(705714) = 20
    - iteration times for a ticker in best case:
      - Log(250) = 8
    - the total iteration times
      - (20 + 8 ) \* 2835 / 2 = 39690
      - so 39690 times per year
      - 119070 times for 3 years
    - in summary
      - Binary search is 17 times faster than linear search

- Flow Chart
<<<<<<< HEAD
=======


>>>>>>> 027fbf2f5991d422d6318b79734cb691bd43d8a6
  ![Flow chart](https://github.com/Simon-Xu-Lan/data-hw2-VBA-Scripting/blob/master/data-hw2-VBA1.png)
