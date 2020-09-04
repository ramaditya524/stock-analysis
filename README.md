# stock-analysis

## Overview of the project

The objective of this project was to help Steve analyze large stock market datasets with relative ease by using macros. 

For this specific challenge, we are refactoring code written from earlier parts of the analysis to make it more efficient and run faster. 

## Results

The results can be divided into two specific areas: stock performance and code performance

### Stock Performance

There is a clear distinction in stock performance between the years 2017 and 2018.

![Analysis_images](Resources/2017_stock_analysis.PNG)

![Analysis_images](Resources/2018_stock_analysis.PNG)

As we can see from the above images, the year 2017 was a very good year in terms of returns on stock value. Only one stock of twelve stocks analyzed had negative returns and some stocks enjoyed returns higher than 100%.

The year 2018, however was quite the opposite. Most of the stocks performed poorly. Only two of the twelve stocks analyzed had positive returns.

### Code Performance

In terms of the code performance, the refactored code had run times that were much quicker than the original code. With increasing dataset set, this difference would keep getting more significant.

![Analysis_images](Resources/VBA_Challenge_2017_old.PNG)
![Analysis_images](Resources/VBA_Challenge_2017.PNG)

As we can see from the images, the refactored code(right image) ran much faster than the original code(left image) for stock analysis of the year 2017.

![Analysis_images](Resources/VBA_Challenge_2018_old.PNG)
![Analysis_images](Resources/VBA_Challenge_2018.PNG)

For the stock analysis of 2018, the same pattern repeats. The refactored code(right image) ran much faster than the original code(left image).

## Summary

In this section we will talk about the advantages and disadvantages of refactoring code and how these apply to in a VBA specific scenario.

### Pros and Cons of refactoring

The big pro of refactoring is that we are trying to reach a state which is the best way accomplish a task using computer code. It's highly efficient, uses the least possible memory, takes as few steps as possible and makes it easy for the user to understand.

The flipside however is that refactoring is often complex for new users and they may find it difficult to understand the code. Refactoring may achieve the most efficient code but it may not be everyone's cup of tea.

### Refactoring in the context of VBA

Refactoring, as we saw in this example, saved a lot of processing time to run the VBA code and that is a huge advantage in this context. VBA often deals with large datasets and the more efficient the code, the better. The drawback is the amount of resources, especially time, required to get it right. There is a lot of debugging needed to ensure things are running smoothly. It's a rare case where a programmer can refactor code to be the best possible in the first attempt. It takes multiple iterations, changes, debug tests and some hair pulling to get it done.
