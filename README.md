# dss-stocks
Built a decision support system (DSS) to predict a stock's next day price based on historical data. The DSS was VBA-based, and consisted of a multilayer perceptron using Google Finance data obtained by a web query. 

There are two modules: Historical and Perceptron.

### Historical
Takes a user-inputted tickery symbol (i.e. AAPL), and queries Google Finance to retrieve historical stock data.

As seen in the image below, the user can click a button to refresh the historical data query for the NASDAQ stock symbol chosen.

![historical](https://user-images.githubusercontent.com/16723379/40272632-643bfa4c-5b7e-11e8-984e-5ef259fdb0f2.PNG)





### Perceptron 
Estimates the stock's next day closing price. Using Excel Solver, the multilayer perceptron's internal weights are optimized based on the chosen metric. The user can select to minimize the total sum of error, or to minimize the maximum error.

Again, the user can select which optimization metric to use by clicking a button.

[](url)
![estimate-next](https://user-images.githubusercontent.com/16723379/40272616-3d7f2eec-5b7e-11e8-9341-fef12fd2860b.PNG)
