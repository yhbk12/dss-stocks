# dss-stocks
Built a decision support system (DSS) to predict a stock's next day price based on historical data. The DSS was VBA-based, and consisted of a multilayer perceptron using Google Finance data obtained by a web query. 

## There are two modules: Historical and Perceptron.

### Historical
Takes a user-inputted tickery symbol (i.e. AAPL), and queries Google Finance to retrieve historical stock data.

### Perceptron 
Estimates the stock's next day closing price. Using Excel Solver, the multilayer perceptron's internal weights are optimized based on the chosen metric. The user can select to minimize the total sum of error, or to minimize the maximum error.
