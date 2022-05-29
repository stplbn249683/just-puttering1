This repository contains an example of a stock charting program written in
Visual Basic.Net that uses historical end-of-day stock data from the 
SQL Server database table described in the repository "just-puttering".

Sometimes I just want to display some stock charts without logging in,
worrying about getting timed out, or getting distracted by continuous
ads. So I wrote this Windows desktop application.  I used the free
version of Visual Studio 2022.

This program makes use of the Skender.Stock.Indicators NuGet package to
calculate the indicators.  I had written some routines of my own but
decided that it would be easier to add additional indicators by using
the Skender package.

Disclaimer: This example includes code that I developed for my own personal
use and I have included it as a source of information.  I am not recommending
that anyone else use the code as written and I am not responsible for any
consequences that result from doing so. This post reflects my current 
understanding and may contain errors.  Since the code was developed for my
own use, I did not attempt to make it elegant or efficient.

There is of course no limit to the features that could be added to a program
like this.  The user enters the ticker symbol (which must have data in the 
database table) and the number of days to chart (ending at the most recent
date). I have used a list of check boxes to select what charts to plot.
Since the panel containing the charts scrolls, I normally just check
all of the check boxes. However, there is an error if the Chaikin money flow
chart is selected when all of the volumes are zero and some other charts are 
meaningless if all of the volumes are zero.

The data souce is read in from the text file named "DataSource.ini" located
in the application directory.  The data source could also be specified
directly in the InitializeDefaults subroutine (see the readme.txt file in
the just-puttering repository for additional comments about the data source).
The options that are selected by the user are stored in the file
"StockChart.ini" in the application directory and are used as the starting
values the next time that the program is run.

The Skender packages returns different numbers of valid points for different
indicators.  I am using "Series.IsXValueIndexed = True" so that the charts
do not contain gaps for days when the market is not open.  However, this 
requires that all of the series on the chart have the same date range. So
I have resized the lists to shorten the date range to the minimum date range
of all the indicators on the same chart.  The exception is the series for 
SMA(50), SMA(100) and SMA(200) which I have simply dropped if the ticker 
symbol does not have enough data points so that the date range matches the 
date ranges of the other indicators on the chart (this can happen for stocks
that were added to the exchange very recently).

I have included a few PNG files in the repository showing screen shots of the
output.  If I find errors in the program or add additional charts then I
will update the files and change the "Last modified" date at the top of 
the file Module1.vb.  I have not included a project file; just the files
for the input form and the code.

I would point out something that I ran into myself when using this program.
If a stock split has occurred in the stock since the historical stock quotes
were first added to the database then all of the historicl stock quotes
for that ticker symbol probably need to be deleted from the database so
that all of the historical stock quotes for that ticker symbol will be
downloaded and added back to the database again.

This can be done very carefully (using ticker symbol NVDA as an example) using
a SQL query like the following:

DECLARE @ticker varchar(10) = 'NVDA'
DELETE FROM market_price
WHERE        (Ticker = @ticker)