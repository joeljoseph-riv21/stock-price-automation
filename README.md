# stock-price-automation
An RPA automation built with UiPath Studio 2026 extracting data from a web application every 30 minutes during trading hours which automatically logs the data into a structured Excel file,eliminating manual data collection with zero human intervention.

![UiPath](https://img.shields.io/badge/UiPath-Studio%202026-FF6200?style=flat&logo=uipath&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=flat)


## What This Project Does

- Opens the RPA Stock Market page on rpachallenge.com using Microsoft Edge
- Extracts the share price of **Exxon RPA Corporation** and **WEX Academy Inc** every 30 minutes
- Stores the date, time and both prices into a structured Excel file each cycle
- A double line chart inside the Excel file auto-updates as data is written
- Runs automatically from 9:00 AM to 4:00 PM and stops on its own


## Activities Used

| Activity | Purpose |
|----------|---------|
| `Assign` | Capture system time, calculate loop cycles, format date and time |
| `If` | Check if current time is within valid trading hours |
| `Do While` | Repeat data collection every 30 minutes until end of trading day |
| `Delay` | Wait 30 minutes between each collection cycle |
| `Use Application/Browser` | Open Microsoft Edge and navigate to the stock market webpage |
| `Click` | Interact with the company dropdown and search button on the webpage |
| `Get Text` | Extract the displayed share price from the webpage |
| `Build Data Table` | Create an in-memory table with 4 columns to stage collected data |
| `Add Data Row` | Add one row of data per cycle into the in-memory DataTable |
| `Use Excel File` | Open and manage the Excel workbook for writing |
| `Append Range Workbook` | Write the DataTable row into the next empty row in Excel |
| `Clear Data Table` | Reset the DataTable after each write so it stays clean for the next cycle |
| `Log Message` | Print cycle status and completion info to the Output panel |


## Variables

| Variable | Type | Purpose |
|----------|------|---------|
| `currentTime` | DateTime | Stores system clock at robot start |
| `minutesRemaining` | Int32 | Minutes left between start time and 4PM |
| `Counter` | Int32 | Total number of 30-min cycles to run |
| `Count` | Int32 | How many cycles have completed so far |
| `extractedDate` | String | Today's date captured each cycle |
| `extractedTime` | String | Timestamp captured each cycle |
| `exxonPrice` | String | Exxon RPA Corp price from webpage |
| `wexPrice` | String | WEX Academy Inc price from webpage |
| `stockDataTable` | DataTable | In-memory table holding one row per cycle |


## Prerequisites

- UiPath Studio 2026 with Excel, UIAutomation, System and Mail packages installed
- Microsoft Excel with `Data.xlsx` pre-created — headers in row 1, double line chart linked to columns C and D, AutoSave off
- Microsoft Edge browser
- Active internet connection to reach rpachallenge.com
- System must stay on from 9AM to 4PM while the robot runs


## How to Run

1. Open the project in UiPath Studio 2026
2. Make sure `Data.xlsx` exists at your configured path with headers and empty chart
3. Click **Run** between 9:00 AM and 4:00 PM
4. The robot validates the time, calculates how many cycles fit before 4PM, then collects prices every 30 minutes automatically
5. Open `Data.xlsx` after the run to see the filled data and populated chart

> For testing, change the Delay from `00:30:00` to `00:01:00` temporarily


## Project Structure
```
StockPriceAutomation/
├── Main.xaml
├── project.json
├── Data/
│   ├── Data.xlsx
│   └── Config.xlsx
└── README.md
```


## Real-Time Use Cases

- **Finance** — Banks and trading teams use similar bots to log live market prices into datasets without manual data entry
- **Retail** — E-commerce teams automate competitor price tracking by scraping product prices at regular intervals
- **Supply Chain** — Procurement teams collect commodity prices from supplier portals on a schedule for cost forecasting
- **HR & Payroll** — Finance teams automate daily currency exchange rate collection for multinational payroll processing
- **Healthcare** — Hospital teams extract operational metrics from web dashboards and log them for daily review


## Future Enhancements

- **Email reporting** — Send `Data.xlsx` as an email attachment at the end of the trading day using SMTP Mail activity *(immediate next step)*
- **Orchestrator scheduling** — Deploy to UiPath Orchestrator and use Time Triggers instead of internal Delay, making the bot more reliable and remotely manageable
- **API-based price fetching** — Replace browser scraping with direct calls to a financial market API for faster and more stable data collection
- **Database storage** — Write collected data directly into SQL Server or SQLite instead of Excel for better querying and multi-user access
- **Power BI integration** — Stream data into a live Power BI dashboard so the trend chart is accessible on any device in real time
- **Multi-company support** — Read the company list dynamically from a config file instead of hardcoding two companies, making the bot scalable to any number of stocks
- **REFramework migration** — Refactor using UiPath's Robotic Enterprise Framework for proper exception handling, retry logic and production-grade reliability
- **AI anomaly detection** — Use UiPath Autopilot to flag unusual price movements and include a natural language summary in the daily report


## Acknowledgements

- Coursera & UiPath — RPA Specialization curriculum and project specification
- RPAChallenge.com — for the RPA Stock Market practice environment

## Author
Joel Joseph R
