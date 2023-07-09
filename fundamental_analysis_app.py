import os
import pandas as pd
from datetime import datetime
from calendar import monthrange
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import yfinance as yf
from utils import scrape_quick_links, scrape_table, get_active_href
import requests
import re
import streamlit as st
import sys

st.set_page_config(layout='wide')
st.title("Fundamental Analysis of the Stocks listed in NSE")
url = st.text_input("Enter stock url from www.moneycontrol.com website: ", '')

if url:
    current_directory = os.getcwd()
    stocks_price_directory = os.path.join(current_directory, 'Stock_Historical_Prices')

    # Replace the URL below with the actual URL of the Moneycontrol page you want to scrape
    # url = 'https://www.moneycontrol.com/india/stockpricequote/computers-software/infosys/IT'
    # url = 'https://www.moneycontrol.com/india/stockpricequote/pesticidesagro-chemicals/piindustries/PII'
    # url = 'https://www.moneycontrol.com/india/stockpricequote/miscellaneous/jubilantfoodworks/JF04'

    response = requests.get(url)
    html_source = response.text

    # Use regular expressions to find the nseId value in the HTML source code
    match = re.search(r'var nseId = "(.*?)";', html_source)

    # Extract the nseId value from the match object
    nse_id = match.group(1)

    # Extract the word NSE name from the nseId value
    nse_name = nse_id.split(':')[0]
    tick_name = nse_name + '.NS'

    # Following three lines are required for stocks, which dont have .NS at the end on yahoo finance
    # url = 'https://www.moneycontrol.com/india/stockpricequote/chemicals/blackroseindustries/BRI01'
    # nse_name = 'BLACKROSE'
    # tick_name = 'BLACKROSE.BO'

    s_name = yf.Ticker(tick_name)
    price_df = s_name.history(period="1000mo")
    if len(price_df) == 0:
        print('Ticker doesnt exists')
        st.warning('Unable to fetch this stock data. Please try other stocks. Thank you for understanding!', icon="🚨")
        sys.exit()
    else:
        st.warning('Working on your request..... Fetching data..... Thanks for your patience!')

    price_df['Date'] = price_df.index
    price_df = price_df.reset_index(drop=True)
    price_df = price_df[['Date', 'Open', 'High', 'Low', 'Close', 'Volume', 'Dividends', 'Stock Splits']]
    price_df = price_df.sort_values(by=['Date'], ascending=True).reset_index(drop=True)
    price_df['Date'] = price_df['Date'].dt.tz_localize(None)

    # file_name = os.path.join(stocks_price_directory, nse_name + '.xlsx')
    # price_df.to_excel(file_name)
    ######################################
    stocks_financials_directory = os.path.join(current_directory, 'Financials')
    CHECK_FOLDER = os.path.isdir(stocks_financials_directory)
    # If folder doesn't exist, then create it.
    if not CHECK_FOLDER:
        os.makedirs(stocks_financials_directory)
        print("created folder : ", stocks_financials_directory)
    else:
        print(stocks_financials_directory, "folder already exists.")

    file_path = os.path.join(stocks_financials_directory, nse_name + '.xlsx')
    if os.path.isfile(file_path):
        os.remove(file_path)


    sheets = ['Balance Sheet', 'Profit & Loss', 'Ratios']
    links_dict = scrape_quick_links(url)
    for sheet_name, financial_url in links_dict.items():
        if sheet_name in sheets:
            first_entry = True
            while financial_url:
                print(financial_url)
                if first_entry:
                    scrape_table(financial_url, nse_name, sheet_name, stocks_financials_directory)
                    first_entry = False
                else:
                    scrape_table(financial_url, nse_name, sheet_name, stocks_financials_directory, True)
                financial_url = get_active_href(financial_url)

    file_name = os.path.join(stocks_financials_directory, nse_name + '.xlsx')
    # create the subplots with two y-axes
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    # add the line chart trace to the first y-axis
    line_trace = go.Scatter(x=price_df["Date"], y=price_df["Close"], marker=dict(color='black'), name="Price")
    fig.add_trace(line_trace, secondary_y=False)

    for sheet_name in sheets:
        sheet_df = pd.read_excel(file_name, sheet_name=sheet_name)
        sheet_df = sheet_df.dropna(axis=1, how='all') # Drop nan columns

        # Find columns containing '\n' and drop them
        for col in sheet_df.columns:
            if '\n' in col:
                sheet_df.drop(col, axis=1, inplace=True)

        # Rename the 'name' column to 'full_name'
        sheet_df = sheet_df.rename(columns={'Unnamed: 1': 'Indicator'})

        df_transposed = sheet_df.transpose()
        # set the first row as the column names
        df_transposed.columns = df_transposed.iloc[0]
        # drop the first row since it is now the column names
        df_transposed = df_transposed.drop(df_transposed.index[0])

        # reverse the row positions
        df_reversed = df_transposed.iloc[::-1]

        # get dates as a new column
        df_reversed = df_reversed.reset_index(drop=False)
        df_reversed = df_reversed.rename(columns={'index': 'Date'})

        # set index to start from 0 onwards
        df_reversed = df_reversed.reset_index(drop=True)

        def convert_dates(month):
            parts = month.split('.')
            if len(parts) > 1:
                month = parts[0]
            date = datetime.strptime(month, '%b %y').date()
            last_dates = date.replace(day=monthrange(date.year, date.month)[1])

            return last_dates

        df_reversed['Date'] = df_reversed['Date'].apply(convert_dates)
        df_reversed['Date'] = pd.to_datetime(df_reversed['Date'])

        if sheet_name == 'Balance Sheet':
            indicators = ['Total Reserves and Surplus', 'Long Term Borrowings', 'Short Term Borrowings', 'Cash And Cash Equivalents']
        elif sheet_name == 'Profit & Loss':
            indicators = ['Total Operating Revenues', 'Profit/Loss For The Period', 'Basic EPS (Rs.)']
        elif sheet_name == 'Ratios':
            indicators = ['Book Value [ExclRevalReserve]/Share (Rs.)', 'Return on Networth / Equity (%)', 'Total Debt/Equity (X)']

        for indicator in indicators:
            # remove commas from the 'Numbers' column
            df_reversed[indicator] = df_reversed[indicator].replace(',', '', regex=True)
            # convert the 'Numbers' column to a numeric data type
            df_reversed[indicator] = pd.to_numeric(df_reversed[indicator])
            df_reversed = df_reversed.drop_duplicates(subset='Date', keep="first")
            df_reversed = df_reversed.sort_values(by=['Date'], ignore_index=True)

            # add the bar chart trace to the second y-axis with uniform width
            bar_trace = go.Bar(x=df_reversed["Date"], y=df_reversed[indicator], marker=dict(opacity=0.5),
                               name=indicator)
            fig.add_trace(bar_trace, secondary_y=True)
            # If we specify width=10000000000 in the bar_trace object, then no need to add the following statement to adjust the bargap
            fig.update_layout(bargap=0.5)

            # set the axis titles and layout
            fig.update_layout(title=nse_name, xaxis_title='Date', yaxis_title="Stock Price (Rs.)", yaxis2_title="Indicator (Rs. Cr.)")

            fig.update_traces(visible="legendonly", selector=lambda t: not t.name in ['Price', 'Total Reserves and Surplus'])

    # fig.show()
    st.plotly_chart(fig, use_container_width=True)

    # Delete all files in the directory
    def delete_files_in_directory(directory_path):
        # Iterate over all files in the directory
        for filename in os.listdir(directory_path):
            file_path = os.path.join(directory_path, filename)

            # Check if the current item is a file
            if os.path.isfile(file_path):
                try:
                    # Delete the file
                    os.remove(file_path)
                    print(f"Deleted: {file_path}")
                except Exception as e:
                    print(f"Failed to delete: {file_path}")
                    print(e)

    # Call the function to delete files in the directory
    delete_files_in_directory(stocks_financials_directory)


