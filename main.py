from yahoofinancials import YahooFinancials
import pandas as pd
import time


tickers_doc = pd.read_excel('Tickers.xlsx')


tickers = tickers_doc['Ticker'].to_list()
tickers = [str(item) for item in tickers]


tickers_error = []


time_all = time.time()
for ticker in tickers:
    start_time = time.time()
    print(ticker)
    try:
        yahoo_financials = YahooFinancials(ticker.replace(' ', ''))
        yahoo_financial_stmts = yahoo_financials.get_financial_stmts('annual', 'income')
        file_my = yahoo_financial_stmts['incomeStatementHistory'][ticker]

        date_1 = pd.DataFrame(file_my[0])
        date_2 = pd.DataFrame(file_my[1])
        date_3 = pd.DataFrame(file_my[2])

        date_1=date_1.set_index(date_1.groupby(level=0).cumcount(),append=True)
        date_2=date_2.set_index(date_2.groupby(level=0).cumcount(),append=True)
        date_3=date_3.set_index(date_3.groupby(level=0).cumcount(),append=True)

        df_prepared = pd.concat([date_1,date_2],1).reset_index(level=1,drop=True)
        df_prepared = df_prepared.set_index(df_prepared.groupby(level=0).cumcount(),append=True)

        df_head = pd.concat([df_prepared,date_3],1).reset_index(level=1,drop=True)


        name_doc = tickers_doc.loc[tickers_doc['Ticker']==ticker,['Name']]
        name_doc = str('File_finnance/' + name_doc.values[0][0]) + '.xlsx'


        writer = pd.ExcelWriter(name_doc, engine='xlsxwriter')

        # Write your DataFrame to a file
        df_head.to_excel(writer, 'Sheet1')

        # Save the result
        writer.save()
        print("--- %s seconds ---" % (time.time() - start_time))
    except (KeyError, TypeError):
        tickers_error.append(ticker)

print("--- %s seconds ---" % (time.time() - time_all))

