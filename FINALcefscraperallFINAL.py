import requests
import pandas as pd
from bs4 import BeautifulSoup
from functools import reduce

df = pd.read_excel('C:/Users/Jacob Steenhuysen/Downloads/CEF Tickers.xlsx', sheet_name='Sheet1')

tickers_list = df['Ticker'].tolist()




headers = {
    "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:80.0) Gecko/20100101 Firefox/80.0"
}

#tickers_list = ["NIE", "ZTR", "SCD", "GLV"]

dfs1 = []
dfs2 = []
dfs3 = []
dfs4 = []
dfs5 = []
dfs6 =[]
dfs7 = []
dfs8 = []
dfs9 = []
dfs10 = []
dfs11 = []
dfs12 = []
dfs13 = []


for ticker in tickers_list:
    try:
        print("Downloading ticker {}".format(ticker))
        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_SummaryGrid')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:]
        df.columns = new_header
        df = df.unstack().to_frame().sort_index(level=1).T
        df.columns = df.columns.map('_'.join)
        #df.set_index([f'{ticker}'], inplace = True)
        df["Ticker"] = ticker
        dfs1.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucFundBasics_dvCapitalStructure')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df["Ticker"] = ticker
        dfs2.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucFundBasics_dvFB1')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        # df = df[1:].astype(str)
        # df.columns = new_header
        df["Ticker"] = ticker
        dfs3.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucFundBasics_dvFB2')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        # df = df[1:].astype(str)
        # df.columns = new_header
        df["Ticker"] = ticker
        dfs4.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucDistributions_dvFB1')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        # df = df[1:].astype(str)
        # df.columns = new_header
        df["Ticker"] = ticker
        dfs5.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucPricing_DiscountGrid')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df = df.unstack().to_frame().sort_index(level=1).T
        df.columns = df.columns.map('_'.join)
        #df.set_index([f'{ticker}'], inplace = True)
        df["Ticker"] = ticker
        dfs6.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucPricing_ZScoreGridView')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df = df.unstack().to_frame().sort_index(level=1).T
        df.columns = df.columns.map('_'.join)
        #df.set_index([f'{ticker}'], inplace = True)
        df["Ticker"] = ticker
        dfs7.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucFundBasics_dvCapitalStructure')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df["Ticker"] = ticker
        dfs8.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucFundBasics_dvFB1')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df["Ticker"] = ticker
        dfs9.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucPricing_DiscountGrid')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df = df.unstack().to_frame().sort_index(level=1).T
        df.columns = df.columns.map('_'.join)
        #df.set_index([f'{ticker}'], inplace = True)
        df["Ticker"] = ticker
        dfs10.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucPricing_ZScoreGridView')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df = df.unstack().to_frame().sort_index(level=1).T
        df.columns = df.columns.map('_'.join)
        #df.set_index([f'{ticker}'], inplace = True)
        df["Ticker"] = ticker
        dfs11.append(df)


        #MAYBE THIS WILL FLAG
        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucPortChar_pcSector_SectorGrid')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df["Ticker"] = ticker
        df.fillna(0)
        dfs12.append(df)

        soup = BeautifulSoup(requests.get(f"https://www.cefconnect.com/fund/{ticker}", headers=headers).content,"html.parser",)
        df = pd.read_html(str(soup.select_one('#ContentPlaceHolder1_cph_main_cph_main_ucPortChar_pcCountry_CountryGrid')))[0]
        df = pd.DataFrame.transpose(df)
        new_header = df.iloc[0].astype(str)
        df = df[1:].astype(str)
        df.columns = new_header
        df["Ticker"] = ticker
        df.fillna(0)
        dfs13.append(df)
    except:
        pass




final_dfs1 = pd.concat(dfs1)
final_dfs2 = pd.concat(dfs2)
final_dfs3 = pd.concat(dfs3)
final_dfs4 = pd.concat(dfs4)
final_dfs5 = pd.concat(dfs5)
final_dfs6 = pd.concat(dfs6)
final_dfs7 = pd.concat(dfs7)
final_dfs8 = pd.concat(dfs8)
final_dfs9 = pd.concat(dfs9)
final_dfs10 = pd.concat(dfs10)
final_dfs11 = pd.concat(dfs11)
final_dfs12 = pd.concat(dfs12)
final_dfs13 = pd.concat(dfs13)



final_dfs1 = final_dfs1.set_index(['Ticker'])
final_dfs2 = final_dfs2.set_index(['Ticker'])
final_dfs3 = final_dfs3.set_index(['Ticker'])
final_dfs4 = final_dfs4.set_index(['Ticker'])
final_dfs5 = final_dfs5.set_index(['Ticker'])
final_dfs6 = final_dfs6.set_index(['Ticker'])
final_dfs7 = final_dfs7.set_index(['Ticker'])
final_dfs8 = final_dfs8.set_index(['Ticker'])
final_dfs9 = final_dfs9.set_index(['Ticker'])
final_dfs10 = final_dfs10.set_index(['Ticker'])
final_dfs11 = final_dfs11.set_index(['Ticker'])
final_dfs12 = final_dfs12.set_index(['Ticker'])
final_dfs13 = final_dfs13.set_index(['Ticker'])

final_dfs = [final_dfs5, final_dfs1, final_dfs2, final_dfs3, final_dfs4, final_dfs6, final_dfs7, final_dfs8, final_dfs9, final_dfs10, final_dfs11, final_dfs12, final_dfs13]

df_merged = reduce(lambda  left,right: pd.merge(left,right,on=['Ticker'], how='outer'), final_dfs).fillna('void')




export_excel = df_merged.to_excel(r'C:/Users/Jacob Steenhuysen/Downloads/CEF Connect Full Scrape.xlsx', sheet_name='Sheet1', index= True)


#print(df_merged)
