import pandas as pd

# Load the Excel file
file_path = 'updated_currency.xlsx'
df = pd.read_excel(file_path)

# Define the tickers for companies where both ADR and foreign tickers are available
tickers = {
    "Toyota": {"adr": "TM", "foreign": "7203.T"},
    "111 INC": {"adr": "YI", "foreign": "1112.HK"},
    "3I GROUP PLC": {"adr": "TGOPF", "foreign": "III.L"},
    "4D PHARMA PLC": {"adr": "LBPS", "foreign": "DDDD.L"},
    "A.P. MOLLER - MAERSK AS": {"adr": "AMKBY", "foreign": "MAERSKB.CO"},
    "AAC TECHNOLOGIES HOLDINGS INC.": {"adr": "AACAY", "foreign": "2018.HK"},
    "ABB LTD.": {"adr": "ABB", "foreign": "ABBN.S"},
    "ABN AMRO BANK N.V.": {"adr": "ABNRY", "foreign": "ABN.AS"},
    "AIA GROUP LIMITED": {"adr": "AAGIY", "foreign": "1299.HK"},
    "AIR FRANCE-KLM": {"adr": "AFLYY", "foreign": "AF.PA"},
    "AIR LIQUIDE SA": {"adr": "AIQUY", "foreign": "AI.PA"},
    "AJINOMOTO CO. INC": {"adr": "AJINY", "foreign": "2802.T"},
    "AKZO NOBEL N.V.": {"adr": "AKZOY", "foreign": "AKZA.AS"},
    "ALIBABA GROUP HOLDING LTD": {"adr": "BABA", "foreign": "9988.HK"},
    "ALLIANZ SE": {"adr": "ALIZY", "foreign": "ALV.DE"},
    "ANHEUSER-BUSCH IN BEV SA/NV": {"adr": "BUD", "foreign": "ABI.BR"},
    "BAIDU INC": {"adr": "BIDU", "foreign": "9888.HK"},
    "BARCLAYS PLC": {"adr": "BCS", "foreign": "BARC.L"},
    "BASF SE": {"adr": "BASFY", "foreign": "BAS.DE"},
    "BAYER AG": {"adr": "BAYRY", "foreign": "BAYN.DE"},
    "BRITISH AMERICAN TOBACCO": {"adr": "BTI", "foreign": "BATS.L"},
    "CEMEX S.A.B. DE C.V.": {"adr": "CX", "foreign": "CEMEXCPO.MX"},
    "CHINA MOBILE LIMITED": {"adr": "CHL", "foreign": "0941.HK"},
    "CRH PLC": {"adr": "CRH", "foreign": "CRG.IR"},
    "DAIMLER AG": {"adr": "DDAIF", "foreign": "DAI.DE"},
    "DEUTSCHE BANK AG": {"adr": "DB", "foreign": "DBK.DE"},
    "DIAGEO PLC": {"adr": "DEO", "foreign": "DGE.L"},
    "ENI S.P.A.": {"adr": "E", "foreign": "ENI.MI"},
    "ERICSSON": {"adr": "ERIC", "foreign": "ERIC-B.ST"},
    "FERROVIAL S.A.": {"adr": "FRRVY", "foreign": "FER.MC"},
    "GLAXOSMITHKLINE PLC": {"adr": "GSK", "foreign": "GSK.L"},
    "GRUPO AEROPORTUARIO DEL PACIFICO": {"adr": "PAC", "foreign": "GAPB.MX"},
    "HONDA MOTOR CO. LTD.": {"adr": "HMC", "foreign": "7267.T"},
    "HSBC HOLDINGS PLC": {"adr": "HSBC", "foreign": "HSBA.L"},
    "ICICI BANK LTD.": {"adr": "IBN", "foreign": "ICICIBANK.NS"},
    "ING GROEP N.V.": {"adr": "ING", "foreign": "INGA.AS"},
    "ISHARES MSCI JAPAN ETF": {"adr": "EWJ", "foreign": "JPXN.S"},
    "ISHARES MSCI SOUTH KOREA ETF": {"adr": "EWY", "foreign": "KOXN.S"},
    "ITAU UNIBANCO HOLDING S.A.": {"adr": "ITUB", "foreign": "ITUB4.SA"},
    "JAPAN TOBACCO INC.": {"adr": "JAPAY", "foreign": "2914.T"},
    "JD.COM, INC": {"adr": "JD", "foreign": "9618.HK"},
    "KAO CORPORATION": {"adr": "KCRPY", "foreign": "4452.T"},
    "KYOCERA CORPORATION": {"adr": "KYOCY", "foreign": "6971.T"},
    "MITSUBISHI UFJ FINANCIAL GROUP, INC.": {"adr": "MUFG", "foreign": "8306.T"},
    "NESTLE S.A.": {"adr": "NSRGY", "foreign": "NESN.S"},
    "NOVARTIS AG": {"adr": "NVS", "foreign": "NOVN.S"},
    "NOVO NORDISK": {"adr": "NVO", "foreign": "NOVOB.CO"},
    "ORANGE S.A.": {"adr": "ORAN", "foreign": "ORA.PA"},
    "PETROLEO BRASILEIRO S.A.": {"adr": "PBR", "foreign": "PETR4.SA"},
    "REPSOL S.A.": {"adr": "REPYY", "foreign": "REP.MC"},
    "RIO TINTO PLC": {"adr": "RIO", "foreign": "RIO.L"},
    "ROYAL DUTCH SHELL PLC": {"adr": "RDS.A", "foreign": "RDSA.L"},
    "SAMSUNG ELECTRONICS": {"adr": "SSNLF", "foreign": "005930.KS"},
    "SANOFI": {"adr": "SNY", "foreign": "SAN.PA"},
    "SIEMENS AG": {"adr": "SIEGY", "foreign": "SIE.DE"},
    "SK TELECOM CO LTD": {"adr": "SKM", "foreign": "017670.KS"},
    "SOCIETE GENERALE S.A.": {"adr": "SCGLY", "foreign": "GLE.PA"},
    "SONY GROUP CORPORATION": {"adr": "SONY", "foreign": "6758.T"},
    "TAIWAN SEMICONDUCTOR MANUFACTURING": {"adr": "TSM", "foreign": "2330.TW"},
    "TELEFONICA S.A.": {"adr": "TEF", "foreign": "TEF.MC"},
    "TOYOTA MOTOR CORPORATION": {"adr": "TM", "foreign": "7203.T"},
    "UNICREDIT S.P.A.": {"adr": "UNCFF", "foreign": "UCG.MI"},
    "VEOLIA ENVIRONNEMENT": {"adr": "VEOEY", "foreign": "VIE.PA"},
    "VODAFONE GROUP PLC": {"adr": "VOD", "foreign": "VOD.L"},
    "WPP PLC.": {"adr": "WPPGY", "foreign": "WPP.L"},
    "ZURICH INSURANCE GROUP AG LTD": {"adr": "ZURVY", "foreign": "ZURN.S"},
    "AXA SA": {"adr": "AXAHY", "foreign": "CS.PA"},
    "BHP GROUP LIMITED": {"adr": "BHP", "foreign": "BHP.AX"},
    "BNP PARIBAS": {"adr": "BNPQY", "foreign": "BNP.PA"},
    "BP PLC": {"adr": "BP", "foreign": "BP.L"},
    "CREDIT SUISSE GROUP AG": {"adr": "CS", "foreign": "CSGN.S"},
    "DEUTSCHE TELEKOM AG": {"adr": "DTEGY", "foreign": "DTE.DE"},
    "E.ON SE": {"adr": "EONGY", "foreign": "EOAN.DE"},
    "ENEL SPA": {"adr": "ENLAY", "foreign": "ENEL.MI"},
    "GSK PLC": {"adr": "GSK", "foreign": "GSK.L"},
    "HITACHI LTD.": {"adr": "HTHIY", "foreign": "6501.T"},
    "ICL GROUP LTD": {"adr": "ICL", "foreign": "ICL.TA"},
    "KONINKLIJKE PHILIPS N.V.": {"adr": "PHG", "foreign": "PHIA.AS"},
    "MIZUHO FINANCIAL GROUP INC.": {"adr": "MFG", "foreign": "8411.T"},
    "NIPPON TELEGRAPH AND TELEPHONE CORPORATION": {"adr": "NTTYY", "foreign": "9432.T"},
    "NTT DOCOMO INC.": {"adr": "DCMYY", "foreign": "9437.T"},
    "SANTANDER BANK POLSKA SA": {"adr": "BPYPY", "foreign": "BZW.WA"}
}
# Filter the DataFrame to only include rows where the 'Company Name' is in the tickers dictionary
df = df[df['Company Name'].isin(tickers.keys())]

# Update the dataframe with the tickers
for company, ticker_info in tickers.items():
    df.loc[df['Company Name'] == company, 'adr'] = ticker_info['adr']
    df.loc[df['Company Name'] == company, 'foreign'] = ticker_info['foreign']

# Save the updated dataframe
df.to_excel('updated_tickers.xlsx', index=False)