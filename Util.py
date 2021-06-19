class Util:

    def companyNameLookUp(companyName):
        companyDict = {
            'ALL': 'All Seasons Realty Corp.',
            'APL': 'Allianz-PNB Life Insurance, Inc. (APLII)',
            'ABI': 'Asia Brewery, Inc. (ABI) and Subsidiaries',
            'BHC': 'Basic Holdings Corp.',
            'CPH': 'Century Park Hotel',
            # "EPP": "Eton Properties Philippines, Inc. (EPPI), Subsidiaries",
            "EPP": "Eton Properties Philippines, Inc. (EPPI), Subsidiaries",
            'FFI': 'Foremost Farms, Inc.',
            'FTC': 'Fortune Tobacco Corp.',
            'GDC': 'Grandspan Development Corp.',
            'HII': 'Himmel Industries, Inc.',
            'LRC': 'Landcom Realty Corp.',
            'LTG': 'LT Group, Inc. (Parent Company)',
            'DIR': 'LTGC Directors',
            'MAC': 'MacroAsia Corp., Subsidiaries & Affiliates',
            'PAL': 'Philippine Airlines, Inc. (PAL), Subsidiaries and Affiliates',
            'PNB': 'Philippine National Bank (PNB) and Subsidiaries',
            'PMI': 'PMFTC',
            'RAP': 'Rapid Movers & Forwarders, Inc.',
            'TYK': 'Tan Yan Kee Foundation, Inc. (TYKFI)',
            'TDI': 'Tanduay Distillers, Inc. (TDI) and subsidiaries',
            'CHI': 'Charter House Inc.',
            'SPV': 'SPV-AMC Group',
            'TMC': 'Topkick Movers Corporation',
            'UNI': 'University of the East (UE)',
            'UER': 'University of the East Ramon Magsaysay Memorial Medical Center (UERMMMC)',
            'VMC': 'Victorias Milling Company, Inc. (VMC)',
            'ZHI': 'Zebra Holdings, Inc.',
            'STN': 'Sabre Travel Network Phils., Inc.',
            'OGC': 'OGC'
        }
        company_Code = ""
        for key, value in companyDict.items():
            if companyName.strip() == value:
                company_Code = key

        return company_Code
