from analysis import Excel

file = 'C:/Users/brand/OneDrive - InsureGood LLC/Documents - Cedar ' \
       'Insights/applications/SentAnalysis_Excel/agency_nps.xlsx'


def run():
    analysis = Excel(file)
    analysis.fileParser()
    analysis.sentiment()
    analysis.fileWriter()

if __name__ == "__main__":
    run()
