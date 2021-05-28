
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime as dt
import QuantLib as qt
import time
from flask import Flask, jsonify
from flask_restful import Api, Resource

app = Flask(__name__)
api = Api(app)

def getSheet(name_sheet):
    scopes = ['https://www.googleapis.com/auth/spreadsheets',
              "https://www.googleapis.com/auth/drive.file",
              "https://www.googleapis.com/auth/drive"]


    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'service_account.json', scopes)
    client = gspread.authorize(creds)
    sheet = client.open(name_sheet).sheet1
    return sheet


sheet1 = getSheet('Fiverr')
values = sheet1.col_values('2')


def dateToint(date):
    return dt.datetime.strptime(date, '%d/%m/%Y')


for i in range(0, 7):
    if(i == 1 or i == 0):
        values[i] = dateToint(values[i])
    else:
        values[i] = float(values[i])


VALUATION_DATE = 0
EXERCISE_DATE = 1
VOLATILITY = 2
UNDERLYING = 3
RISK_FREE_RATE = 4
DIVIDEND_RATE = 5
STRIKE = 6

def getdata(info):
    valuation_date = qt.Date(
        info[VALUATION_DATE].day, info[VALUATION_DATE].month, info[VALUATION_DATE].year)
    qt.Settings.instance().evaluationDate = valuation_date+2

    qt.calendar = qt.Canada()
    volatility = info[VOLATILITY]/100
    day_count = qt.Actual365Fixed()

    underlying = info[UNDERLYING]
    risk_free_rate = info[RISK_FREE_RATE]/100
    dividend_rate = info[DIVIDEND_RATE]/100

    exercise_date = qt.Date(
        info[EXERCISE_DATE].day, info[EXERCISE_DATE].month, info[EXERCISE_DATE].year)
    strike = info[STRIKE]
    option_type = qt.Option.Put

    payoff = qt.PlainVanillaPayoff(option_type, strike)
    exercise = qt.EuropeanExercise(exercise_date)
    european_option = qt.VanillaOption(payoff, exercise)

    spot_handle = qt.QuoteHandle(qt.SimpleQuote(underlying))

    flat_ts = qt.YieldTermStructureHandle(qt.FlatForward(
        valuation_date, risk_free_rate, day_count))
    dividend_yield = qt.YieldTermStructureHandle(
        qt.FlatForward(valuation_date, dividend_rate, day_count))
    flat_vol_ts = qt.BlackVolTermStructureHandle(qt.BlackConstantVol(
        valuation_date, qt.calendar, volatility, day_count))
    bsm_process = qt.BlackScholesMertonProcess(
        spot_handle, dividend_yield, flat_ts, flat_vol_ts)

    # European option
    european_option.setPricingEngine(qt.AnalyticEuropeanEngine(bsm_process))
    bs_price = european_option.NPV()
    sheet1.update('B10',bs_price)
    print("European option price is ", bs_price)

    # American option
    binomial_engine = qt.BinomialVanillaEngine(bsm_process, "crr", 50)
    am_exercise = qt.AmericanExercise(valuation_date, exercise_date)
    american_option = qt.VanillaOption(payoff, am_exercise)
    american_option.setPricingEngine(binomial_engine)
    crr_price = american_option.NPV()
    sheet1.update('B11',crr_price)
    print("American option price is ", crr_price)

    time.sleep(2)
    return 0

class callApi(Resource):
    def post(self):
        print(values)
        getdata(values)
        return jsonify({'data': 'Done'})

api.add_resource(callApi, "/callApi/")

if __name__ == '__main__':
    app.run()