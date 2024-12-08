import pandas as pd
import requests
from matplotlib import pyplot as plt

from datetime import date, timedelta
d1 = date(2021, 1, 1)  # начальная дата
d2 = date(2025, 1, 1)
delta = 3

dates = [d1]

while d1 < d2:
  d1 += timedelta(delta)
  dates.append(d1)


tickers = ['YNDX','GAZP', 
           'ROSN', 'SBER', 'LKOH', 'SIBN', 'NVTK', 'PLZL', 
           'GMKN', 'YDEX', 'TATN', 'CHMF', 'SNGS', 'PHOR', 'NLMK', 'AKRN', 
           'OZON', 'UNAC', 'RUAL', 'TCSG', 'MGNT', 'SNGSP', 'MOEX', 'IRAO',
           'ALRS', 'VTBR', 'MAGN', 'MTSS', 'BANE', 'IRKT', 'CBOM']

def make_ticket(ticket):

  for i in range(len(dates) - 1):
    j = requests.get(f'http://iss.moex.com/iss/engines/stock/markets/shares/securities/{ticket}/candles.json?from={dates[i]}&till={dates[i+1]}').json()
    data = [{k : r[i] for i, k in enumerate(j['candles']['columns'])} for r in j['candles']['data']]
    frame = pd.DataFrame(data)
    ll = [frame['begin'].iloc[i][:7] for i in range(len(frame))]
    ll2 = [frame['begin'].iloc[i][8:11] for i in range(len(frame))]

    frame['month'] = ll
    frame['day'] = ll2

    try:
      frame1 = pd.concat([frame, frame1])
    except:
      frame1 = frame

  frame_dates = frame1.groupby('month')['end'].agg('max')
  frame2 = frame1.join(frame_dates, on='month', lsuffix='_caller', rsuffix='_other')
  frame3 = frame2[frame2['end_caller'] == frame2['end_other']]
  frame3 = frame3.reset_index()
  del frame3['index']

  frame_final = frame3.drop_duplicates()
  frame_final['ticket'] = [ticket for i in range(len(frame_final))]

  return frame_final

for i in tickers:
  try:
      frame = make_ticket(i)
      frame_top = pd.concat([frame, frame_top])
  except:
      frame_top = make_ticket(i)

file_path = "все_акции.xlsx" 
frame_top.to_excel(file_path, index=False)

import pandas as pd

def calculate_returns_and_beta(file_path, sheet_name):
    # Загрузка данных из Excel с указанного листа
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Убедитесь, что дата имеет правильный формат (например, datetime)
    df['дата'] = pd.to_datetime(df['дата'])

    # Сортировка данных по активам и датам
    df = df.sort_values(by=['ticket', 'дата'])

    # Создание пустых колонок для доходности
    df['Доходность актива'] = None
    df['Доходность индекса'] = None

    # Словарь для хранения ковариации, дисперсии и беты для каждого актива
    betas = {}

    # Вычисление доходности для каждого актива
    for asset in df['ticket'].unique():
        asset_data = df[df['ticket'] == asset].reset_index(drop=True)

        # Для каждой строки актива, начиная со второй (индекса 1)
        for i in range(1, len(asset_data)):
            # Доходность актива
            price_asset_today = asset_data.loc[i, 'close']
            price_asset_yesterday = asset_data.loc[i - 1, 'close']
            asset_data.loc[i, 'Доходность актива'] = (price_asset_today - price_asset_yesterday) / price_asset_yesterday

            # Доходность индекса
            price_index_today = asset_data.loc[i, 'индекс MOEX']
            price_index_yesterday = asset_data.loc[i - 1, 'индекс MOEX']
            asset_data.loc[i, 'Доходность индекса'] = (price_index_today - price_index_yesterday) / price_index_yesterday

        # Обновление данных в основном DataFrame
        df.update(asset_data)

        # Рассчитываем ковариацию между доходностью актива и индекса для этого актива
        covariance = asset_data[['Доходность актива', 'Доходность индекса']].cov().iloc[0, 1]

        # Рассчитываем дисперсию доходности индекса для этого актива
        variance_index = asset_data['Доходность индекса'].var()

        # Рассчитываем бета для этого актива
        beta = covariance / variance_index

        # Рассчитываем стандартное отклонение для актива и индекса
        std_asset = asset_data['Доходность актива'].std()
        std_index = df['Доходность индекса'].std()

        # Сохраняем бету в словарь
        betas[asset] = {
            'Ковариация': covariance,
            'Дисперсия индекса': variance_index,
            'Бета': beta,
            'Стандартное отклонение актива': std_asset,
            'Стандартное отклонение индекса': std_index

        }

    return df, betas

# Пример использования
file_path = 'все_акции.xlsx'  # Путь к вашему файлу
sheet_name = 'все'  # Имя листа

# Вычисление доходностей, ковариации, дисперсии и беты
result_df, betas = calculate_returns_and_beta(file_path, sheet_name)

# Сохраняем результат в новый Excel файл
result_df.to_excel('с отклонениями.xlsx', index=False)

# Выводим результаты
for asset, metrics in betas.items():
    print(f"Актив: {asset}")
    print(f"Ковариация: {metrics['Ковариация']}")
    print(f"Дисперсия индекса: {metrics['Дисперсия индекса']}")
    print(f"Бета: {metrics['Бета']}")
    print(f"Стандартное отклонение актива: {metrics['Стандартное отклонение актива']}")
    print(f"Стандартное отклонение индекса: {metrics['Стандартное отклонение индекса']}")
    print("-" * 40)
