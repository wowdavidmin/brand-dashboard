#!/usr/bin/env python3
import urllib.request, json
url = 'https://wttr.in/Seoul?format=j1'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
resp = urllib.request.urlopen(req, timeout=10)
data = json.loads(resp.read().decode())
cur = data['current_condition'][0]
print(f'----- 현재 서울 날씨 -----')
print(f'날씨: {cur["weatherDesc"][0]["value"]}')
print(f'온도: {cur["temp_C"]}°C (체감 {cur["FeelsLikeC"]}°C)')
print(f'습도: {cur["humidity"]}%')
print(f'풍속: {cur["windspeedKmph"]} km/h')
print()
print('----- 내일(7/1) 예보 -----')
for day in data['weather'][1:2]:
    print(f'최고 {day["maxtempC"]}°C / 최저 {day["mintempC"]}°C')
    print(f'시간별:')
    for h in day['hourly']:
        t = int(h['time'])
        print(f'  {t//100:02d}:00  {h["weatherDesc"][0]["value"]:10s} {h["tempC"]:>3s}°C  강수{h["chanceofrain"]}%')
