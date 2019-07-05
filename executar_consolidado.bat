echo off
echo Starting now... $(stamp)
echo Generating month...
node gerar_month.js
echo Generating week_day...
node gerar_week_day.js
pause