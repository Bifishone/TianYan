@echo off
chcp 65001 > nul
echo.
python3 TianYan.py --file ./CompanyName.txt --out-xlsx ./天眼查公司数据.xlsx --out-domains ./域名列表.txt --include-app --include-wechat --include-weibo
echo.
pause