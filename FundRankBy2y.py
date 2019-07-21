import requests
import json
from bs4 import BeautifulSoup
import os
import xlwt
import time
import random

head = {};
head['User-Agent'] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36";

def main():
    writeEXCEL(getRankFund());

#获取相关基金数据
def getRankFund():
    #爬取两年内排名前一百的基金数据
    try:
        page=requests.get("https://danjuanapp.com/djapi/v3/filter/fund?type=1&order_by=2y&size=100&page=1",headers=head);
    except:
        pass;
    data=json.loads(page.text);
    return data;

#将获取到的数据写入Excel
def writeEXCEL(data):
    os.chdir("D://")
    print("正在写入......\n");
    f = xlwt.Workbook();
    sheet1 = f.add_sheet('股票型',cell_overwrite_ok = True);
  
    #写入第一行
    row0 = ["基金名称","基金代码","净值","日涨幅","基金管理者","近一周","近一月","近三月","近六月","今年以来","近一年","近两年","近三年","近五年","成立以来"];
    for i in range(0,len(row0)):
        sheet1.write(0,i,row0[i]);
  
    #将各个基金数据写入
    for i in range(0,100):
        fundCode=data['data']['items'][i]['fd_code'];
        time.sleep(random.random()*3)
        l=fundInfoProcess(fundCode);
        for j in range(0,14):
            sheet1.write(i+1,j,l[j]);
            
    f.save('基金数据.xls');

    print("写入完成!");


#获取某基金涨幅数据
def getFundTrend(fundCode):
    url1="https://danjuanapp.com/djapi/fund/derived/"+str(fundCode);
    url2="https://danjuanapp.com/djapi/fund/nav/history/"+str(fundCode);
    try:
        fundPage1=requests.get(url1,headers=head);
        fundPage2=requests.get(url2,headers=head);
    except:
        pass;
    fundInfoJson=json.loads(fundPage1.text);
    fundDailyInfoJson=json.loads(fundPage2.text);
    fundInfo = {};
    #获取日涨幅数据
    fundInfo['daily']=fundDailyInfoJson['data']['items'][0]['percentage'];
    #获取净值
    fundInfo['value']=fundDailyInfoJson['data']['items'][0]['value'];
    #获取其他涨幅数据
    try:
        fundInfo['1y']=fundInfoJson['data']['nav_grl1y'];#近一年涨幅
    except:
        fundInfo['1y']=0;
    try:
        fundInfo['1m']=fundInfoJson['data']['nav_grl1m'];#近一个月涨幅
    except:
        fundInfo['1m']=0;
    try:
        fundInfo['1w']=fundInfoJson['data']['nav_grl1w'];#近一周涨幅
    except:
         fundInfo['1w']=0;
    try:
        fundInfo['2y']=fundInfoJson['data']['nav_grl2y'];#近二年涨幅
    except:
        fundInfo['2y']=0;
    try:
        fundInfo['3m']=fundInfoJson['data']['nav_grl3m'];#近三个月涨幅
    except:
         fundInfo['3m']=0;
    try:
        fundInfo['6m']=fundInfoJson['data']['nav_grl6m'];#近六个月涨幅
    except:
        fundInfo['6m']=0;
    try:
        fundInfo['3y']=fundInfoJson['data']['nav_grl1y'];#近三年涨幅
    except:
        fundInfo['3y']=0;
    try:
        fundInfo['5y']=fundInfoJson['data']['nav_grl5y'];#近五年涨幅
    except:
        fundInfo['5y']=0;
    try:
        fundInfo['nav_grlty']=fundInfoJson['data']['nav_grlty'];#今年以来涨幅
    except:
        fundInfo['nav_grlty']=0;
    try:
        fundInfo['nav_grbase']=fundInfoJson['data']['nav_grbase'];#成立以来涨幅
    except:
        fundInfo['nav_grbase']=0;
    return fundInfo;

#获取某基金管理者数据
def getFundManager(fundCode):
    url = "https://danjuanapp.com/djapi/fund/"+str(fundCode);
    try:
        fundPage=requests.get(url,headers=head);
    except:
        pass;
    fundManagerInfo = json.loads(fundPage.text);
    return fundManagerInfo['data']['manager_name'];

#获取基金名称
def getFundName(fundCode):
    url="https://danjuanapp.com/djapi/fund/"+str(fundCode);
    fundPage=None;
    try:
        fundPage=requests.get(url,headers=head);
    except:
        pass;
    fundNameJson=json.loads(fundPage.text);
    fundName=fundNameJson['data']['fd_name'];
    return fundName;


#对某基金数据进行重新封装，返回一个list
def fundInfoProcess(fundCode):
    fundInfoList=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14];
    #获取某基金涨幅数据
    dic=getFundTrend(fundCode);
    #对基金数据进行封装
    fundInfoList[0]=getFundName(fundCode);
    fundInfoList[1]=fundCode ;
    fundInfoList[2]=dic['value'];
    fundInfoList[3]=dic['daily'];
    fundInfoList[4]=getFundManager(fundCode);
    fundInfoList[5]=dic['1w'];
    fundInfoList[6]=dic['1m'];
    fundInfoList[7]=dic['3m'];
    fundInfoList[8]=dic['6m'];
    fundInfoList[9]=dic['nav_grlty'];
    fundInfoList[10]=dic['1y'];
    fundInfoList[11]=dic['2y'];
    fundInfoList[12]=dic['3y'];
    fundInfoList[13]=dic['5y'];
    fundInfoList[14]=dic['nav_grbase'];
    return fundInfoList;


if __name__ == "__main__":
    main();