import requests
import json

import xlrd
import xlwt



urlBase='http://stat.119.gov.cn/prod-api/police/situation/jqxxListDatails?jqbh='
url1 = "http://stat.119.gov.cn/prod-api/police/situation/jqxxListDatails?jqbh=J20212101107049"


payload={}
headers = {
  'Authorization': 'Bearer eyJhbGciOiJIUzUxMiJ9.eyJsb2dpbl91c2VyX2tleSI6IjkwYTdhZDc4LTEwYWYtNGI3Mi04NzEzLTAwYWExODA2Nzc1ZiJ9.yl7iB47GgOi_GlxuTfbIKHvYSTnsBBDPlCTtnX3zVucgPIbwjHsW3l1QyHbTIPzUdEWa3dfJnsNIx2h4abT7_A',
  'Cookie': 'JSESSIONID=4245CA6F4AE85788E00135BC4BAFE2ED'
}


url= 'http://stat.119.gov.cn/prod-api/data/entry/zqxx/getHisDataDetail?zqbh=Z20212101101245'
print(url)
response = requests.request("GET", url, headers=headers, data=payload)

#print(response.text)
res = response.text.encode("utf-8")
  #f.write(res+'\n')

try:
  js = json.loads(res)
  # print(js)
  print(js['data'])
  if(len(js['data'])>0):
    print(js['data'][0]['zqlbdm'])#灾情类别代码
    print(js['data'][0]['dwdm'])#单位代码
    print(js['data'][0]['xzqydm'])#行政区划代码
    print(js['data'][0]['qhcslb'])#起火场所类别
    print(js['data'][0]['qhcsms'])#起火场所描述
    print(js['data'][0]['qhcsdm'])#起火场所代码
    print(js['data'][0]['jjlxdm'])#经济类型代码
    print(js['data'][0]['sfsjycdm'])#是否世界遗产代码
    print(js['data'][0]['qhwms'])#起火物描述
    print(js['data'][0]['qhwfldm'])#起火物分类代码
    print(js['data'][0]['qhwmsItem1'])#起火物分类代码
    print(js['data'][0]['hzyyms'])#火灾原因描述
    print(js['data'][0]['hzyyfldm'])#火灾原因分类代码
    print(js['data'][0]['czdcjg'])#处置调查经过
    print(js['data'][0]['fireProcess'])#处置经过
    print(js['data'][0]['qydm'])#区域代码 1 城市市区 2县城城区 3集镇镇区 4农村  5开发区、旅游区 6其他
    #print(js['data'][0]['swsj'])#区域代码 1 事故现场 2非事故现场 3事故7天内死亡
    print(js['data'][0]['jdjcqk'])#监督检查情况 1 消防 2派出所 3非监督
    print(js['data'][0]['sgqtdcbm'])#事故牵头调查部门 1 应急管理部门 2消防 3住建4公安 5交通 6其他
    print(js['data'][0]['sstd'])#疏散通道是否符合规定 1符合 2 不符合
    print(js['data'][0]['yjck'])#紧急出口是否符合规定 1符合 2 不符合
    print(js['data'][0]['yjsszm'])#应急疏散照明是否符合规定 1符合 2 不符合
  #  print(js['data'][0]['jzyt'])#建筑用途 1居住 2公用 3工业 4农业
    print(js['data'][0]['sflw'])#是否联网。 1联网 2未联网
  #  print(js['data'][0]['xfssqk'])#是否安装消防设施。 1安装 2未安装
  else:
    a='1'
except :
  print(  "Error: 没有找到文件或读取文件失败")




