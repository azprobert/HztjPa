import requests
import json

import xlrd
import xlwt







class JingQingXinXi:
  #'警情的基类'
  jqbh=''#警情编号
  jjsj=''#接警时间
  xzqydm=''#灾情区域（街道）
  jqfsdd=''#警情发生地点
  sjlbdm=''#事件类别代码
  sjlbmc=''#事件类别名称
  qtsjlbsm=''#其他事件类别说明
  hzbh=""# 火灾编号
  def __init__(self, jqbh):
      self.jqbh = jqbh
  def print(self):
    print('警情编号:'+ self.jqbh)
    print('接警时间:'+ self.jjsj)
    print('灾情区域（街道）:'+ self.xzqydm)
    print('警情发生地点:'+ self.jqfsdd)
    print('事件类别代码:'+ self.sjlbdm)
    print('事件类别名称:'+ self.sjlbmc)
    print('其他事件类别说明:'+ self.qtsjlbsm)
    print('火灾编号:'+ self.hzbh)


urlBase='http://stat.119.gov.cn/prod-api/police/situation/jqxxListDatails?jqbh='
url1 = "http://stat.119.gov.cn/prod-api/police/situation/jqxxListDatails?jqbh=J20212101107049"


payload={}
headers = {
  'Authorization': 'xx',
  'Cookie': 'JSESSIONID=xxx'
}

#f = open('d:/test.txt','a')
readbook = xlrd.open_workbook(r'd:\统计数据\综合查询-1627622500201.xlsx')
sheet = readbook.sheet_by_name('综合查询')#名字的方式
print(sheet.cell(1,3).value)#获取i行3列的表格值
nrows = sheet.nrows#行
print(nrows)
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet1')
style = xlwt.XFStyle()
saveFileLoadPath='Excel_Workbook_hztjLiaoning20210730.xls'





for i in range(1,nrows):
  jqbm = sheet.cell(i,0).value #'J20213702104601'
#  k = str(i).zfill(6)
  #print(k)
  url= urlBase+jqbm
  print(url)
  response = requests.request("GET", url, headers=headers, data=payload)

  print(response.text)
  res = response.text.encode("utf-8")
  #f.write(res+'\n')

  try:
    js = json.loads(res)
    # print(js)
    if (len(js['rows']) > 0):
      record = js['rows'][0]
      # print(record)

      worksheet.write(i, 0, jqbm, style)
      worksheet.write(i, 1, record['jjsj'], style)
      worksheet.write(i, 2, record['xzqydm'], style)
      worksheet.write(i, 3, record['jqfsdd'], style)
      worksheet.write(i, 4, record['sjlbdm'], style)
      worksheet.write(i, 5, record['sjlbmc'], style)
      worksheet.write(i, 6, record['qtsjlbsm'], style)

      #   zq = JingQingXinXi(jqbm)
      #   zq.jjsj = record['jjsj']
      #   zq.xzqydm = record['xzqydm']
      #    zq.jqfsdd = record['jqfsdd']
      #    zq.sjlbdm = record['sjlbdm']
      #    zq.sjlbmc = record['sjlbmc']
      #   zq.qtsjlbsm = record['qtsjlbsm']

      jlist = record['jlist']
      # print("jlist",jlist)
      for item in jlist:
        if (item['xfdwlx'] == '火灾报告'):
          worksheet.write(i, 7, item['cdbh'], style)
          hzbh =item['cdbh']
          urlHzbh='http://stat.119.gov.cn/prod-api/data/entry/zqxx/getHisDataDetail?zqbh='+hzbh
          print(urlHzbh)
          responseHzbh = requests.request("GET", urlHzbh, headers=headers, data=payload)
          resHzbh = responseHzbh.text.encode("utf-8")
          print(resHzbh)
          try:
            jsHzbh = json.loads(resHzbh)
          except:
            print('json读取有异常')
          if (len(jsHzbh['data']) > 0):
            # print('第2步')
            try:
              worksheet.write(i, 8, jsHzbh['data'][0]['zqlbdm'], style)  # 灾情类别代码
            except:
              print('灾情类别代码  有异常')
            try:
              worksheet.write(i, 9, jsHzbh['data'][0]['dwdm'], style)  # 单位代码
            except:
              print('单位代码  有异常')
            try:
              worksheet.write(i, 10, jsHzbh['data'][0]['xzqydm'], style)  # 行政区划代码
            except:
              print('行政区划代码  有异常')
            try:
              worksheet.write(i, 11, jsHzbh['data'][0]['qhcslb'], style)  # 起火场所类别
            except:
              print('起火场所类别  有异常')
            try:
              worksheet.write(i, 12, jsHzbh['data'][0]['qhcsms'], style)  # 起火场所描述
            except:
              print('起火场所描述  有异常')
            try:
              worksheet.write(i, 13, jsHzbh['data'][0]['qhcsdm'], style)  # 起火场所代码
            except:
              print('起火场所代码  有异常')
            try:
              worksheet.write(i, 14, jsHzbh['data'][0]['jjlxdm'], style)  # 经济类型代码
            except:
              print('经济类型代码  有异常')
            try:
              worksheet.write(i, 15, jsHzbh['data'][0]['sfsjycdm'], style)  # 是否世界遗产代码
            except:
              print('是否世界遗产代码  有异常')
            try:
              worksheet.write(i, 16, jsHzbh['data'][0]['qhwms'], style)  # 起火物描述
            except:
              print('起火物描述  有异常')
            try:
              worksheet.write(i, 17, jsHzbh['data'][0]['qhwfldm'], style)  # 起火物分类代码
            except:
              print('起火物分类代码  有异常')
            try:
              worksheet.write(i, 18, jsHzbh['data'][0]['hzyyms'], style)  # 火灾原因描述
            except:
              print('火灾原因描述  有异常')
            try:
              worksheet.write(i, 19, jsHzbh['data'][0]['hzyyfldm'], style)  # 火灾原因分类代码
            except:
              print('火灾原因分类代码  有异常')
            try:
              worksheet.write(i, 20, jsHzbh['data'][0]['czdcjg'], style)  # 处置调查经过
            except:
              print('处置调查经过  有异常')
            try:
              worksheet.write(i, 21, jsHzbh['data'][0]['fireProcess'], style)  # 处置经过
            except:
              print('处置经过  有异常')
            try:
              worksheet.write(i, 22, jsHzbh['data'][0]['qydm'], style)  # 区域代码
            except:
              print('处置经过  有异常')
            try:
              worksheet.write(i, 23, jsHzbh['data'][0]['sstd'], style)  # 疏散通道是否符合规定 1符合 2 不符合
            except:
              print('疏散通道是否符合规定  有异常')
            try:
              worksheet.write(i, 24, jsHzbh['data'][0]['qydm'], style)  # # 区域代码 1 城市市区 2县城城区 3集镇镇区 4农村  5开发区、旅游区 6其他
            except:
              print('区域代码  有异常')
            try:
              worksheet.write(i, 25, jsHzbh['data'][0]['jdjcqk'], style)  # 监督检查情况 1 消防 2派出所 3非监督
            except:
              print('监督检查情况  有异常')
            try:
              worksheet.write(i, 26, jsHzbh['data'][0]['sgqtdcbm'], style)  # 事故牵头调查部门 1 应急管理部门 2消防 3住建4公安 5交通 6其他
            except:
              print('事故牵头调查部门  有异常')
            try:
              worksheet.write(i, 27, jsHzbh['data'][0]['yjck'], style)  # 紧急出口是否符合规定 1符合 2 不符合
            except:
              print('紧急出口是否符合规定  有异常')
            try:
              worksheet.write(i, 28, jsHzbh['data'][0]['yjsszm'], style)  # 应急疏散照明是否符合规定 1符合 2 不符合
            except:
              print('应急疏散照明是否符合规定  有异常')
            try:
              worksheet.write(i, 29, jsHzbh['data'][0]['sflw'], style)  # 是否联网。 1联网 2未联网
            except:
              print('是否联网  有异常')
            try:
              if(len(jsHzbh['data'][0]['zqxxRyswList'])>0):
                deathCount =0;
                injuryCount =0;
                for rysw in jsHzbh['data'][0]['zqxxRyswList']:
                  if(rysw['swfl']==0):
                    deathCount = deathCount+1
                  else :
                    injuryCount = injuryCount +1
                worksheet.write(i, 30, deathCount, style)  # 伤亡人员性别
                worksheet.write(i, 31, injuryCount, style)  # 伤亡人员性别
                  # worksheet.write(i, 30, rysw['xb'], style)  # 伤亡人员性别
                  # worksheet.write(i, 31, rysw['sfzhm'], style)  # 伤亡人员身份证号
                  # worksheet.write(i, 32, rysw['zydm'], style)  # 伤亡人员职业代码
                  # worksheet.write(i, 32, rysw['jkzkdm'], style)  # 伤亡人员健康状况代码
                  # worksheet.write(i, 32, rysw['sjycddm'], style)  # 伤亡人员受教育程度代码
                  # worksheet.write(i, 32, rysw['swyydm'], style)  # 伤亡人员受伤或死亡原因代码
                  # worksheet.write(i, 32, rysw['rklydm'], style)  # 伤亡人员人口来源代码
                  # worksheet.write(i, 32, rysw['swwzdm'], style)  # 伤亡人员死亡位置代码




            except:
              print('是否联网  有异常')
      workbook.save(saveFileLoadPath)
  except IOError:
    print(  "Error: 没有找到文件或读取文件失败")
  else:
   # worksheet.write(i, 0, jqbm, style)
    workbook.save(saveFileLoadPath)



  # print(jlist)

  #zq.print()



