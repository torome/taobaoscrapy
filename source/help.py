# -*- coding:utf-8 -*-
import urllib.request, urllib.parse, http.cookiejar
import os, time,re
import http.cookies
import xlsxwriter as wx
from PIL import Image

__author__ = 'hunterhug'
# http://python.jobbole.com/81344/
# 拆分JSON
import xml.dom.minidom
import json
from openpyxl import Workbook
from openpyxl import load_workbook


# 找出文件夹下所有html后缀的文件
def listfiles(rootdir, prefix='.json'):
	file = []
	for parent, dirnames, filenames in os.walk(rootdir):
		for filename in filenames:
			if filename.endswith(prefix):
				file.append(parent.replace('\\','/')+'/'+filename)
	path=file[0].split('/')[-2]
	return  file,path

def writeexcel(path,dealcontent):
	workbook = wx.Workbook(path)
	top = workbook.add_format({'border':1,'align':'center','bg_color':'white','font_size':11,'font_name': '微软雅黑'})
	red = workbook.add_format({'font_color':'white','border':1,'align':'center','bg_color':'800000','font_size':11,'font_name': '微软雅黑','bold':True})
	image = workbook.add_format({'border':1,'align':'center','bg_color':'white','font_size':11,'font_name': '微软雅黑'})
	formatt=top
	formatt.set_align('vcenter') #设置单元格垂直对齐
	worksheet = workbook.add_worksheet()        #创建一个工作表对象
	width=len(dealcontent[0])
	worksheet.set_column(0,width,38.5)            #设定列的宽度为22像素
	for i in range(0,len(dealcontent)):
		if i==0:
			formatt=red
		else:
			formatt=top
		for j in range(0,len(dealcontent[i])):
			if i!=0 and j==len(dealcontent[i])-1:
				if dealcontent[i][j]=='':
					worksheet.write(i,j,' ',formatt)
				else:
					try:
						worksheet.insert_image(i,j,dealcontent[i][j])
					except:
						worksheet.write(i,j,' ',formatt)
			else:
				if dealcontent[i][j]:
					worksheet.write(i,j,dealcontent[i][j].replace(' ',''),formatt)
				else:
					worksheet.write(i,j,'无',formatt)
	workbook.close()


def getHtml(url,postdata={}):
	"""
    抓取网页：支持cookie
    第一个参数为网址，第二个为POST的数据

    """
	# COOKIE文件保存路径
	filename = 'cookie.txt'

	# 声明一个MozillaCookieJar对象实例保存在文件中
	cj = http.cookiejar.MozillaCookieJar(filename)
	# cj =http.cookiejar.LWPCookieJar(filename)

	# 从文件中读取cookie内容到变量
	# ignore_discard的意思是即使cookies将被丢弃也将它保存下来
	# ignore_expires的意思是如果在该文件中 cookies已经存在，则覆盖原文件写
	# 如果存在，则读取主要COOKIE
	if os.path.exists(filename):
		cj.load(filename, ignore_discard=True, ignore_expires=True)
	# 读取其他COOKIE
	if os.path.exists('../subcookie.txt'):
		cookie = open('../subcookie.txt', 'r').read()
	else:
		cookie='ddd'
	# 建造带有COOKIE处理器的打开专家
	opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

	# 打开专家加头部
	opener.addheaders = [('User-Agent',
						  'Mozilla/5.0 (iPad; U; CPU OS 4_3_3 like Mac OS X; en-us) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8J2 Safari/6533.18.5'),
						 ('Referer',
						  'http://s.m.taobao.com'),
						 ('Host', 'h5.m.taobao.com'),
						 ('Cookie',cookie)]

	# 分配专家
	urllib.request.install_opener(opener)
	# 有数据需要POST
	if postdata:
		# 数据URL编码
		postdata = urllib.parse.urlencode(postdata)

		# 抓取网页
		html_bytes = urllib.request.urlopen(url, postdata.encode()).read()
	else:
		html_bytes = urllib.request.urlopen(url).read()

	# 保存COOKIE到文件中
	cj.save(ignore_discard=True, ignore_expires=True)
	return html_bytes

# 去除标题中的非法字符 (Windows)
def validateTitle(title):
	rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/\:*?"<>|'
	new_title = re.sub(rstr, "", title)
	return new_title

# 递归创建文件夹
def createjia(path):
	try:
		os.makedirs(path)
	except:
		print('目录已经存在：'+path)

def timetochina(longtime,formats='{}天{}小时{}分钟{}秒'):
	day=0
	hour=0
	minutue=0
	second=0
	try:
		if longtime>60:
			second=longtime%60
			minutue=longtime//60
		else:
			second=longtime	
		if minutue>60:
			hour=minutue//60
			minutue=minutue%60
		if hour>24:
			day=hour//24
			hour=hour%24
		return formats.format(day,hour,minutue,second)
	except:
		raise Exception('时间非法')

def begin():
    sangjin = '''
		-----------------------------------------
		| 欢迎使用自动抓取手机淘宝关键字程序   	|
		| 时间：2015年12月23日                  |
		| 新浪微博：一只尼玛                    |
		| 微信/QQ：569929309                    |
		-----------------------------------------
	'''
    print(sangjin)


if __name__ == '__main__':
	begin()
	a=time.clock()
	today=time.strftime('%Y%m%d%H%M', time.localtime())
	root='../help'
	try:
		files,path= listfiles(root, '.json')
	except:
		print('错误！！请把需要重新处理的文件拉到Help文件夹')
		input('')
		exit()
	needpic=input("需要下载图片吗？需要请按1，否则按其他键：")
	mulu=root+'/'+path+'pic'+today
	if needpic=='1':
		createjia(mulu)
	total = []
	total.append(['页数', '店名', '商品标题', '商品打折价', '发货地址', '评论数', '原价', '手机折扣', '售出件数', '政策享受', '付款人数', '金币折扣','URL地址','图像URL','图像'])
	for filename in files:
		print('正在处理：'+filename)
		try:
			doc = open(filename, 'rb')
			doccontent = doc.read().decode('utf-8', 'ignore')
			product = doccontent.replace(' ', '').replace('\n', '')
			product = json.loads(product)
			onefile = product['listItem']
		except:
			print('抓不到' + filename)
			continue
		for item in onefile:
			itemlist = [filename, item['nick'], item['title'], item['price'], item['location'], item['commentCount']]
			itemlist.append(item['originalPrice'])
			itemlist.append(item['mobileDiscount'])
			itemlist.append(item['sold'])
			itemlist.append(item['zkType'])
			itemlist.append(item['act'])
			itemlist.append(item['coinLimit'])
			itemlist.append(item['auctionURL'])
			picpath=item['pic_path'].replace('60x60','720x720')
			itemlist.append(picpath)
			#http://g.search2.alicdn.com/img/bao/uploaded/i4/i4/TB13O7bJVXXXXbJXpXXXXXXXXXX_%21%210-item_pic.jpg_180x180.jpg
			if needpic=='1':
				url=urllib.parse.quote(picpath).replace('%3A',':')
				urllib.request.urlcleanup()
				try:
					pic=urllib.request.urlopen(url)
					picno=time.strftime('%H%M%S', time.localtime())
					filenamep=mulu+'/'+picno+validateTitle(item['nick']+'-'+item['title'])
					filenamepp=filenamep+'.jpeg'
					sfilename=filenamep+'s.jpeg'
					filess=open(filenamepp,'wb')
					filess.write(pic.read())
					filess.close()
					img = Image.open(filenamepp)
					w, h = img.size
					size=w/6,h/6
					img.thumbnail(size, Image.ANTIALIAS)
					img.save(sfilename,'jpeg')
					itemlist.append(sfilename)
					print('抓到图片：'+sfilename)
				except Exception as e:
					if hasattr(e, 'code'):
						print('页面不存在或时间太长.')
						print('Error code:', e.code)
					elif hasattr(e, 'reason'):
							print("无法到达主机.")
							print('Reason:  ', e.reason)
					else:
						print(e)
					itemlist.append('')
			else:
				itemlist.append('')
			# print(itemlist)
			total.append(itemlist)
	if len(total) > 1:
		writeexcel(root+'/'+path+ '淘宝手机商品.xlsx', total)
	else:
		print('什么都抓不到')
	b=time.clock()
	print('运行时间：'+timetochina(b-a))
	input('请关闭窗口')
