# -*- coding:utf-8 -*-
import urllib.request, urllib.parse, http.cookiejar
import os, time,re
import http.cookies
import xlsxwriter as wx
from PIL import Image
import pymysql
import socket
__author__ = 'hunterhug'
# http://python.jobbole.com/81344/
# 拆分JSON
import xml.dom.minidom
import json
from openpyxl import Workbook
from openpyxl import load_workbook

def password():
	print('请输入你的账号和密码')
	user=input('账号：')
	pwd=input('密码：')
	if user=='jinhan' and pwd=='6833066':
		print('欢迎你：'+user)
		return
	try:
		mysql = pymysql.connect(host="192.168.1.177", user="dataman", passwd="123456",db='qingmu', charset="utf8")
		cur = mysql.cursor()
		isuser="SELECT * FROM mtaobao where user='{0}' and pwd='{1}'".format(user,pwd)
		cur.execute(isuser)
		mysql.commit()
		if cur.fetchall():
			print('欢迎你：'+user)
			localIP = socket.gethostbyname(socket.gethostname())#这个得到本地ip
			ipList = socket.gethostbyname_ex(socket.gethostname())
			s=''
			for i in ipList:
				if i != localIP and i!=[]:
					s=s+(str)(i)
			timesss=time.strftime('%Y%m%d-%H%M%S', time.localtime())
			update="UPDATE mtaobao SET `times` = `times`+1,`dates`='{0}',`ip` ='{1}' where user='{2}'".format(timesss,s.replace("'",''),user)
			#print(update)
			cur.execute(update)
			mysql.commit()
			cur.close()
			mysql.close()
			return
		else:
			raise
	except Exception as e:
		#print(e)
		mysql.rollback()
		cur.close()
		mysql.close()
		print('密码错误')
		password()

# 找出文件夹下所有html后缀的文件
def listfiles(rootdir, prefix='.xml'):
	file = []
	for parent, dirnames, filenames in os.walk(rootdir):
		if parent == rootdir:
			for filename in filenames:
				if filename.endswith(prefix):
					file.append(rootdir + '/' + filename)
			return file
		else:
			pass

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
	

def getHtml(url,daili='',postdata={}):
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
	proxy_support = urllib.request.ProxyHandler({'http':'http://'+daili})
	# 开启代理支持
	if daili:
		print('代理:'+daili+'启动')
		opener = urllib.request.build_opener(proxy_support, urllib.request.HTTPCookieProcessor(cj), urllib.request.HTTPHandler)
	else:
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
	password()
	today=time.strftime('%Y%m%d', time.localtime())
	a=time.clock()
	keyword = input('请输入关键字：')
	sort = input('按销量优先请按1，按价格低到高抓取请按2，价格高到低按3，信用排序按4，综合排序按5：')
	try:
		pages =int(input('需要抓取的页数（默认100页）：'))
		if pages>100 or pages<=0:
			print('页数应该在1-100之间')
			pages=100
	except:
		pages=100
	try:
		man=int(input('请设置抓取暂停时间：默认4秒（4）：'))
		if man<=0:
			man=4
	except:
		man=4
	zp=input('抓取图片按1，不抓取按2：')
	if sort == '1':
		sortss = '_sale'
	elif sort == '2':
		sortss = 'bid'
	elif sort=='3':
		sortss='_bid'
	elif sort=='4':
		sortss='_ratesum'
	elif sort=='5':
		sortss=''
	else:
		sortss = '_sale'
	namess=time.strftime('%Y%m%d%H%S', time.localtime())
	root = '../data/'+today+'/'+namess+keyword
	roota='../excel/'+today
	mulu='../image/'+today+'/'+namess+keyword
	createjia(root)
	createjia(roota)
	for page in range(0, pages):
		time.sleep(man)
		print('暂停'+str(man)+'秒')
		if sortss=='':
			postdata = {
				'event_submit_do_new_search_auction': 1,
				'search': '提交查询',
				'_input_charset': 'utf-8',
				'topSearch': 1,
				'atype': 'b',
				'searchfrom': 1,
				'action': 'home:redirect_app_action',
				'from': 1,
				'q': keyword,
				'sst': 1,
				'n': 20,
				'buying': 'buyitnow',
				'm': 'api4h5',
				'abtest': 16,
				'wlsort': 16,
				'style': 'list',
				'closeModues': 'nav,selecthot,onesearch',
				'page': page
			}
		else:
			postdata = {
				'event_submit_do_new_search_auction': 1,
				'search': '提交查询',
				'_input_charset': 'utf-8',
				'topSearch': 1,
				'atype': 'b',
				'searchfrom': 1,
				'action': 'home:redirect_app_action',
				'from': 1,
				'q': keyword,
				'sst': 1,
				'n': 20,
				'buying': 'buyitnow',
				'm': 'api4h5',
				'abtest': 16,
				'wlsort': 16,
				'style': 'list',
				'closeModues': 'nav,selecthot,onesearch',
				'sort': sortss,
				'page': page
			}
		postdata = urllib.parse.urlencode(postdata)
		taobao = "http://s.m.taobao.com/search?" + postdata
		print(taobao)
		try:
			content1 = getHtml(taobao)
			file = open(root + '/' + str(page) + '.json', 'wb')
			file.write(content1)
		except Exception as e:
				if hasattr(e, 'code'):
					print('页面不存在或时间太长.')
					print('Error code:', e.code)
				elif hasattr(e, 'reason'):
						print("无法到达主机.")
						print('Reason:  ', e.reason)
				else:
					print(e)

	# files=listfiles('201512171959','.json')
	files = listfiles(root, '.json')
	total = []
	total.append(['页数', '店名', '商品标题', '商品打折价', '发货地址', '评论数', '原价', '手机折扣', '售出件数', '政策享受', '付款人数', '金币折扣','URL地址','图像URL','图像'])
	for filename in files:
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
			if zp=='1':
				if os.path.exists(mulu):
					pass
				else:
					createjia(mulu)
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
		writeexcel(roota +'/'+namess+keyword+ '淘宝手机商品.xlsx', total)
	else:
		print('什么都抓不到')
	b=time.clock()
	print('运行时间：'+timetochina(b-a))
	input('请关闭窗口')
