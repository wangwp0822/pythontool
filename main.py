#coding=utf-8

import cookielib, urllib2, urllib, re, os, xlwt
from urllib import urlencode
from config import _loginweb,_userinfo,_desgion,_client,_server
import time, datetime

#
# other way
# http://stackoverflow.com/questions/10957522/how-to-convert-string-datetime-to-timestamp-in-python
# >>> timestamp = time.mktime(time.strptime(datetimestring, '%a, %d %b %Y %H:%M:%S GMT'))
### other way
#import dateutil.parser as dateparser
#dt = dateparser.parse(dstr)
#timestampnum = time.mktime(dt.timetuple())
#return datetime.datetime.fromtimestamp(int(timestampnum)).strftime('%Y-%m-%d %H:%M:%S')
#
#'Thu Mar 31 17:31:45 CST 2016' => 20160331173145
def write_current_issue_name(issue):
    wfile = open(".t_issue.log", "a")
    wfile.write(issue + "\n")
    wfile.close()

	
def convertdatastring(dstr):
    timestamp = time.mktime(time.strptime(dstr, '%a %b %d %H:%M:%S CST %Y'))
    result = datetime.datetime.fromtimestamp(int(timestamp)).strftime('%Y%m%d%H%M%S')
    return result

def issue(opener, issue):
	op = opener.open(_loginweb + "browse/" + issue + "?page=com.atlassian.jira.plugin.ext.subversion:subversion-commits-tabpanel")
	text = op.read()
	title = issue + re.compile('<title>\[.*\] (.*) - .*</title>').findall(text)[0]	
	#dstr_t = re.compile('<td bgcolor="#ffffff" width="10%" valign="top" rowspan="3">([a-zA-Z].*[0-9]{4})</td>').findall(text)	
	dstr_t = re.compile('([a-zA-Z]{3}\s[a-zA-Z]{3}\s[0-9]{2}\s[0-9]{2}:[0-9]{2}:[0-9]{2}\sCST\s[0-9]{4})').findall(text)
	dstr = "--"
	if len(dstr_t) > 0:
		dstr = convertdatastring(dstr_t[len(dstr_t) - 1])
	#authors = re.compile('<td bgcolor="#ffffff" width="10%" valign="top" rowspan="3">([a-z]+)</td>').findall(text)
	authors = re.compile('(u003e[a-z]+)').findall(text)
	result = {}
	for v in authors:
		s = v[5:]
		if not result.has_key(s):
			result[s] = s
			
	#get tester if exist
	tester = "--"
	#tester_t = re.compile('<dt title="测试相关负责人">[\s\S]*">(.*)</span></span>').findall(text)
	#write_current_issue_name("text " + tester_t + " ...")	
	#if len(tester_t) > 0:
		#tester = tester_t[0]	
	return title, dstr, result, tester

#***** 1--策划, 2-客户端，3-服务器 *****
def find_author(name):
	for v in _desgion:
		if name == v:
			return 1,_desgion[v]
	for v in _client:
		if name == v:
			return 2,_client[v]
	for v in _server:
		if name == v:
			return 3,_server[v]
	for v in _qa:
		if name == v:
			return 4,_qa[v]
	return 1, name
	
def write_excel_header(sheet):
    sheet.write(0, 0,"单号名称")
    sheet.write(0, 1, "策划")
    sheet.write(0, 2, "客户端")
    sheet.write(0, 3, "服务器")
    #sheet.write(0, 4, "测试")
    sheet.write(0, 4, "最后提交时间")

def write_excel_row(sheet, raw, col, text):
	sheet.write(raw, col, text)

def get_xls_name():
    if not os.path.exists("result.xls"):
        return "result.xls"
    i = 1
    func_name = lambda arg: "result" + '%02d'%arg + '.xls'
    while os.path.exists(func_name(i)):
        i += 1
    return func_name(i)



def main():
    cj = cookielib.CookieJar()
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))
    opener.addheaders=[('User-agent','Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)')]
    data = urllib.urlencode(_userinfo)
    opener.open(_loginweb + "login.jsp", data)

    xls_file = xlwt.Workbook(encoding='utf-8')
    sheet= xls_file.add_sheet('Sheet1',cell_overwrite_ok=True)
    write_excel_header(sheet)
    current_line = 1

    rfile = open("input.txt", 'r')
    line = rfile.readline()
    while line:
        line = line.replace(" ", "")
        line = line.replace("\t", "")
        line = line.replace("\r", "")
        line = line.replace("\n", "")       
        write_current_issue_name("Processing " + line + " ...")
        line_map = {}
        title, dstr, result, tester = issue(opener, line)
        #write_excel_row(sheet, current_line, 4, tester)
        write_excel_row(sheet, current_line, 4, dstr)
        write_excel_row(sheet, current_line, 0, title)		
        for v in result:
            author_type, author_name = find_author(v)
            if line_map.has_key(author_type):
                line_map[author_type] = line_map[author_type] + "、" + author_name
            else:
                line_map[author_type] = author_name
        
        for v in line_map:
            write_excel_row(sheet, current_line, v, line_map[v])
        for i in range(1,4):
            if not line_map.has_key(i):
                write_excel_row(sheet, current_line, i, "--")

        current_line = current_line + 1
        line = rfile.readline()
    write_current_issue_name("Writing to excel, please wait... ╰(￣▽￣)╮")
    xlsfilename = get_xls_name()
    xls_file.save(xlsfilename)
    rfile.close()
    write_current_issue_name("Well Done! (^o^)")
    os.startfile(xlsfilename)
    print("everything is done!")
