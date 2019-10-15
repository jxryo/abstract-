# coding=utf-8
import urllib2
import urllib
import math
import os
import xlwt
import datetime


# ����ҳ
def url_open(url):
    user_agent = 'User-Agent', 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
    header = {'User-Agent': user_agent}
    values = {'username': '', 'password': ''}
    data = urllib.urlencode(values)
    try:
        req = urllib2.Request(url, data, header)
    except:
        print
        "URL_ERROR"
        urlx = 'http://tennisabstract.com/reports/wtaRankings.html'
        reqx = urllib2.Request(urlx, data, header)
        repx = urllib2.urlopen(reqx)
        return repx.read()
    rep = urllib2.urlopen(req)
    return rep.read()


# ��ȡHTML
def get_html(url):
    html = url_open(url).decode('iso-8859-1')
    return html


# ��ȡԭʼ�����е������ַ���
def cut_name(url):
    a = url.find('p=')
    b = url.find('">')
    return url[a + 2:b]


# ��ʽ��ʱ��
def tran_time(str):
    time_data = {
        '01': 'Jan',
        '02': 'Feb',
        '03': 'Mar',
        '04': 'Apr',
        '05': 'May',
        '06': 'Jun',
        '07': 'Jul',
        '08': 'Aug',
        '09': 'Sep',
        '10': 'Oct',
        '11': 'Nov',
        '12': 'Dec'
    }
    try:
        day = str[-2:]
        month = time_data[str[-4:-2]]
        year = str[:4]
        new_str = day + '-' + month + '-' + year
    except:
        new_str = ''
    return new_str


# ��ȡmessage
def f_opp(item, a):
    opp = 11
    oseed = 13
    oentry = 14
    ocountry = 18
    score = 9
    wl = 4
    seed = 6
    entry = 7

    pentry = ''
    if item[seed] != '':
        pentry = '(' + item[seed] + ')'
    elif item[entry] != '':
        pentry = '(' + item[entry] + ')'
    pname = names[a]

    oppentry = ''
    if item[oseed] != '':
        oppentry = '(' + item[oseed] + ')'
    elif item[oentry] != '':
        oppentry = '(' + item[oentry] + ')'
    oppcc = '[' + item[ocountry] + ']'
    detlink = ''
    if item[score] == '':
        detlink = 'vs'
    else:
        detlink = 'd.'

    opplink = item[11]

    opp_all = ''
    if item[wl] == 'W' or item[wl] == 'U':
        opp_all = pentry + pname + ' ' + detlink + ' ' + oppentry + opplink + oppcc
    else:
        opp_all = oppentry + opplink + oppcc + ' ' + detlink + ' ' + pentry + pname
    return opp_all


# ��ʽ������
def alignRound(num, dec, perc):
    if perc == 1:
        num *= 100
    intrate = round(num * math.pow(10, dec)) / math.pow(10, dec)
    extra = intrate + math.pow(10, -1 * (dec + 1))
    strx = str(extra)
    indexdot = strx.find('.')
    done = ''
    if perc == 1:
        done = strx[0:indexdot + dec + 1] + '%'
    else:
        done = strx[0:indexdot + dec + 1]
    if done[0] == 'N' or done[0] == '%':
        return '-'
    elif dec == 0:
        return done[0:-2] + done[-1]
    else:
        return done


# ��ȡ����ʱ��
def show_time(times):
    try:
        times = eval(times[20])
        hours = str(int(times / 60))
        minutes = int(times % 60)
        if minutes < 10:
            minutes = '0' + str(minutes)
        else:
            minutes = str(minutes)
        showtime = hours + ':' + minutes
        return showtime
    except:
        return ''


def scores_change(scores):
    for i in range(scores.__len__()):
        if scores[i] == '-':
            asc = scores[i - 1]
            bsc = scores[i + 1]
            fsc = scores[:i - 1]
            lsc = scores[i + 2:]
            scores = fsc + bsc + scores[i] + asc + lsc
    return scores


def anl_num(strs):
    if strs[0] == '0':
        return strs[1]
    else:
        return strs


# matchhead ԭʼ���ݶ�Ӧ�б��ֵ�
matchhead = ["date", "tourn", "surf", "level", "wl", "rank", "seed", "entry", "round",
             "score", "max", "opp", "orank", "oseed", "oentry", "ohand", "obday",
             "oht", "ocountry", "oactive", "time", "aces", "dfs", "pts", "firsts", "fwon",
             "swon", 'games', "saved", "chances", "oaces", "odfs", "opts", "ofirsts",
             "ofwon", "oswon", 'ogames', "osaved", "ochances", "obackhand", "chartlink",
             "pslink", "whserver", "matchid"]
# #test port
# count =0
# for i in matchhead:
#     print '<' +str(count)+'>'
#     print i
#     count+=1
# get time
# ��ȡ����ʱ��
trs = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
Dates = trs[0:4] + '.' + trs[5:7] + '.' + trs[8:10]
# url
urls = [['http://tennisabstract.com/reports/wtaRankings.html', 'WTA'],
        ['http://tennisabstract.com/reports/atpRankings.html', 'ATP']]
# urlst = [['http://tennisabstract.com/reports/atpRankings.html', 'ATP'],
#          ['http://tennisabstract.com/reports/wtaRankings.html', 'WTA']]
root_name = "ATP WTA+" + Dates
# �������ļ���
try:
    os.mkdir(root_name)
except:
    print
    "dir has been create"
excels = []
for i in range(20):
    excels.append(xlwt.Workbook())
exc_index = 0
# data get ��ȡԭʼ����
W_A = True
for U in urls:
    if U[1] == 'ATP':
        W_A = False
    Rank_low = 1
    Rank_high = 30
    try:
        data = get_html(U[0])
    except:
        print
        '404 error'
        break
    arr = []
    c = data.find('style="vertical-align:top"')
    data = data[c:]
    for i in range(800):
        a = data.find('http:')
        b = data.find('&nbsp;')
        arr.append(data[a:b])
        data = data[b + 5:]
    names = []
    urls = []
    j = 1
    # ��ÿ�����ݽ��д���
    while names.__len__() <= 300:
        c = arr[j].find('">')
        urls.append(arr[j][:c])
        name = cut_name(arr[j])
        if name != '':
            names.append(cut_name(arr[j]))
        j += 1

    t = 0
    dir_name = root_name + '/' + U[1] + '+' + Dates
    if U[1] == 'ATP':
        exc_index += 1
    try:
        os.mkdir(dir_name)
    except:
        print
        "dir has been created"
    error_names = []
    for i in range(300):
        # try:
        print
        names[t] + ' start:'
        py_url = 'http://www.minorleaguesplits.com/tennisabstract/cgi-bin/jsmatches/' + names[t] + '.js'
        try:
            py_data = get_html(py_url)
        except:
            print
            "404 error"
            i += 1
            continue
        c = py_data.find('chmx =') + 6
        py_data = py_data[c:-2]
        arr_pydata = eval(py_data)

        # more 2 years
        lastyear = eval(arr_pydata[1][0][:4]) - 2

        # print lastyear
        # print type(lastyear)
        lpy_url = 'http://www.minorleaguesplits.com/tennisabstract/cgi-bin/jsmatches/' + names[t] + 'Career.js'
        try:
            lpy_data = get_html(lpy_url)
            lc = lpy_data.find('mx') + 5
            lpy_data = lpy_data[lc:-2]
            arr_lpydata = eval(lpy_data)
            index = 0
            for ix in arr_lpydata:
                index += 1
                year = eval(ix[0][:4])
                if year >= lastyear:
                    break
            arr_lpydata = arr_lpydata[index:]
            arr_pydata = arr_lpydata + arr_pydata
        except Exception as e:
            print
            e

        # ����2019-1-19 ��þ������е�����ʱ��,������ӱ�����ҳ��������
        if W_A == False:
            try:
                old_latest_time = str(arr_pydata[-1][0])
                new_url = 'http://www.tennisabstract.com/cgi-bin/player.cgi?p=' + names[t]
                new_data = get_html(new_url)
                nlc = new_data.find('matchmx') + 9
                new_data = new_data[nlc:]
                nld = new_data.find('fourspaces') - 8
                new_data = new_data[:nld]
                arr_new_data = eval(new_data)
                new_index = 0
                for nd in arr_new_data:
                    new_index += 1
                    # print new_index
                    cp_string = str(nd[0])
                    now_time = eval(nd[0][:4])
                    now_month = eval(anl_num(nd[0][4:6]))
                    now_day = eval(anl_num(nd[0][6:]))

                    # ��bug��ʱ���
                    if now_time <= 2018:
                        if now_month <= 11:
                            if now_day <= 12:
                                break
                            elif now_month < 11:
                                break

                    if cp_string == old_latest_time:
                        break
                # for ixx in arr_new_data:
                #     print ixx
                # print new_index
                arr_new_data = arr_new_data[:new_index - 1]
                arr_new_data = list(reversed(arr_new_data))
                arr_pydata = arr_pydata + arr_new_data
            except Exception as e:
                print
                e

        res_data = [
            ['Date', 'Tournament', 'surface', 'Rd', 'Rk', 'vRk', '', 'score', 'DR', 'A%', 'DF%', '1stln', '1st%',
             '2nd%', 'BPsaved', 'TPW', 'RPW', 'vA%', 'v1st%', 'v2nd%', 'BPConv', 'Time', '', 'dataorder']]
        lenth = arr_pydata.__len__()
        if i % 30 == 0 and i != 0:
            Rank_high += 30
            Rank_low += 30
            exc_index += 1
        table = excels[exc_index].add_sheet(str(i + 1) + ' ' + names[t], True)
        for i in range(1, lenth + 1):
            dfs = eval(str(arr_pydata[i - 1][22]) + '.0')
            aces = eval(str(arr_pydata[i - 1][21]) + '.0')
            if str(arr_pydata[i - 1][23]).__len__() < 4:
                pts = eval(str(arr_pydata[i - 1][23]) + '.0')
            else:
                pts = 0.0
            firsts = eval(str(arr_pydata[i - 1][24]) + '.0')
            fwon = eval(str(arr_pydata[i - 1][25]) + '.0')
            if str(arr_pydata[i - 1][26]).__len__() < 4:
                swon = eval(str(arr_pydata[i - 1][26]) + '.0')
                saved = eval(str(arr_pydata[i - 1][28]) + '.0')
                chances = eval(str(arr_pydata[i - 1][29]) + '.0')
                oaces = eval(str(arr_pydata[i - 1][30]) + '.0')
                odfs = eval(str(arr_pydata[i - 1][31]) + '.0')
                opts = eval(str(arr_pydata[i - 1][32]) + '.0')
                ofirsts = eval(str(arr_pydata[i - 1][33]) + '.0')
                ofwon = eval(str(arr_pydata[i - 1][34]) + '.0')
                if str(arr_pydata[i - 1][35]).__len__() < 3:
                    oswon = eval(str(arr_pydata[i - 1][35]) + '.0')
                else:
                    oswon = 0.0
                osaved = eval(str(arr_pydata[i - 1][37]) + '.0')
                ochances = eval(str(arr_pydata[i - 1][38]) + '.0')
            else:
                swon = 0.0
            Date = tran_time(arr_pydata[i - 1][0])  # Date
            Tournament = arr_pydata[i - 1][1]  # Tournament
            surface = arr_pydata[i - 1][2]  # surface
            Rd = arr_pydata[i - 1][8]  # Rd
            Rk = arr_pydata[i - 1][5]  # Rk
            vRk = arr_pydata[i - 1][12]  # vRk
            message = f_opp(arr_pydata[i - 1], t)  # message

            if opts != 0.0:
                rpw = 1 - (int(ofwon) + int(oswon)) / opts
            else:
                rpw = '-'
            if pts != 0.0:
                spl = 1 - (int(fwon) + int(swon)) / pts
            else:
                spl = '-'

            if spl != 0.0 and rpw != '-' and spl != '-':
                num = rpw / spl
                dec = 2
                DR = alignRound(num, dec, 0)  # DR
            else:
                DR = '-'

            score = arr_pydata[i - 1][9]  # score
            try:
                pass
                # if score!='W/O' and float (DR) <=1.00:
                #     score = scores_change(score)
                #     if float(DR)>=0.97:
                #         score = scores_change(score)
                # elif float(DR) <=1.03:
                #     score = scores_change(score)
            except:
                pass
            if score == 'W/O':
                all_arr = [Date, Tournament, surface, Rd, Rk, vRk, message, score, '', '', '', '', '', '', '', '',
                           '',
                           '', '', '', '', '', '']

            else:
                if pts != 0.0:

                    A_p = alignRound(aces / pts, 1, 1)  # A%
                else:
                    A_p = '-'
                if pts != 0.0:
                    DF_p = alignRound(dfs / pts, 1, 1)  # DF%
                else:
                    DF_p = '-'
                if pts != 0.0:
                    stln = alignRound(firsts / pts, 1, 1)  # 1stln
                else:
                    print
                    Date + ':'
                    print
                    'errorFirst:' + str(firsts)
                    print
                    'errorPts:' + str(pts)
                    print
                    arr_pydata[i - 1][23]
                    stln = '-'
                if firsts != 0.0:
                    st_p = alignRound(fwon / firsts, 1, 1)  # 1st%
                else:
                    st_p = '-'
                if pts - firsts != 0.0:

                    nd_p = alignRound(swon / (pts - firsts), 1, 1)  # 2nd%
                else:
                    nd_p = '-'
                if chances != 0.0:
                    if W_A:
                        BPsaved = alignRound(saved / chances, 1, 1) + '(' + str(int(saved)) + '/' + str(
                            int(chances)) + ')'  # BPsaved
                    else:
                        BPsaved = str(int(saved)) + '/' + str(int(chances))  # BPsaved
                else:
                    BPsaved = '-'
                # Time1 = show_time(arr_pydata[i - 1])  # Time
                pointswon = int(fwon) + int(swon) + (opts - ofwon - oswon)
                if int(pts) + int(opts) != 0.0:
                    TPW = alignRound((pointswon) / (int(pts) + int(opts)), 1, 1)  # TPW
                else:
                    TPW = '-'
                if opts != 0.0:
                    RPW = alignRound(1 - ((int(ofwon) + int(oswon)) / opts), 1, 1)  # RPW
                else:
                    RPW = '-'

                if opts != 0.0:
                    vA_p = alignRound((oaces / opts), 1, 1)  # vA%
                else:
                    vA_p = '-'
                if ofirsts != 0.0:
                    v1st_p = alignRound(1 - (ofwon / ofirsts), 1, 1)  # v1st%
                else:
                    v1st_p = '-'
                if opts - ofirsts != 0.0:
                    v2nd_p = alignRound(1 - (oswon / (opts - ofirsts)), 1, 1)  # v2nd%
                else:
                    v2nd_p = '-'
                if ochances != 0.0:
                    if W_A:
                        BPConv = alignRound(1 - (osaved / ochances), 1, 1) + '(' + str(
                            int(ochances - osaved)) + '/' + str(int(ochances)) + ')'  # BPConv
                    else:
                        BPConv = str(int(ochances - osaved)) + '/' + str(int(ochances))  # BPConv  # BPConv
                else:
                    BPConv = '-'
                Time2 = show_time(arr_pydata[i - 1])  # Time
                dataorder = str(arr_pydata[i - 1][-1])
                all_arr = [Date, Tournament, surface, Rd, Rk, vRk, message, score, DR, A_p, DF_p, stln, st_p, nd_p,
                           BPsaved, TPW,
                           RPW, vA_p, v1st_p, v2nd_p, BPConv, Time2, '', dataorder]
            if all_arr[0] == '22-Oct-2018':
                print
                arr_pydata[i - 1]
            if all_arr[11] != '-' and all_arr[11] != '0%' and all_arr[11] != '':
                # ������������
                res_data.append(all_arr)
            else:
                # print Date+':1stln��'+all_arr[11]
                pass
            # print str(i)+" is ok"
        count = 0
        table.write(count, 0, names[t])
        count += 1
        # ������д�����
        for i in res_data:
            for j in range(i.__len__()):
                table.write(count, j, i[j])
            count += 1
        try:
            # ���±��
            excels[exc_index].save(
                dir_name + '/' + U[1] + ' ' + str(Rank_low) + '-' + str(Rank_high) + '+' + Dates + '.xls')
        except:
            print
            'save data error'
        print
        names[t] + " is ok " + str(t)
        t += 1

    # except:
    #     print "UNKOWN ERROR"
    #     t += 1
    #     error_names.append(names[t])

    for i in error_names:
        print
        i
