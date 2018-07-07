import os
import re
from bs4 import BeautifulSoup
import requests
from urllib.request import urlopen
import os
import pandas as pd
import csv
from collections import OrderedDict
from openpyxl import load_workbook
from xlrd import open_workbook



def downloadStats(stats_file_path):
    if not os.path.exists(os.path.dirname(stats_file_path)):
        os.makedirs(os.path.dirname(stats_file_path))

    # points part
    points = 'https://www.cbssports.com/nhl/stats/playersort/nhl/year-2017-season-regularseason-category-points?print_rows=9999'

    html = urlopen(points)

    soup = BeautifulSoup(html, 'lxml')

    title = soup.find('td', {'colspan': re.compile('.*?')})
    title1 = title.get_text()

    cells = soup.find_all('td', {'align': re.compile('.*?')})
    rows = soup.find_all('tr', {'id': re.compile('.*?')})

    PLAYER = []
    POS = []
    TEAM = []
    GP = []
    G = []
    A = []
    PTS = []
    Positive_negative = []
    PIM = []
    SOG = []
    PCT = []
    WG = []
    TOI = []
    PPG = []
    PPA = []
    SHG = []
    SHA = []
    SG = []
    SA = []
    List = []

    n = 0

    for c in cells:
        c = c.get_text().strip()

        List.append(c)
        n = n + 1

        if n == 19:
            PLAYER.append(List[0])
            POS.append(List[1])
            TEAM.append(List[2])
            GP.append(List[3])
            G.append(List[4])
            A.append(List[5])
            PTS.append(List[6])
            Positive_negative.append(List[7])
            PIM.append(List[8])
            SOG.append(List[9])
            PCT.append(List[10])
            WG.append(List[11])
            TOI.append(List[12])
            PPG.append(List[13])
            PPA.append(List[14])
            SHG.append(List[15])
            SHA.append(List[16])
            SG.append(List[17])
            SA.append(List[18])

            n = 0
            List = []

    rank1 = list(range(1, len(PLAYER) + 1))

    cols = ['RANK', 'PLAYER', 'POS', 'TEAM', 'GP', 'G', 'A', 'PTS', '+/-', 'PIM', 'SOG', 'PCT', 'WG', 'TOI', 'PPG',
            'PPA', 'SHG', 'SHA', 'SG', 'SA']

    NHL_df = pd.DataFrame({
        'RANK': rank1,
        'PLAYER': PLAYER,
        'POS': POS,
        'TEAM': TEAM,
        'GP': GP,
        'G': G,
        'A': A,
        'PTS': PTS,
        '+/-': Positive_negative,
        'PIM': PIM,
        'SOG': SOG,
        'PCT': PCT,
        'WG': WG,
        'TOI': TOI,
        'PPG': PPG,
        'PPA': PPA,
        'SHG': SHG,
        'SHA': SHA,
        'SG': SG,
        'SA': SA,
    }, columns=cols)

    NHL_df['RANK'] = NHL_df['RANK'].astype(str)

    # penalty part
    penalty = 'https://www.cbssports.com/nhl/stats/playersort/nhl/year-2017-season-regularseason-category-penaltyminutes?print_rows=9999'

    html = urlopen(penalty)

    soup = BeautifulSoup(html, 'lxml')

    title = soup.find('td', {'colspan': re.compile('.*?')})
    title2 = title.get_text()

    cells = soup.find_all('td', {'align': re.compile('.*?')})
    rows = soup.find_all('tr', {'id': re.compile('.*?')})

    PLAYER = []
    POS = []
    TEAM = []
    GP = []
    PEN = []
    PIM = []
    MAJ = []
    MIN = []
    MIS = []
    PPG = []
    PIM_divide_G = []
    FM = []
    CHG = []
    BD = []
    INST = []
    PPG1 = []

    List = []

    n = 0

    for c in cells:
        c = c.get_text().strip()

        List.append(c)
        n = n + 1

        if n == 16:
            PLAYER.append(List[0])
            POS.append(List[1])
            TEAM.append(List[2])
            GP.append(List[3])
            PEN.append(List[4])
            PIM.append(List[5])
            MAJ.append(List[6])
            MIN.append(List[7])
            MIS.append(List[8])
            PPG.append(List[9])
            PIM_divide_G.append(List[10])
            FM.append(List[11])
            CHG.append(List[12])
            BD.append(List[13])
            INST.append(List[14])
            PPG1.append(List[15])

            n = 0
            List = []

    rank2 = list(range(1, len(PLAYER) + 1))

    cols2 = ['RANK', 'PLAYER', 'POS', 'TEAM', 'GP', 'PEN', 'PIM', 'MAJ', 'MIN', 'MIS', 'PPG', 'PIM/G', 'FM', 'CHG',
             'BD','INST', 'PPG1']

    NHL_df2 = pd.DataFrame({
        'RANK': rank2,
        'PLAYER': PLAYER,
        'POS': POS,
        'TEAM': TEAM,
        'GP': GP,
        'PEN': PEN,
        'PIM': PIM,
        'MAJ': MAJ,
        'MIN': MIN,
        'MIS': MIS,
        'PPG': PPG,
        'PIM/G': PIM_divide_G,
        'FM': FM,
        'CHG': CHG,
        'BD': BD,
        'INST': INST,
        'PPG1': PPG1,
    }, columns=cols2)

    NHL_df2['RANK'] = NHL_df2['RANK'].astype(str)
    # wins part
    wins = 'https://www.cbssports.com/nhl/stats/playersort/nhl/year-2017-season-regularseason-category-wins?print_rows=9999'

    html = urlopen(wins)

    soup = BeautifulSoup(html, 'lxml')

    title = soup.find('td', {'colspan': re.compile('.*?')})
    title3 = title.get_text()

    cells = soup.find_all('td', {'align': re.compile('.*?')})
    rows = soup.find_all('tr', {'id': re.compile('.*?')})

    PLAYER = []
    POS = []
    TEAM = []
    GP = []
    GS = []
    W = []
    L = []
    OT = []
    SA = []
    GA = []
    GAA = []
    SVPCT = []
    SO = []
    G = []
    A = []
    PIM = []
    TOI = []

    List = []

    n = 0

    for c in cells:
        c = c.get_text().strip()

        List.append(c)
        n = n + 1

        if n == 17:
            PLAYER.append(List[0])
            POS.append(List[1])
            TEAM.append(List[2])
            GP.append(List[3])
            GS.append(List[4])
            W.append(List[5])
            L.append(List[6])
            OT.append(List[7])
            SA.append(List[8])
            GA.append(List[9])
            GAA.append(List[10])
            SVPCT.append(List[11])
            SO.append(List[12])
            G.append(List[13])
            A.append(List[14])
            PIM.append(List[15])
            TOI.append(List[15])

            n = 0
            List = []

    rank3 = list((range(1, len(PLAYER) + 1)))

    cols3 = ['RANK', 'PLAYER', 'POS', 'TEAM', 'GP', 'GS', 'W', 'L', 'OT', 'SA', 'GA', 'GAA', 'SVPCT', 'SO', 'G',
             'A', 'PIM', 'TOI']

    NHL_df3 = pd.DataFrame({
        'RANK': rank3,
        'PLAYER': PLAYER,
        'POS': POS,
        'TEAM': TEAM,
        'GP': GP,
        'GS': GS,
        'W': W,
        'L': L,
        'OT': OT,
        'SA': SA,
        'GA': GA,
        'GAA': GAA,
        'SVPCT': SVPCT,
        'SO': SO,
        'G': G,
        'A': A,
        'PIM': PIM,
        'TOI': TOI,
    }, columns=cols3)

    NHL_df3['RANK']= NHL_df3['RANK'].astype(str)


    print(type(NHL_df3['RANK']))

#write
    write_path = os.path.abspath(stats_file_path)
    print(write_path)
    writer = pd.ExcelWriter(write_path, engine='xlsxwriter')

    NHL_df.to_excel(writer, sheet_name=title1, index = False)
    NHL_df2.to_excel(writer, sheet_name=title2, index = False)
    NHL_df3.to_excel(writer, sheet_name=title3, index = False)

    writer.save()



def print_m():
    print("=============== MENU =================")
    print("1. Display Points Leaders")
    print("2. Display Wins Leaders")
    print("3. Display Penalty Minutes Leaders")
    print("r. Refresh data")
    print("q. Quit")



def display_data(stats_file_path, sheet_name, top):
    f = None
    try:

        wb = open_workbook(stats_file_path)

        worksheet = wb.sheet_by_name(sheet_name)

        #datas = [t for t in worksheet.get_rows()]
        value = []
        print('**************',sheet_name.upper(),'*****************')
        for row_num in range(top + 1):
            row_value = worksheet.row_values(row_num)
            #row_value = row_value.remove('')
            if type(row_value) == float :
                row_value = int(row_value)
            print(*row_value, sep=', ')

    except Exception as e:
        print(e)

    if f is not None:
        f.close()


def main():
    stats_file_path = "downloads/nhl2/nhl_stats.xlsx"
    downloadStats(stats_file_path)
    while 1:
        print_m()
        option = str(input("Please select an option: "))
        top = int(input('Enter the maximum number of players to display (0 for all): '))
        if option == 'q':
            break
        elif option == 'r':
            main()
        elif option == '1':
            display_data(stats_file_path,'points',top)
        elif option == '2':
            display_data(stats_file_path,'wins',top)
        elif option == '3':
            display_data(stats_file_path,'penaltie minutes',top)
        else:
            print('invalid option', option)




if __name__ == '__main__':
    main()








