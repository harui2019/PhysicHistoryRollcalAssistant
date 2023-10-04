import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import json
import argparse
import re
from typing import Optional, Literal
from colorama import Fore, Style


def fileLocFind() -> dict[Literal['attend', 'test', 'hw'], str]:
    print("-"*40)
    print("| Script location:", __file__)
    print("| Current working directory:",
          Fore.YELLOW + os.getcwd() + Style.RESET_ALL)
    print("| Change working directory to script location")
    os.chdir(os.path.dirname(__file__))
    print("| Current working directory:",
          Fore.YELLOW + os.getcwd() + Style.RESET_ALL)
    print("-"*40)
    print("| Search 'ta_support_files.json'")
    if os.path.isfile('./ta_support_files.json'):
        fileLocations = json.load(open('./ta_support_files.json'))
    else:
        print(Fore.RED + "| 'ta_support_files.json' not found" + Style.RESET_ALL)
        fileLocations = {}
    print("-"*40)
    print("| Search excel files in 'ta_support_files.json")
    for k, v in fileLocations.items():
        if os.path.isfile(v):
            print("| Found file: ", k, v)
        else:
            print("| File not found: ", k, v)
            print(Fore.RED + "| Please check 'ta_support_files.json' and check the file location of excel files." + Style.RESET_ALL)
            # print("| exit")
            # exit()

    return fileLocations


def mode_and_target(
    mode: str,
    title: str,
    fileLocations: dict[Literal['attend', 'test', 'hw'], str],
    reserved_col: list[str] = ['序號', '系級', '學號', '姓名'],
) -> tuple[pd.DataFrame, Literal['attend', 'test', 'hw'], str, str]:

    path: str = ''
    while all([mode != "hw", mode != "attend", mode != "test"]):
        mode = input(
            Fore.BLUE + "| 輸入 'attend' 開始點名 / 'hw' 開始登記作業: " + Style.RESET_ALL)
        print(Fore.RED + Style.BRIGHT +
              'ERROR! Plz check the mode!' + Style.RESET_ALL)
    if mode == "hw":
        path = fileLocations['hw']
    elif mode == "attend":
        path = fileLocations['attend']
    elif mode == "test":
        path = fileLocations['test']
        
    revised = pd.read_excel(path)

    while len(title) == 0:
        print('| Following are the column head of the excel file:')
        print('| {}'.format(revised.columns))
        if mode == "hw":
            titleraw = input(Fore.BLUE+"| hw number: "+Style.RESET_ALL)
        elif mode == "attend":
            titleraw = input(Fore.BLUE+'| Date: '+Style.RESET_ALL)
        else:
            titleraw = input(Fore.BLUE+'| Test: '+Style.RESET_ALL)
            
        if titleraw in reserved_col:
            print(
                Fore.RED + Style.BRIGHT +
                '| It is reserved column head. Please choose other one!' + Style.RESET_ALL)
        elif titleraw in revised.columns:
            tmp = input(
                Fore.YELLOW +
                '| It is already in the column head. Do you want to edit or use other title?\n'+
                "| 'y' for edit, 'n' for using other title to create a new one." +
                Style.RESET_ALL+Fore.BLUE +
                "\n>>> "+Style.RESET_ALL
            )
            if tmp == 'y':
                title = titleraw    
        else:
            title = titleraw
            

    print("| mode:", mode)
    print("| target:", title)
    print("| path:", path)
    print("| Begin to write excel file")

    return revised, mode, title, path


class MyProgramArgs(argparse.Namespace):
    mode: Optional[Literal['attend', 'test', 'hw']]
    title: Optional[str]
    check: bool


if __name__ == '__main__':
    fileLocations = fileLocFind()
    # if len(fileLocations) == 0:
    #     exit()

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-c", "--check",
        help="show config",
        action="store_true",
    )
    parser.add_argument(
        "-m", "--mode",
        help="choose mode: 'attend' or 'hw'",
        type=str,
        default='',
    )
    parser.add_argument(
        "-t", "--title",
        help="write column title",
        type=str,
        default='',
    )

    args: MyProgramArgs = parser.parse_args()
    # print(args)
    # print(args.mode)

    if args.check:
        print(fileLocations)
        exit()

    reserved_col = ['序號', '系級', '學號', '姓名']
    revised, mode, title, path = mode_and_target(
        mode=args.mode,
        title=args.title,
        fileLocations=fileLocations,
    )

    if title in revised.columns:
        print(Fore.YELLOW + "| Edit column '{}'.".format(title) + Style.RESET_ALL)
    else:
        today_col = pd.DataFrame(
            np.zeros(np.array(revised).shape[0]), columns=[title])
        today_col[title] = today_col[title].astype(str)
        revised = pd.concat([revised, today_col], axis=1)
    revised['學號'] = revised['學號'].astype(str)
    # 政大學號系所代碼部分存在非數字的情況，但非常罕見

    p = re.compile('[a-zA-Z0-9_]+')
    running = True
    while running:
        students = []
        raw = input(
            Fore.YELLOW +
            "| Input the student id, you can input multiple at once and divid them by ','.\n" +
            "| Input '0' or 'end' to stop." +
            Style.RESET_ALL+Fore.BLUE +
            "\n>>> "+Style.RESET_ALL
        )
        serials = raw.split(',')
        serials = [s.strip() for s in serials]

        exit_sign = False
        if len(serials) == 1:
            if len(serials[0]) == 0:
                continue
            exit_sign = serials[0] == '0' or serials[0] == 'end'

        if not exit_sign:
            for i, s in enumerate(serials):
                reresult = p.match(s)
                if reresult is not None:
                    res = reresult.group()
                else:
                    res = ''

                if len(res) == 9:
                    students.append(res)
                else:
                    print(
                        Fore.RED +
                        "|  - input {}: '{}' is detected as '{}' which is unrecongnized.".format(i, s, res) +
                        Style.RESET_ALL
                    )
        else:
            print('| exit')
            running = False

        if len(students) > 0:
            filtered = revised[revised['學號'].isin(students)]
            if len(filtered) > 0:
                print("| Search:\n", filtered)
                for student in students:
                    check = input(
                        Fore.YELLOW +
                        '| Does number {} match: \n{}'.format(
                            student, revised.loc[(revised['學號'] == student)]
                        ) +
                        Style.RESET_ALL+Fore.BLUE +
                        "\n| ENTER to yes, 'n' for no, 'l' for day-off" +
                        "\n>>> "+Style.RESET_ALL
                    )
                    if len(check) == 0:
                        revised.loc[(revised['學號'] == student), title] = "1"
                    elif check == 'l':
                        revised.loc[(revised['學號'] == student), title] = "假"
                    else:
                        print(Fore.RED + Style.BRIGHT +
                              '| No assign for {}'.format(student) + Style.RESET_ALL)
            else:
                print(
                    Fore.RED + Style.BRIGHT +
                    '| No found any student for following numbers:\n {}'.format(students) +
                    Style.RESET_ALL)
        else:
            ...

    revised.to_excel(path, index=False)
    print(Fore.BLUE + Style.BRIGHT + '| The data has been saved !' + Style.RESET_ALL)
    exit()
