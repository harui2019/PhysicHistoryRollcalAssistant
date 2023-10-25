"""
============================================================================
TA support script for NCCU Course
    - The History of Physics and Human Civilization
============================================================================

## Usage

1. Put the script in the same folder with the excel files.
2. Run the script.
3. Choose the mode.
    - 'attend': attendance
    - 'hw': homework
    - 'group': group score
4. Input the student id or name.
5. Input the score, check the attendance and homework.
6. Input '0' or 'end' to stop.

"""
import re
import os
import sys
import json
import argparse
from typing import Optional, Literal, Iterable
import pandas as pd
from colorama import Fore, Style

RESERVED_COL = ['序號', '組別', '系級', '學號', '姓名']

# pylint: disable=line-too-long


def damerau_levenshtein_distance_py(
    seq1: Iterable[str],
    seq2: Iterable[str],
) -> int:
    """Calculate the Damerau-Levenshtein distance between sequences.

    This distance is the number of additions, deletions, substitutions,
    and transpositions needed to transform the first sequence into the
    second. Although generally used with strings, any sequences of
    comparable objects will work.

    Transpositions are exchanges of *consecutive* characters; all other
    operations are self-explanatory.

    This implementation is O(N*M) time and O(M) space, for N and M the
    lengths of the two sequences.

    >>> dameraulevenshtein('ba', 'abc')
    2
    >>> dameraulevenshtein('fee', 'deed')
    2

    It works with arbitrary sequences too:
    >>> dameraulevenshtein('abcd', ['b', 'a', 'c', 'd', 'e'])
    2

    This implementation is based on Michael Homer's implementation
    (https://web.archive.org/web/20150909134357/http://mwh.geek.nz:80/2009/04/26/python-damerau-levenshtein-distance/)
    and inspired by https://github.com/lanl/pyxDamerauLevenshtein, 
    a Cython implementation of same algorithm.

    For more powerful string comparison, including Levenshtein distance,
    We recommend using the https://github.com/maxbachmann/RapidFuzz,
    It's a library that wraps the C++ Levenshtein algorithm and other string processing functions.
    The most efficient Python implementation (using Cython) currently.

    Args:
        seq1 (Iterable): Sequence of items to be compared.
        seq2 (Iterable): Sequence of items to be compared.

    Returns:
        int: The distance between the two sequences.
    """
    # pylint: enable=line-too-long
    if seq1 is None:
        return len(seq2)
    if seq2 is None:
        return len(seq1)

    first_differing_index = 0
    while all([
        first_differing_index < len(seq1) - 1,
        first_differing_index < len(seq2) - 1,
        seq1[first_differing_index] == seq2[first_differing_index]
    ]):
        first_differing_index += 1

    seq1 = seq1[first_differing_index:]
    seq2 = seq2[first_differing_index:]

    two_ago, one_ago, this_row = [], [
    ], (list(range(1, len(seq2) + 1)) + [0])
    for x, _ in enumerate(seq1):
        two_ago, one_ago, this_row = one_ago, this_row, [0] * len(seq2) + [x + 1]
        for y, _ in enumerate(seq2):
            del_cost = one_ago[y] + 1
            add_cost = this_row[y - 1] + 1
            sub_cost = one_ago[y - 1] + (seq1[x] != seq2[y])
            # fun fact: isinstance(bool(...), int) == True
            this_row[y] = min(del_cost, add_cost, sub_cost)

            if all([
                x > 0, y > 0,
                seq1[x] == seq2[y - 1],
                seq1[x - 1] == seq2[y],
                seq1[x] != seq2[y]
            ]):
                this_row[y] = min(this_row[y], two_ago[y - 2] + 1)

    return this_row[len(seq2) - 1]


def file_location_find(
    supported_file: str = './ta_support_files.json'
) -> dict[Literal['attend', 'test', 'hw', 'group'], str]:
    """Find the file location of the supported files.

    Args:
        supported_file (str, optional): The file location of the supported files.

    Returns:
        dict[Literal['attend', 'test', 'hw', 'group'], str]: 
            The file location of the supported files.
    """

    print("-"*40)
    print("| Script location:", __file__)
    print(
        "| Current working directory:",
        Fore.YELLOW + os.getcwd() + Style.RESET_ALL)
    print("| Change working directory to script location")
    os.chdir(os.path.dirname(__file__))
    print(
        "| Current working directory:",
        Fore.YELLOW + os.getcwd() + Style.RESET_ALL)
    print("-"*40)
    print(f"| Search '{supported_file}'")
    if os.path.isfile(supported_file):
        with open(supported_file, encoding='utf-8') as f:
            print(f"| Found '{supported_file}'")
            file_location: dict[
                Literal['attend', 'test', 'hw', 'group'], str
            ] = json.load(f)
    else:
        print(Fore.RED + f"| '{supported_file}' not found" + Style.RESET_ALL)
        file_location = {}
    print("-"*40)
    print(f"| Search excel files in '{supported_file}")
    for k, v in file_location.items():
        if os.path.isfile(v):
            print("| Found file: ", k, v)
        else:
            print("| File not found: ", k, v)
            print(
                Fore.RED +
                "| Please check 'ta_support_files.json' " +
                "and check the file location of excel files." +
                Style.RESET_ALL
            )

    return file_location


def mode_and_target(
    mode: Literal['attend', 'test', 'hw', "group"],
    file_locations: dict[Literal['attend', 'test', 'hw', "group"], str],
    reserved_col: Optional[list[str]] = None,
) -> tuple[pd.DataFrame, Literal['attend', 'test', 'hw', "group"], str]:
    """Setup the mode and target file.

    Args:
        mode (str): Description
        file_locations (dict[Literal['attend', 'test', 'hw', "group"], str]): Description
        reserved_col (Optional[list[str]], optional): Description

    Returns:
        tuple[pd.DataFrame, Literal['attend', 'test', 'hw', "group"], str]: Description
    """

    if reserved_col is None:
        reserved_col = RESERVED_COL

    target_path: str = ''
    while all([mode != "hw", mode != "attend", mode != "test", mode != "group"]):
        mode = input(
            Fore.BLUE + "| 輸入 'attend' 開始點名 / 'hw' 開始登記作業 / 'group' 開始登記小組互評: " + Style.RESET_ALL)
        if all([mode != "hw", mode != "attend", mode != "test", mode != "group"]):
            print(
                Fore.RED + Style.BRIGHT +
                'ERROR! Plz check the mode!' + Style.RESET_ALL
            )
    if mode == "hw":
        target_path = file_locations['hw']
    elif mode == "attend":
        target_path = file_locations['attend']
    elif mode == "group":
        target_path = file_locations['group']
    else:
        target_path = file_locations['test']

    target = pd.read_excel(target_path)

    return target, mode, target_path


def check_col(
    target: pd.DataFrame,
    col: str,
    mode: Literal['attend', 'test', 'hw', "group"],
) -> tuple[pd.DataFrame, bool]:
    """Check the column in the target file.

    Args:
        target (pd.DataFrame): Dataframe of the target file.
        col (str): The column name to be checked.
        mode (Literal['attend', 'test', 'hw', "group"]): The mode of the target file.

    Returns:
        tuple[pd.DataFrame, bool]:
            The dataframe of the target file and the flag of whether the column is added.
    """

    is_add = False
    if not col in target.columns:
        is_add_new_col: Literal['y', 'n'] = ''
        while is_add_new_col not in ["y", "n"]:
            is_add_new_col = input(
                Fore.YELLOW + Style.BRIGHT +
                f"| '{col}' not found in excel file. Add it? [y/n]\n" +
                Style.RESET_ALL+Fore.BLUE+">>> " + Style.RESET_ALL
            )
            if is_add_new_col == 'y':
                target[col] = 0.0
                if mode == 'group':
                    target[col] = target[col].astype(float)
                else:
                    target[col] = target[col].astype(str)
                is_add = True
            elif is_add_new_col == 'n':
                print(
                    Fore.YELLOW + Style.BRIGHT +
                    f"| {col} would not be added." + Style.RESET_ALL
                )
                is_add = False
    target['學號'] = target['學號'].astype(str)
    target['組別'] = target['組別'].astype(str)

    p = re.compile('[a-zA-Z0-9]+')
    target['組別'] = target['組別'].apply(lambda x: p.findall(x)[0])

    return target, is_add


def handle_input(
    target: pd.DataFrame,
    mode: Literal['attend', 'test', 'hw', "group"],
    target_path: str,
    reserved_col: Optional[list[str]] = None,
    title_parse: re.Pattern = re.compile('[a-zA-Z0-9_-_._/]+'),
    input_parse: re.Pattern = re.compile('[a-zA-Z0-9_-_.]+'),
):
    """Handle the input from user.

    Args:
        target (pd.DataFrame): Dataframe of the target file.
        mode (Literal['attend', 'test', 'hw', "group"]): The mode of the target file.
        target_path (str): The path of the target file.
        reserved_col (Optional[list[str]], optional): The reserved column name of the target file.
        title_parse (re.Pattern, optional): The pattern of the title.
        input_parse (re.Pattern, optional): The pattern of the input.
    """

    if reserved_col is None:
        reserved_col = RESERVED_COL
    if not isinstance(title_parse, re.Pattern):
        raise TypeError(
            f"'title_parse' should be re.Pattern. not '{type(title_parse)}'.")
    if not isinstance(input_parse, re.Pattern):
        raise TypeError(
            f"'input_parse' should be re.Pattern. not '{type(input_parse)}'.")

    titles = []
    colunm_not_decide = True
    while colunm_not_decide:
        hint_for_group = ' (multiple title divided by \',\')' if mode == 'group' else ''
        raw_col = input(
            Fore.BLUE + "| 輸入欲修改的欄位: "+hint_for_group+"\n>>> "+Style.RESET_ALL)
        titles_raw = raw_col.split(',')
        titles_raw = [s.strip() for s in titles_raw]
        titles_raw = [title_parse.findall(s)[0] for s in titles_raw]

        if mode == 'group':
            is_confirmed = input(
                Fore.YELLOW + f"| Check the following column: \n{titles_raw}\n" +
                Style.RESET_ALL+Fore.BLUE+"| ENTER to yes, 'n' for no to reset" +
                "\n>>> "+Style.RESET_ALL)
            colunm_not_decide = is_confirmed == 'n'
            if not is_confirmed in ['', 'n']:
                print(
                    Fore.RED + Style.BRIGHT +
                    "| Input error. Reset the title." + Style.RESET_ALL
                )
                colunm_not_decide = True
            if colunm_not_decide:
                print(Fore.YELLOW + "| Title reset." + Style.RESET_ALL)
        else:
            if len(titles_raw) == 1:
                colunm_not_decide = False
            else:
                print(
                    Fore.RED + Style.BRIGHT +
                    "| Only mode 'group' allow multiple title." + Style.RESET_ALL
                )

        for i, t in enumerate(titles_raw):
            target, is_add = check_col(target, t, mode)
            if is_add:
                titles.append(title_parse)

        if len(titles) == 0:
            colunm_not_decide = True
            print(
                Fore.RED + Style.BRIGHT +
                "| No title selected. Reset the title." + Style.RESET_ALL
            )

    running = True
    while running:

        student = input(
            Fore.YELLOW +
            "| Input the student id or name, .\n" +
            "| Input '0' or 'end' to stop." +
            Style.RESET_ALL+Fore.BLUE +
            "\n>>> "+Style.RESET_ALL
        )
        student_id = ''

        exit_sign = student in ['0', 'end']
        if exit_sign:
            print('| exit')
            running = False
            continue
        if len(student) == 0:
            continue

        filtered_id = target[target['學號'].str.contains(student)]
        filtered_name = target[target['姓名'].str.contains(student)]
        if mode == 'group':
            chech_hint = (
                f"\n| Enter score for '{titles}' or '-1' for be-scored, divided by ','" +
                "\n| 'n' for no assign. "
            )
        elif mode in ('hw', 'attend'):
            chech_hint = "\n| ENTER to yes, 'n' for no, 'l' for day-off"
        else:
            chech_hint = ''

        if len(filtered_id) > 0:
            check = input(
                Fore.YELLOW +
                f"| Does number {student} match: \n{filtered_id}" +
                Style.RESET_ALL+Fore.BLUE+chech_hint +
                "\n>>> "+Style.RESET_ALL
            )
            student_id = filtered_id["學號"].iloc[0]
        elif len(filtered_name) > 0:
            check = input(
                Fore.YELLOW +
                f"| Does name '{student}' match: \n{filtered_name}" +
                Style.RESET_ALL+Fore.BLUE +
                Style.RESET_ALL+Fore.BLUE+chech_hint +
                "\n>>> "+Style.RESET_ALL
            )
            student_id = filtered_name["學號"].iloc[0]
        else:
            print(
                Fore.RED + Style.BRIGHT +
                f'| No found any student for following numbers:\n    {student}' +
                Style.RESET_ALL)
            similar_id = target['學號'].astype(str).apply(
                lambda x: damerau_levenshtein_distance_py(student, x) <= 3
            )
            similar_name = target['姓名'].astype(str).apply(
                lambda x: damerau_levenshtein_distance_py(student, x) <= 1
            )
            if len(similar_id) > 0:
                print("| Similar id:\n", target[similar_id])
            if len(similar_name) > 0:
                print("| Similar name:\n", target[similar_name])
            continue

        if check == 'n':
            continue

        if mode == 'group':
            scores = check.split(',')
            scores = [s.strip() for s in scores]
            scores = [input_parse.findall(s)[0] for s in scores]
            check_scores = [
                (k, k.replace(".", "").isnumeric() or k == '-1')
                for k in scores]

            if len(check_scores) != len(scores):
                print(
                    Fore.RED + Style.BRIGHT +
                    f'| Some score is not numeric: {check_scores}' +
                    Style.RESET_ALL
                )
                continue
            if len(scores) != len(titles):
                print(
                    Fore.RED + Style.BRIGHT +
                    f'| The number of score is not match: {check_scores}' +
                    Style.RESET_ALL
                )
                continue

            if all(k.replace(".", "").isnumeric() or k == '-1' for k in scores):
                for i, s in enumerate(scores):
                    target.loc[(target['學號'] == student_id),
                               titles[i]] = float(s)
                print("| Score added.")
                print(target[(target['學號'] == student_id)][
                    list(reserved_col)+titles])
                continue

            print(
                Fore.RED + Style.BRIGHT +
                "| Not all input are numerical. Score not added:\n    " +
                f"{scores}"+Style.RESET_ALL)

        else:
            if len(check) == 0:
                target.loc[(target['學號'] == student_id), titles[0]] = "1"
            elif check == 'l':
                target.loc[(target['學號'] == student), titles[0]] = "假"
            else:
                print(Fore.RED+Style.BRIGHT +
                      f'| No assign for {student}'+Style.RESET_ALL)

        target.to_excel(target_path, index=False)


class MyProgramArgs(argparse.Namespace):
    """args
    """
    mode: Optional[Literal['attend', 'test', 'hw', "group"]]
    title: Optional[str]
    check: bool


if __name__ == '__main__':
    fileLocations = file_location_find()
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
        help="choose mode: 'attend', 'hw', 'group'",
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

    if args.check:
        print(fileLocations)
        sys.exit()

    revised, mode_selected, path = mode_and_target(
        mode=args.mode,
        file_locations=fileLocations,
        reserved_col=RESERVED_COL,
    )

    handle_input(
        target=revised,
        mode=mode_selected,
        target_path=path,
        reserved_col=RESERVED_COL,
    )

    revised.to_excel(path, index=False)
    print(Fore.BLUE + Style.BRIGHT + "| File exported." + Style.RESET_ALL)
