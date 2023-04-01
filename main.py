# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from decimal import Decimal


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
def is_1(gou_shou_dian_all_right):
    gou_shou_dian_all_right = False
    pass


def format_string(input_string):
    return input_string.lower().replace(' ', '_')


# 这个函数接受一个字母作为参数，并将字母转换为其在字母表中的顺序。如果给定的字母不是小写字母，函数会返回 `None`。
def get_letter_order(letter):
    alphabet = set("abcdefghijklmnopqrstuvwxyz")
    if letter in alphabet:
        return ord(letter) - ord("a") + 1
    else:
        return None


# pyinstaller --name "二级市场购售电力销售CheckV1.0" --onefile --hidden-import=openpyxl --hidden-import=wx --add-data="./report/check.py;./report/" ./MainWindow/MainWindow.py
if __name__ == '__main__':
    excel_list_file = [1, 2, "3"]
    print(format_string("Full internet access"))
    print(get_letter_order("t"))
    print("3" in excel_list_file)

    # print_hi('PyCharm')
    # name ="（4）从省级以下电网企业购电（含趸售企业）"
    # print(name.count("从省级以下电网企业购电"))
    # print(name.find("从省级以下电网企业购电"))
    # print(name in "从省级以下电网企业购电")

    # 这是范围
    # for row in range(1, 11, 4):
    #     print(row, end=",")
    #
    # print("")

    # list = [1, 2, 3, 4]
    # for row in list:
    #     if str(row) == "2":
    #         continue
    #     print(row)
    #     if str(row) == "3":
    #         break

    pass

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
