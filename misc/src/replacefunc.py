#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
def targetFunc():  # この関数の中の関数を置換する。
    print("ターゲットの関数")
def newFunc(arg):  # 置換後の新しい関数。
    print("置換後の関数で{}を出力。".format(arg))
    
from functools import wraps
import inspect   
def replaceFunc(oldfunc, newfunc):  # targetfunc内のoldfuncをnewfuncに置換する。
    def decorate(func):
        module_name = inspect.getmodulename(__file__)  # newfuncのあるモジュール名を取得。
        srclines = inspect.getsource(func).splitlines()  # tagergetfuncのソースを1行ずつ要素にしたリストを取得。
        srclines.insert(1,"""
    from {0} import {2}
    {1} = {2}
        """.format(module_name, oldfunc.__name__, newfunc.__name__))  # targetfuncのソースを改変する。
        src = '\n'.join(srclines)  # targetfuncのソースを作成。
        temp = dict()  # exec()の仮想モジュールの名前空間を受けとる辞書。
        exec(compile(src,'virtual_module','exec'), temp, temp)  # srcをコンパイルしてtempに取得。
        @wraps(func)
        def wrapper(*args, **kwargs):
            return temp[func.__name__](*args, **kwargs)  # 作り替えたtargetfuncを返す。
        return wrapper
    return decorate

if __name__ == "__main__":
    targetFunc = replaceFunc(print, newFunc)(targetFunc)
    targetFunc()