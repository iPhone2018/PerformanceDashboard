# -*- coding: utf-8 -*-
"""数据目录的唯一权威定义。

看板（app.py）和采集（timer_get_data.py）都从这里取，**不各写一份**。
"采集写到一个地方、看板读另一个地方"是最难查的一类问题 —— 两边都不报错，
页面就是空的，日志里采集还全是"已保存"。只留一处定义，这类问题就不可能发生。

取值优先级：
    1. 环境变量 DASHBOARD_DATA_DIR（要换位置就设它，两个进程都要设成同一个）
    2. Windows：固定路径 C:\\Users\\Administrator\\Desktop\\data
       打包成 exe 后 __file__ 指向临时解包目录 _MEIxxxxxx，不能用它推路径
    3. 其他平台：本文件所在目录下的 data/（本地开发自测用）

兼容 Python 3.8+，只用标准库 —— 采集端没装 flask/pandas 时也要能导入。
"""

import os

DEFAULT_DATA_ROOT = r"C:\Users\Administrator\Desktop\data"

ENV_VAR = "DASHBOARD_DATA_DIR"


def data_root() -> str:
    """当前进程该用哪个数据目录。"""
    override = os.environ.get(ENV_VAR)
    if override:
        return override
    if os.name == "nt":
        return DEFAULT_DATA_ROOT
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


def describe() -> str:
    """一句话说明这个路径是怎么定下来的，启动时打出来备查。"""
    if os.environ.get(ENV_VAR):
        return f"{data_root()}（来自环境变量 {ENV_VAR}）"
    if os.name == "nt":
        return f"{data_root()}（Windows 固定路径）"
    return f"{data_root()}（本文件所在目录）"


def subdirs() -> tuple:
    """(日报目录, 时报目录, 日志目录)。"""
    root = data_root()
    return (os.path.join(root, "daily"),
            os.path.join(root, "hourly"),
            os.path.join(root, "logs"))
