#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SSH连接脚本

用法:
    python ssh_connect.py [命令]

示例:
    python ssh_connect.py                              # 交互式连接
    python ssh_connect.py "ls -la /root"              # 执行单个命令
    python ssh_connect.py "uptime; free -h"           # 执行多个命令
"""

import sys
import os
import subprocess

HOST = '192.168.2.2'
PORT = 22
USER = 'root'
PASSWORD = 'admincw'

def check_ssh_client():
    """检查系统是否有SSH客户端"""
    if shutil.which('ssh'):
        return 'ssh'
    if shutil.which('plink'):
        return 'plink'
    return None

def connect_ssh(command=None):
    """SSH连接函数"""
    import shutil

    ssh_client = check_ssh_client()

    if not ssh_client:
        print('❌ 未找到SSH客户端')
        print('请安装 OpenSSH (ssh) 或 PuTTY (plink)')
        return False

    if ssh_client == 'ssh':
        cmd = ['ssh', f'{USER}@{HOST}', '-p', str(PORT)]
        if command:
            cmd.append(command)
        result = subprocess.run(cmd)
        return result.returncode == 0

    elif ssh_client == 'plink':
        cmd = ['plink', '-ssh', f'{USER}@{HOST}', '-P', str(PORT), '-pw', PASSWORD]
        if command:
            cmd.append(command)
        result = subprocess.run(cmd, capture_output=True, text=True)
        print(result.stdout)
        if result.stderr:
            print(result.stderr)
        return result.returncode == 0

def interactive_connect():
    """交互式SSH会话"""
    print(f'🔐 连接 {USER}@{HOST}:{PORT}...')
    import shutil

    ssh_client = check_ssh_client()

    if ssh_client == 'ssh':
        cmd = ['ssh', f'{USER}@{HOST}', '-p', str(PORT)]
        print(f'📍 执行: {" ".join(cmd)}')
        os.system(' '.join(cmd))
    elif ssh_client == 'plink':
        cmd = f'plink -ssh {USER}@{HOST} -P {PORT} -pw {PASSWORD}'
        print(f'📍 执行: {cmd}')
        os.system(cmd)
    else:
        print('❌ 未找到SSH客户端')
        return False

    return True

def main():
    if len(sys.argv) > 1:
        command = ' '.join(sys.argv[1:])
        print(f'📤 执行命令: {command}')
        success = connect_ssh(command)
        if success:
            print('✅ 命令执行成功')
        else:
            print('❌ 命令执行失败')
    else:
        interactive_connect()

if __name__ == '__main__':
    import shutil
    main()