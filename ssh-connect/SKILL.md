---
name: "ssh-connect"
description: "SSH connection tool for remote server management. Invoke when user needs to connect to remote servers via SSH or mentions 'armbian'."
---

# SSH Connect

Simple SSH connection tool for accessing remote servers.

## Connection Details

- **Host**: 192.168.2.2
- **Hostname**: armbian
- **Port**: 22
- **Username**: root
- **Password**: admincw

## Auto-Connect Keywords

当用户提到以下关键词时，自动连接到 armbian (192.168.2.2)：
- "ssh到armbian"
- "连接armbian"
- "登录armbian"
- "远程连接armbian"
- "armbian服务器"

## Usage

### Interactive Session

```bash
python ssh_connect.py
```

### Execute Command

```bash
python ssh_connect.py "ls -la /root"
python ssh_connect.py "uptime; free -h; df -h"
```

## Security Warning

⚠️ **This skill contains hardcoded credentials. Use only in trusted environments.**

- Credentials are stored in plain text
- Suitable for internal/private networks only
- Consider using SSH keys for production environments

## Troubleshooting

### Connection Timeout

- Check network connectivity
- Verify firewall settings on both sides
- Ensure SSH service is running on remote server

### Authentication Failed

- Verify username and password
- Check if SSH permits password authentication
- Ensure the account is not locked

### Permission Denied

- Verify user has SSH access
- Check user groups and permissions on remote server