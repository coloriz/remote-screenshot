Remote Screenshot
---
This program captures a screenshot of the remote computer's session using ssh connection.

## Requirements

### Local
- Python 3.9
- OpenCV (for JPEG decoding, GUI Viewer)
- OpenSSH Client
- Credential of privileged user on the remote computer

### Remote
- Windows 10 (build 1809 and later)
- OpenSSH Server Feature
- `PsExec.exe` in %PATH%
- Set `DefaultShell` registry key

## Usage
```text
usage: main.py [-h] [--shell-type {cmd,powershell}] [--check-host-key] [--ssh-executable SSH_EXECUTABLE] [--ssh-args SSH_ARGS] [--control-path CONTROL_PATH] [--control-path-dir CONTROL_PATH_DIR] [--timeout TIMEOUT] host

positional arguments:
  host                  Hostname/ip to connect to.

optional arguments:
  -h, --help            show this help message and exit
  --shell-type {cmd,powershell}
                        The `DefaultShell` configured on the host. (default: powershell)
  --check-host-key      Let ssh check host keys. (default: False)
  --ssh-executable SSH_EXECUTABLE
                        The location of the ssh binary. (default: ssh)
  --ssh-args SSH_ARGS   Arguments to pass to ssh subprocess calls. (default: -C -o ControlMaster=auto -o ControlPersist=60s -o PreferredAuthentications=publickey -o PasswordAuthentication=no)
  --control-path CONTROL_PATH
                        ssh's ControlPath socket filename. (default: ssh-%h-%p-%r.sock)
  --control-path-dir CONTROL_PATH_DIR
                        The directory to use for ssh control path. (default: .)
  --timeout TIMEOUT     The ammount of time to wait while establishing an ssh connection. (default: 10)
```

### Example
```text
python main.py my_host
```
Add option `--shell-type cmd` when `DefaultShell` is `cmd.exe` or unset.

On success, it pops up the screenshot taken from the remote host. Press `ESC` or `Q` to quit, 
or press any other key to re-take.
