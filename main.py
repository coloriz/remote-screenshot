import sys
from argparse import ArgumentParser, ArgumentDefaultsHelpFormatter
from pathlib import Path
from subprocess import CalledProcessError

import cv2 as cv

from rpcstub import CMDStub, PowerShellStub
from utils import (
    print_error,
    print_quser,
    convert_base64_to_image,
)


def main():
    parser = ArgumentParser(formatter_class=ArgumentDefaultsHelpFormatter)
    parser.add_argument('host',
                        help='Hostname/ip to connect to.')
    parser.add_argument('--shell-type', default='powershell', choices=['cmd', 'powershell'],
                        help='The `DefaultShell` configured on the host.')
    parser.add_argument('--check-host-key', action='store_true',
                        help='Let ssh check host keys.')
    parser.add_argument('--ssh-executable', default='ssh',
                        help='The location of the ssh binary.')
    parser.add_argument('--ssh-args', default='-C -o ControlMaster=auto -o ControlPersist=60s '
                                              '-o PreferredAuthentications=publickey -o PasswordAuthentication=no',
                        help='Arguments to pass to ssh subprocess calls.')
    parser.add_argument('--control-path', default='ssh-%h-%p-%r.sock',
                        help="ssh's ControlPath socket filename.")
    parser.add_argument('--control-path-dir', default='.',
                        help='The directory to use for ssh control path.')
    parser.add_argument('--timeout', default=10, type=int,
                        help='The ammount of time to wait while establishing an ssh connection.')
    opt = parser.parse_args()

    ssh_executable = opt.ssh_executable
    ssh_args = opt.ssh_args
    host = opt.host

    if not opt.check_host_key:
        ssh_args += ' -o StrictHostKeyChecking=no'

    ssh_args += f' -o ConnectTimeout={opt.timeout}'

    if 'ControlPath' not in ssh_args:
        control_path_dir = Path(opt.control_path_dir).resolve()
        control_path = control_path_dir / opt.control_path
        ssh_args += f' -o "ControlPath={control_path}"'

    if opt.shell_type == 'cmd':
        stub = CMDStub(ssh_executable, ssh_args, host)
    elif opt.shell_type == 'powershell':
        stub = PowerShellStub(ssh_executable, ssh_args, host)
    else:
        raise NotImplementedError

    try:
        env = stub.gather_env()
        qusers = stub.get_quser()

        if not qusers:
            print_error(f'No user exists.')
            return 1

        # Print user sessions
        header_row = {
            'Id': 'ID',
            'UserName': 'UserName',
            'SessionName': 'SessionName',
            'State': 'State',
            'LogonTime': 'Logon Time',
        }
        print_quser(header_row)
        for quser in qusers:
            print_quser(quser)

        # Select session to capture
        if len(qusers) == 1:
            quser = qusers[0]
        else:
            while True:
                ans = input('Select session id to capture: ')
                # Find selected quser object
                for quser in qusers:
                    if str(quser['Id']) == ans:
                        break
                else:
                    print_error(f'{ans!r} is not valid!')
                    continue
                break

        window_name = f"session_id={quser['Id']} ({quser['UserName']}@{env['COMPUTERNAME']})"

        while True:
            data = stub.take_screenshot(quser['Id'])
            timestamp = data['LastWriteTime']
            img = convert_base64_to_image(data['Data'])
            print(f'{timestamp=}')
            cv.imshow(window_name, img)

            key = cv.waitKey() & 0xFF
            if key in (27, 113):  # ESC, q
                break

        cv.destroyAllWindows()
    except CalledProcessError as e:
        print(e.stderr)
        return e.returncode
    finally:
        code = stub.cleanup()
        if code:
            print_error(f'Exit code of cleanup is non-zero: {code}')

    return 0


if __name__ == '__main__':
    sys.exit(main())
