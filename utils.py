import base64
import sys

import cv2 as cv
import numpy as np


def print_error(*args, **kwargs):
    kwargs['file'] = sys.stderr
    print(*args, **kwargs)


def print_quser(quser):
    quser_format = '{id:>4}  {user:<20}  {session:<20}  {state:<8}  {logon_time}'
    row = quser_format.format(
        id=quser['Id'],
        user=quser['UserName'],
        session=quser['SessionName'],
        state=quser['State'],
        logon_time=quser['LogonTime']
    )
    print(row)


def convert_base64_to_image(b64_str):
    raw_bytes = base64.b64decode(b64_str)
    return cv.imdecode(np.frombuffer(raw_bytes, np.uint8), cv.IMREAD_UNCHANGED)
