"""主程序：将 PDF 转换为 PNG 图片，然后逐张调用截图工具进行处理"""

import dis
import os
import time
import threading
import cv2
import shutil
import glob
import argparse
from notebooklm2ppt.cli import main

if __name__ == "__main__":
    main()
