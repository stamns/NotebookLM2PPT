import sys
from PIL import Image, ImageTk

# 根据 Python 版本兼容性导入 tkinter
if sys.version_info[0] == 2:
    import Tkinter as tkinter
else:
    import tkinter

def showPIL(pilImage):
    root = tkinter.Tk()
    
    # 1. 获取屏幕分辨率
    w, h = root.winfo_screenwidth(), root.winfo_screenheight()
    
    # 2. 设置无边框全屏
    root.attributes('-fullscreen', True)
    root.configure(background='black')
    root.focus_set()
    
    # 3. 绑定退出事件：按键盘上的 Esc 键关闭窗口
    root.bind("<Escape>", lambda e: (e.widget.withdraw(), e.widget.quit()))
    
    # 4. 创建画布并填充背景为黑色，去除所有边框
    canvas = tkinter.Canvas(root, width=w, height=h,
                           highlightthickness=0, borderwidth=0)
    canvas.pack(fill='both', expand=True)
    canvas.configure(background='black')
    
    # 5. 调整图片大小以适应屏幕（保持纵横比）
    imgWidth, imgHeight = pilImage.size
    if imgWidth > w or imgHeight > h:
        ratio = min(w / float(imgWidth), h / float(imgHeight))
        imgWidth = int(imgWidth * ratio)
        imgHeight = int(imgHeight * ratio)
        # 注意：Image.ANTIALIAS 在新版 Pillow 中可能变更为 Image.LANCZOS
        pilImage = pilImage.resize((imgWidth, imgHeight), Image.LANCZOS)
    
    # 6. 将 PIL 图片转换为 Tkinter 图片对象
    image = ImageTk.PhotoImage(pilImage)
    
    # 7. 在画布中央显示图片
    imagesprite = canvas.create_image(w / 2, h / 2, image=image)
    
    root.mainloop()

# 使用示例
if __name__ == "__main__":
    # 请将 "your_image.png" 替换为你本地的图片路径
    try:
        pilImage = Image.open(r"C:\Users\18350\Desktop\微信图片_20251212143558_365_141.jpg")
        showPIL(pilImage)
    except FileNotFoundError:
        print("未找到图片文件，请检查路径。")