# # encoding=utf-8
# import subprocess
#
#
# # import win32clipboard as w
# # import win32con
#
# class Clipboard(object):
#     '''
#     模拟Windows设置剪切板
#     '''
#
#     # 读取剪切板
#     @staticmethod
#     def getText():
#         # # 打开剪切板
#         # w.OpenClipboard()
#         # # 获取剪切板中的数据
#         # d = w.GetClipboardData(win32con.CF_TEXT)
#         # # 关闭剪切板
#         # w.CloseClipboard()
#         # # 返回剪切板数据给调用者
#         # return d
#         p = subprocess.Popen(['pbpaste'], stdout=subprocess.PIPE)
#         retcode = p.wait()
#         data = p.stdout.read()
#         # 这里的data为bytes类型，之后需要转成utf-8操作
#         return data
#
#     # 设置剪切板内容
#     @staticmethod
#     def setText(aString):
#         # # 打开剪切板
#         # w.OpenClipboard()
#         # # 清空剪切板
#         # w.EmptyClipboard()
#         # # 将数据aString写入剪切板
#         # w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
#         # # 关闭剪切板
#         # w.CloseClipboard()
#         p = subprocess.Popen(['pbcopy'], stdin=subprocess.PIPE)
#         p.stdin.write(aString.encode("utf-8"))
#         p.stdin.close()
#         p.communicate()
#
# if __name__ == '__main__':
#     Clipboard().setText("a.txt")
