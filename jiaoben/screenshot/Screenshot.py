# coding: utf - 8
import os
import itertools

'''
自动截图n张
'''


class Screenshot():

    def Screenshot(self):
        i = 1
        mylist = ("".join(x) for x in
                  itertools.product("123abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", repeat=16))
        while i < 500:
            a = (next(mylist))
            i = i + 1
            # os.system("adb shell screencap -p /sdcard/%s.png" % (a))
            os.system("adb shell screencap -p /sdcard/DCIM/Screenshots/%s.png" % (a))
            # print(a)


if __name__ == '__main__':
    c = Screenshot()
    c.Screenshot()
