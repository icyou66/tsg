import json
import time
import winreg
import tkinter
import requests
from tkinter import *
from tkinter import ttk
from Gui import Pygui
from tkinter import messagebox
from threading import Thread
from bs4 import BeautifulSoup
import datetime


class Tkinter:
    seat_list = dict()
    headers = {
        "Connection": "keep-alive",
        "Host": "hevttc.q2q365.com",
        "Accept-Language": "zh-CN,zh",
        "Accept-Encoding": "gzip, deflate",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 8_3 like Mac OS X) AppleWebKit/600.1.4 (KHTML, like Gecko) Version/8.0 Mobile/12F70 Safari/600.1.4",
    }
    room = [
        ("昌黎-二楼-201", 79),
        ("昌黎-三楼-301", 80),
        ("昌黎-三楼-302", 84),
        ("昌黎-三楼-303", 85),
        ("昌黎-四楼-403", 81),
        ("秦皇岛-二楼-204", 95),
        ("秦皇岛-二楼-大厅", 96),
        ("秦皇岛-三楼-305", 93),
        ("秦皇岛-四楼-401", 92),
        ("秦皇岛-四楼-403", 98),
        ("秦皇岛-五楼-502", 97)
    ]

    def __init__(self):
        self.root = Tk()
        self.root.title("图书馆抢座系统")
        window_height = self.root.winfo_screenheight()
        window_width = self.root.winfo_screenwidth()
        width, height = list([500, 600])
        self.root.geometry(
            "%dx%d+%d+%d"
            % (width, height, (window_width - width) / 2, (window_height - height) / 2)
        )
        self.root.attributes("-alpha", 1)

        self.create_place()  # 创建窗口
        self.root.mainloop()  # 使窗口等待

    def create_place(self):
        self.info_place_func()
        self.log_place_func()
        self.read_info()

    def info_place_func(self):
        place = Frame(self.root)
        place.pack(side="top", fill=BOTH)
        info_place = LabelFrame(place, text="信息配置")
        info_place.pack(ipadx=5, ipady=5, fill="x", padx=10)

        top = Frame(info_place)
        top.pack(side="top")

        Label(top, text="座位号：").grid(row=0, column=0, padx=(5, 0), pady=5)
        self.seat = tkinter.StringVar()
        Entry(top, textvariable=self.seat).grid(row=0, column=1)

        self.select = tkinter.StringVar()
        select_label = Label(top, text="房间号：")
        select_label.grid(row=1, column=0, padx=(5, 0), pady=5, sticky="w")
        select_value = ttk.Combobox(top, width=17, textvariable=self.select, state="readonly")
        select_value['values'] = [x[0] for x in self.room]
        select_value.grid(row=1, column=1, sticky="w")

        Button(top, text="保存信息", command=lambda: self.save_info(), relief="groove", background="#C0E3FF").grid(
            row=0, column=5, rowspan=2, ipadx=5, ipady=10, padx=10, pady=5)

        bot = Frame(info_place)
        bot.pack(side="top")

        Button(bot, text="运行程序", command=lambda: self.run_thread(self.run), relief="groove",
               background="#D4DAFF").grid(row=2, column=0, padx=5, pady=5, ipadx=10)
        Button(bot, text="取消预约", command=lambda: self.run_thread(self.cancel), relief="groove",
               background="#FFE1F4").grid(row=2, column=1, padx=5, pady=5, ipadx=10)
        Button(bot, text="重复占座", command=lambda: self.run_thread(self.repeat), relief="groove",
               background="#FFFBBB").grid(row=2, column=2, padx=5, pady=5, ipadx=10)

    def log_place_func(self):
        # 日志空间
        self.log_place = Frame(self.root)
        self.log_place.pack(expand=tkinter.YES, fill=tkinter.BOTH, side="bottom", padx=5, pady=5, ipadx=5, ipady=5)
        self.log_data = Text(self.log_place, font="微软雅黑 9", height=16, fg="blue", bg="#F4F3FF")
        self.log_data.pack(expand=tkinter.YES, fill=tkinter.BOTH, side="bottom", padx=5, pady=5, ipadx=5, ipady=5)
        self.log_data.tag_config("blue", foreground="blue")
        self.log_data.tag_config("red", foreground="red")
        self.log_data.tag_config("green", foreground="green")
        self.pprint("Tip：操作后这里显示日志")

        roll_logy = tkinter.Scrollbar(self.log_data)
        roll_logy.pack(side="right", fill="y")
        self.log_data.config(yscrollcommand=roll_logy.set)
        roll_logy.config(command=self.log_data.yview)

    def read_info(self):
        self.pprint("读取配置中...")
        open(f"./cookie.txt", "a+", encoding="utf-8-sig").close()
        open(f"./openid.txt", "a+", encoding="utf-8-sig").close()
        open(f"./sid.txt", "a+", encoding="utf-8-sig").close()
        open(f"./info.json", "a+", encoding="utf-8-sig").close()
        with open(f"./cookie.txt", "r+", encoding="utf-8-sig") as file:
            self.cookie = file.read()
        with open(f"./openid.txt", "r+", encoding="utf-8-sig") as file:
            self.openid = file.read()
        with open(f"./sid.txt", "r+", encoding="utf-8-sig") as file:
            self.sid = file.read()
        with open(f"./info.json", "r+", encoding="utf-8-sig") as file:
            info = file.read()
        if info:
            info = json.loads(info)
            self.select.set(value=info.get("select"))
            self.seat.set(value=info.get("seat"))

        if not self.cookie or not self.openid or not self.sid:
            self.run_thread(self.waitcookie)
            # messagebox.showwarning("提示", "cookie等必要参数未读取成功！请按照流程来！")
            # self.root.destroy()
        else:
            self.pprint("信息配置读取成功！正在效验cookie有效性...", "green")
            self.headers['Cookie'] = self.cookie
            url = f"http://hevttc.q2q365.com/qckj/seat/person?sid={self.sid}&openid={self.openid}"
            result = requests.get(url, headers=self.headers).text
            if "座位管理平台" in result:
                self.pprint("cookie效验成功！即将进行下一步操作...\n", "green")
                if not self.seat_check(result):
                    self.run_thread(self.run)
            else:
                self.run_thread(self.waitcookie)
                # messagebox.showwarning("提示", "cookie已过期，请重新按照流程操作一遍！")
                # self.root.destroy()

    def seat_check(self, result):
        if "已签到" in result:
            soup = BeautifulSoup(result, "html.parser")
            search = soup.find("p", class_="daohang_step_des1")
            search2 = soup.find_all("p", class_="daohang_step_des2")
            if search:
                self.pprint(f"你已经占位，信息如下：\n{search.text} 签到时间：{search2[1].text}")
        elif "已预约" in result:
            soup = BeautifulSoup(result, "html.parser")
            search = soup.find_all("p", class_="daohang_step_des2")
            if search:
                self.pprint(f"你已经占位，信息如下：\n{search[0].text}\n{search[1].text}")
        else:
            return False
        return True

    def waitcookie(self):
        try:
            proxy_address = "127.0.0.1:8080"
            internet_settings = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                               r'Software\Microsoft\Windows\CurrentVersion\Internet Settings', 0,
                                               winreg.KEY_ALL_ACCESS)
            winreg.SetValueEx(internet_settings, "ProxyServer", 0, winreg.REG_SZ, proxy_address)
            winreg.SetValueEx(internet_settings, "ProxyEnable", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(internet_settings)
            self.pprint("检测到cookie无效！即将运行cookie获取模式...")
            time.sleep(1)
            gui = Pygui()
            if gui.preg_match():
                self.pprint(f"已捕获刷新按钮，坐标：({gui.x}, {gui.y})。运行期间请勿关闭或遮挡刷新按钮。否则将会失效！")
            else:
                self.end("未捕获到刷新按钮，请重新按照流程操作！")
            self.pprint("已打开本地代理，等待5：20时将运行下一步...\n")
            self.time_sleep(datetime.time(10, 46, 25))
            if gui.preg_match():
                self.pprint(f"已捕获刷新按钮，坐标：({gui.x}, {gui.y})。")
                gui.click()
                internet_settings = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                                   r'Software\Microsoft\Windows\CurrentVersion\Internet Settings', 0,
                                                   winreg.KEY_ALL_ACCESS)
                winreg.SetValueEx(internet_settings, "ProxyEnable", 0, winreg.REG_DWORD, 0)
                winreg.CloseKey(internet_settings)
                self.pprint(f"已点击刷新按钮。10秒后重新检测cookie....")
                time.sleep(10)
                self.read_info()
            else:
                self.pprint("刷新按钮未捕获成功！cookie运行失败！")
        except Exception as e:
            self.end(str(e))

    @staticmethod
    def run_thread(func):
        Thread(target=func).start()

    def run(self):
        seat = self.seat.get().strip()
        select = self.select_room()
        if not seat or not select:
            self.end("信息未填写完整，请填写后点击“运行程序” ！")
        url = f"http://hevttc.q2q365.com/qckj/seat/seat_status?sid={self.sid}&openid={self.openid}&l_id={select}"
        self.pprint("\n尝试进入房间...")
        result = requests.get(url, headers=self.headers).text
        if "seat-row" not in result:
            self.end("进入房间失败！请联系管理员！", "red")
        soup = BeautifulSoup(result, "html.parser")

        seat_list = [x.text for x in soup.find_all("p", class_="seat_no")]
        if seat not in seat_list:
            self.end(f"座位号{seat}不存在！请注意不到百位的座位号要在前面补0，例如21座位需要填：021")

        search = soup.find_all("a", attrs={"data-dialog-ajax": "true"})
        for seats in search:
            p = seats.find("p")
            if p:
                seat_p = seats.find("p").text
                url = seats['href']
                self.seat_list[seat_p] = url
        url = self.seat_list.get(seat, False)
        if not url:
            self.end(f"{seat}已经被占用！")
        self.select_submit(url, seat)

    def cancel(self):
        self.pprint("\n检查座位状态...")
        url = f'http://hevttc.q2q365.com/qckj/seat/person?sid={self.sid}&openid={self.openid}&r=' + str(
            int(time.time()))
        result = requests.get(url, headers=self.headers, timeout=5).text
        if "高校鸿芒" not in result:
            self.end("未知错误！请联系管理员！", "red")
        if "已预约" in result:
            soup = BeautifulSoup(result, "html.parser")
            search = soup.find_all("p", class_="daohang_step_des2")
            if search:
                self.pprint(f"你的座位信息如下：\n{search[0].text}\n{search[1].text}")
                self.pprint("正在退座中...")
            url = soup.find("a", class_="autodialog")['href']
            result = requests.get(url, headers=self.headers, timeout=5).text
            # 取消预约form表单
            soup = BeautifulSoup(result, "html.parser")
            # 拼接url
            base_url = "http://hevttc.q2q365.com"
            param_url = soup.find("form", attrs={"id": "subscribecancle"})['action']
            url = base_url + param_url
            # 获取data内容
            data = dict()
            data['openid'] = soup.find("input", attrs={"name": "openid"})['value']
            data['locate'] = soup.find("input", attrs={"name": "locate"})['value']
            data['floor'] = soup.find("input", attrs={"name": "floor"})['value']
            data['typeseat'] = soup.find("input", attrs={"name": "typeseat"})['value']
            data['op'] = soup.find("input", attrs={"name": "op"})['value']
            data['form_build_id'] = soup.find("input", attrs={"name": "form_build_id"})['value']
            data['form_id'] = soup.find("input", attrs={"name": "form_id"})['value']

            # 提交退座表单
            requests.post(url, data=data, headers=self.headers)

            # 访问首页确认是否退座成功
            result = self.check()
            if "扫座位码选座前请先扫描阅览室确认码进行签到" in result:
                self.pprint("退座成功！")
            else:
                self.end("退座失败！")
        else:
            self.end("你还未预约座位！")

    def repeat(self):
        self.cancel()
        self.run()

    def select_submit(self, url, seat):
        self.time_check()
        self.pprint(f"\n尝试提交座位：{seat}\n")
        while True:
            result = requests.get(url, headers=self.headers, timeout=5).text
            if "预约成功后需在30分钟内签到" in result:
                soup = BeautifulSoup(result, "html.parser")
                url = soup.find("div", class_="subscribe_today").find("a")['href']
                requests.get(url, headers=self.headers, timeout=5)

                # 访问首页确认是否选座成功
                result = self.check()
                soup = BeautifulSoup(result, "html.parser")
                search = soup.find_all("p", class_="daohang_step_des2")
                if search:
                    self.end(f"占座成功！信息如下：{search[0].text}\n{search[1].text}")
            elif "此座位已被预约" in result:
                self.end("此座位已被预约！")
            elif "实时预约" in result:
                self.pprint(f"【{self.get_time()}】未到预约时间！", delete=True)
                time.sleep(0.5)

    def check(self):
        url = f"http://hevttc.q2q365.com/qckj/seat/person?sid={self.sid}&openid={self.openid}"
        result = requests.get(url, headers=self.headers).text
        return result

    def select_room(self):
        select = self.select.get()
        for s in self.room:
            if s[0] == select:
                return s[1]
        return False

    @staticmethod
    def get_time():
        current_time = datetime.datetime.now()
        formatted_time = current_time.strftime("%Y-%m-%d %H:%M:%S")
        return formatted_time

    def time_check(self):
        target_time = datetime.time(6, 00, 00)
        end_time = datetime.time(21, 00, 00)
        current_time = datetime.datetime.now().time()
        if current_time > end_time:
            self.end(f"\n当前时间为【{self.get_time()}】\n已经过了抢座时间，请明天再来！")

        if current_time < target_time:
            self.pprint(f"未到抢座时间，设定06:00:00开始运行占座程序...\n")
        while True:
            current_time = datetime.datetime.now().time()
            self.pprint(f"当前时间：{self.get_time()}", delete=True)
            if current_time > target_time:
                return True
            else:
                time.sleep(0.2)

    def time_sleep(self, sleep_time):
        while True:
            self.pprint(f"当前时间：{self.get_time()}", delete=True)
            current_time = datetime.datetime.now().time()
            if current_time > sleep_time:
                if current_time.hour < sleep_time.hour + 2:
                    return True
            time.sleep(1)

    def save_info(self):
        seat = self.seat.get().strip()
        select = self.select.get()
        if not seat or not select:
            messagebox.showinfo("提示", "你还没有填好默认信息！")
        else:
            info = json.dumps(dict(seat=seat, select=select))
            open(f"./info.json", "a+", encoding="utf-8-sig").close()
            with open(f"./info.json", "r+", encoding="utf-8-sig") as file:
                file.write(info)
            messagebox.showinfo("提示", "保存成功，下次打开软件时将自动为你设置该配置！")

    def pprint(self, msg, color=None, delete=False):
        if not color:
            color = "blue"
        if delete:
            start = self.log_data.index("end-2l")
            end = self.log_data.index("end-1c")
            self.log_data.delete(start, end)
        self.log_data.insert(END, msg + "\n", color)
        self.log_data.update()
        self.log_data.see(END)

    def end(self, msg, color=None, delete=False):
        if not color:
            color = "blue"
        if delete:
            start = self.log_data.index("end-2l")
            end = self.log_data.index("end-1c")
            self.log_data.delete(start, end)
        self.log_data.insert(END, msg + "\n", color)
        self.log_data.update()
        self.log_data.see(END)
        raise SystemExit


if __name__ == "__main__":
    Tkinter()
