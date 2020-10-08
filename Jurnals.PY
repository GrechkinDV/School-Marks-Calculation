from kivy.app import App
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.uix.textinput import TextInput
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.uix.popup import Popup
from kivy.uix.image import Image
from kivy.uix.button import ButtonBehavior, Button
from kivy.uix.label import Label
from kivy.uix.dropdown import DropDown
from kivy.config import Config
from kivy.uix.checkbox import CheckBox
import sys
import os
import codecs
import math
import openpyxl
import xlsxwriter
import icu
import string
from win32api import GetSystemMetrics

Config.set('graphics', 'multisamples', '0')


gl_ids = {}

lb = """
Label:
    text: "Məktəbin Adı"
    size_hint:(0.2, 0.1)
    font_size: self.size[1]//3.5
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        Rectangle:
            pos: self.pos
            size: self.size
"""

logotype = """
Label:
    text: "Məktəbi Jurnal"
    font_size: self.size[0]//13
    size_hint: (0.5, 0.1)
    pos_hint:{"x":0.25, "top":0.93}
    canvas.before:
        Color:
            rgba: (60/255, 58/255, 37/255, 1)
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [17,17,17, 17]
"""

Builder.load_string("""
<P>:
    com: com
    next: next
    Button:
        id: next
        text: "Növbəti"
        background_color: (60/255, 58/255, 37/255, 1)
        font_size: root.width//25
        size_hint: 0.26,0.2
        pos_hint: {"x":0.37, "top":0.36}
    Label:
        font_size: root.width//45
        size_hint: 0.82, 0.4
        pos_hint: {"x":0.09, "top":0.83} 
        id: com
        text_size: [root.width*0.78, None]
        size_hint_y: None
    Label:
        text: "Produced by Dmitry Grechkin Vladimirovich. Future GDV Inc."
        font_size: root.width//38
        size_hint: 0.5, 0.07
        pos_hint: {"x":0.25, "top":.1}
        color: (1,43/255,18/255,1)
        
<AddJournal>:
    FloatLayout:
        TextInput:
            id: grade
            font_size: root.width//20
            size_hint: 0.16, 0.2
            pos_hint: {"top": 1.4, "x": 0.76}
            multiline: False
            padding: (self.width//5, self.height//5)
        Label:
            text: "Sinif:"
            pos_hint: {"top":1.4, "x":0.5}
            size_hint:(0.2, 0.2)
            font_size: root.width//20
            canvas.before:
                Color:
                    rgba: (12/255,63/255,82/255,1)
                Rectangle:
                    pos: self.pos
                    size: self.size
        TextInput:
            id: pow
            font_size: root.width//30
            size_hint: 0.06, 0.12
            pos_hint: {"top": 1.5, "x":0.94}
            multiline: False
            padding: (self.width//3, self.height//10)
        Label:
            text: "Şagirdlər Sayı:"
            pos_hint: {"top":1.1, "x":0.5}
            size_hint:(0.2, 0.2)
            font_size: root.width//35
            canvas.before:
                Color:
                    rgba: (12/255,63/255,82/255,1)
                Rectangle:
                    pos: self.pos
                    size: self.size        
        TextInput:
            id: say
            font_size: root.width//20
            size_hint: 0.16, 0.2
            pos_hint: {"top": 1.1, "x": 0.76}
            multiline: False
            padding: (self.width//3, self.height//5)
        Label:
            text: "Produced by Dmitry Grechkin Vladimirovich. Future GDV Inc."
            font_size: root.width//38
            size_hint: 0.5, 0.07
            pos_hint: {"x":0.61, "top":0.61}
            color: (1,43/255,18/255,1)
        Button:
            id: yarat
            on_press: root.create(say.text, grade.text, pow.text)
            text: "Yarat"
            background_color: (60/255, 58/255, 37/255, 1)
            font_size: root.width//25
            size_hint: 0.26,0.2
            pos_hint: {"x":0.74, "top":0.83}

<Yarat>:
    FloatLayout:
        Label:
            text: "Şagirdlərin Siyahısı Doldurun(Proqram əlifba sırasını özu qurur)"
            font_size: self.size[0]//33
            size_hint: (0.7, 0.1)
            pos_hint:{"x":0.2, "top":1.04}
            canvas.before:
                Color:
                    rgba: (60/255, 58/255, 37/255, 1)
                RoundedRectangle:
                    pos: self.pos
                    size: self.size
                    radius: [17,17,17, 17]
    Button:
        id: next
        text: "Təsdiqlə"
        background_color: (60/255, 58/255, 37/255, 1)
        font_size: root.width//29
        size_hint: 0.26,0.1
        pos_hint: {"x":0.37, "top":0.2}
        on_press: root.listing()
        
<Subject>:
    crt: crt
    Label:
        text: "Produced by Dmitry Grechkin Vladimirovich. Future GDV Inc."
        font_size: root.width//38
        size_hint: 0.5, 0.07
        pos_hint: {"x":0.25, "top":.1}
        color: (1,43/255,18/255,1)
        
    Button:
        id: crt
        text: "Yarat"
        on_press: root.add_subject()
        background_color: (60/255, 58/255, 37/255, 1)
        font_size: root.width//25
        size_hint: 0.26,0.15
        pos_hint: {"x":0.37, "top":0.3}
        
<Parameters>:
    canvas.before:
        Color:
            rgba: (1, 1, 1, 1)
        Line:
            rectangle: (root.width*0.351, root.height//6.9, 1, root.height*0.83)
            width: 3
            
        Line:
            rectangle: (root.width*0.702, root.height//6.9, 1, root.height*0.83)
            width: 3

    Label:
        text: "Produced by Dmitry Grechkin Vladimirovich. Future GDV Inc."
        font_size: root.width//38
        size_hint: 0.5, 0.07
        pos_hint: {"x":0.25, "top":.1}
        color: (1,43/255,18/255,1)

    Label:
        text: "Sinif Rəhbəri Dəyişmək"
        size_hint:  (0.19, 0.11)
        pos_hint: {"top": 0.92, "x": 0.06}
        font_size: self.size[1]//3.5
        canvas.before:
            Color:
                rgba: (12/255,63/255,82/255,1)
            RoundedRectangle:
                pos: self.pos
                size: self.size   
                radius: [20,20,20,20]

    Label:
        text: "Fənni Silmək"
        size_hint:  (0.19, 0.11)
        pos_hint: {"top": 0.92, "x": 0.747}
        font_size: self.size[1]//3.35
        canvas.before:
            Color:
                rgba: (12/255,63/255,82/255,1)
            RoundedRectangle:
                pos: self.pos
                size: self.size 
                radius: [20,20,20,20]
       
    Label:
        text: "Şagirlərin Siyahısı Dəyişmək"
        size_hint:  (0.19, 0.11)
        pos_hint: {"top": 0.92, "x": 0.4}
        font_size: self.size[1]//4
        canvas.before:
            Color:
                rgba: (12/255,63/255,82/255,1)
            RoundedRectangle:
                pos: self.pos
                size: self.size 
                radius: [20,20,20,20]
                
    Button:
        text: "İndiki Jurnalı Sil"
        color: (112/255,8/255,5/255,1)
        bold: True
        on_press: app.delete_journal(root.grade)
        font_size: self.size[0]//10
        size_hint: 0.3, 0.1
        pos_hint: {"x":0.345, "top":0.33}
        background_color: (0,0,0,0)
        canvas.before:
            Color:
                rgba: (112/255,8/255,5/255,1)
            Line:
                rounded_rectangle: (self.pos[0],self.pos[1], self.size[0], self.size[1],12)
                width: 3
    Button:
        text: "Dəyiş"
        background_color: (60/255, 58/255, 37/255, 1)
        font_size: root.width//70
        size_hint: 0.25,0.08
        pos_hint: {"x":0.0288, "top":0.62}    
        on_press: root.deyish()
    
    Button:
        text: "Sil"
        background_color: (60/255, 58/255, 37/255, 1)
        font_size: root.width//65
        size_hint: 0.25,0.08
        pos_hint: {"x":0.72, "top":0.62}    
        on_press: root.sil()
        
<FinalMarks>:
    Label:
        text: "Müəllim"
        size_hint:(0.06, 0.07)
        pos_hint: {"x":0.05, "top":0.9}
        font_size: self.size[1]//3.5
        canvas.before:
            Color:
                rgba: (12/255,63/255,82/255,1)
            Rectangle:
                pos: self.pos
                size: self.size
""")

sm = ScreenManager()


class Parameters(FloatLayout):
    def __init__(self, gr):
        super(Parameters, self).__init__()
        self.dd = DropDown()
        self.dd2 = DropDown()
        self.dd3 = DropDown()
        self.dd.max_height = Window.height//5.5
        self.dd2.max_height = Window.height//5.5
        self.dd3.max_height = Window.height//5.5
        self.dd4 = DropDown()
        self.dd4.max_height = Window.height//5.5
        self.teachers = []
        self.grade = gr
        self.head_teacher = ""
        self.dlt = None
        self.add_s()

        self.exit = EditableImage()
        self.exit.source = "sfx/close2.png"
        self.exit.size_hint = (0.04, 0.08)
        self.exit.pos_hint = {"x": 0.96, "top": 1}
        self.exit.on_press = ListSubjects.config_popup.dismiss
        self.add_widget(self.exit)

        self.mnbutton = Button(size_hint=(0.25, 0.11), pos_hint={"x": 0.0275, "top": 0.76},
                               font_size=Window.width//73,
                               background_normal="sfx/selection2.png")
        self.mnbutton.text = "Indiki Sinif Rəhbəri:\n" + self.head_teacher
        self.mnbutton.bind(on_release=self.dd.open)
        self.add_widget(self.mnbutton)

        self.add_s2()
        self.mnbutton2 = Button(size_hint=(0.25, 0.11), pos_hint={"x": 0.72, "top": 0.76},
                                font_size=Window.width//73,
                                background_normal="sfx/selection2.png")
        self.mnbutton2.text = "Silinən Fənn:\n"
        self.mnbutton2.bind(on_release=self.dd2.open)
        self.add_widget(self.mnbutton2)

        self.mnbutton3 = Button(size_hint=(0.25, 0.11), pos_hint={"x": 0.372, "top": 0.76},
                                font_size=Window.width//73,
                                background_normal="sfx/selection2.png")
        self.mnbutton3.text = "Əməliyyəti Seçin:\n"
        self.mnbutton3.bind(on_release=self.dd3.open)
        self.add_widget(self.mnbutton3)

        for i in ["Şagirdi Əlavə Elə", "Şagirdi Sil"]:
            self.subj = Button(text=i, font_size=Window.height//45,
                               background_color=(12/255, 63/255, 82/255, 1),
                               size_hint_y=None, height=37)
            self.subj.bind(on_release=lambda subj: self.dd3.select(subj))
            self.dd3.add_widget(self.subj)

        self.dd3.bind(on_select=lambda instance, x: self.release3(x))
        self.add_widget(self.dd3)

    def release3(self, x):
        setattr(self.mnbutton3, "text", self.mnbutton3.text[:18] + x.text)
        self.dd4 = DropDown()
        self.dd4.max_height = Window.height//5.5
        if "Sil" in x.text:
            try:
                self.remove_widget(self.add2)
                self.remove_widget(self.yeni)
            except:
                pass
            self.opt = Button(size_hint=(0.25, 0.11), pos_hint={"x": 0.372, "top": 0.62},
                              font_size=Window.width//73,
                              background_normal="sfx/selection2.png")
            self.opt.text = "Şagird ASA seçin:\n"
            self.opt.bind(on_release=self.dd4.open)
            self.add_widget(self.opt)

            self.clean2 = Button(text="Şagirdi Sil",
                                 background_color=(60/255, 58/255, 37/255, 1),
                                 font_size=Window.width//65, size_hint=(0.25, 0.08),
                                 pos_hint={"x": 0.372, "top": 0.48},
                                 on_press=self.sil_shagird)
            self.add_widget(self.clean2)

            self.main_sheet = self.wb[self.wb.sheetnames[0]]

            for i in range(3, self.main_sheet.max_row+1):
                if "---" not in self.main_sheet.cell(row=i, column=1).value:
                    self.subj = Button(text=self.main_sheet.cell(row=i, column=1).value, font_size=Window.height//45,
                                       background_color=(12/255, 63/255, 82/255, 1),
                                       size_hint_y=None, height=37)
                    self.subj.bind(on_release=lambda subj: self.dd4.select(subj))
                    self.dd4.add_widget(self.subj)

                else:
                    continue
                self.dd4.bind(on_select=lambda instance, x: self.slc(x))

        if "Elə" in x.text:
            try:
                self.remove_widget(self.opt)
                self.remove_widget(self.clean2)
            except:
                pass
            self.yeni = Builder.load_string("""
Label:
    text: "Yeni Şagird ASA:"
    size_hint:(0.09, 0.066)
    pos_hint: {"x": 0.372, "top": 0.62}
    font_size: self.size[1]//3
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        Rectangle:
            pos: self.pos
            size: self.size""")
            self.add_widget(self.yeni)

            self.yeni_input = TextInput(size_hint=(0.152, 0.066),
                                        pos_hint={"x": 0.47, "top": 0.62},
                                        font_size=Window.height//47,
                                        padding=(Window.width//150,2),
                                        multiline=False)
            self.add_widget(self.yeni_input)

            self.add2 = Button(text="Şagirdi Əlavə Elə",
                               background_color=(60/255, 58/255, 37/255, 1),
                               font_size=Window.width//65, size_hint=(0.25, 0.08),
                               pos_hint={"x": 0.372, "top": 0.52},
                               on_press=self.add_sheetlisting)

            self.add_widget(self.add2)

    def add_sheetlisting(self, instance):
        self.names = []
        self.main_sheet = self.wb[self.wb.sheetnames[0]]
        for i in range(3, self.main_sheet.max_row+1):
            self.names.append(self.main_sheet.cell(row=i, column=1).value)
        self.names.append(self.yeni_input.text)

        self.collator = icu.Collator.createInstance(icu.Locale('az_Latn_AZ'))
        self.names = sorted(self.names, key=self.collator.getSortKey)
        self.storage = [""] * self.main_sheet.max_column
        self.entrance = [""] * self.main_sheet.max_column
        self.entrance[0] = self.yeni_input.text
        self.new_index = self.names.index(self.yeni_input.text) + 1

        for f in range(self.new_index+2, self.main_sheet.max_row+2):
            for h in range(1, self.main_sheet.max_column+1):
                self.storage[h-1] = self.main_sheet.cell(row=f, column=h).value
                self.main_sheet.cell(row=f, column=h).value = self.entrance[h-1]
            self.entrance = self.storage.copy()

        for sh in range(1, len(self.wb.sheetnames)):
            self.sht = self.wb[self.wb.sheetnames[sh]]
            self.storage = [""] * self.sht.max_column
            self.entrance = [""] * self.sht.max_column
            self.entrance[0] = self.yeni_input.text
            for f in range(self.new_index+1, self.sht.max_row+2):
                for h in range(1, self.sht.max_column+1):
                    self.storage[h-1] = self.sht.cell(row=f, column=h).value
                    self.sht.cell(row=f, column=h).value = self.entrance[h-1]
                self.entrance = self.storage.copy()

        self.wb.save(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")
        self.remove_widget(self.yeni_input)
        self.remove_widget(self.add2)
        self.remove_widget(self.yeni)

    def slc(self, x):
        setattr(self.opt, "text",self.opt.text[:18] + x.text)
        self.dlt = x

    def sil_shagird(self, instance):
        self.remove_widget(self.dd4)
        self.dd4.remove_widget(self.dlt)
        i = 3
        for i in range(i, self.main_sheet.max_row+1):
            if self.opt.text[18:] in self.main_sheet.cell(row=i, column=1).value:
                self.main_sheet.cell(row=i, column=1).value += "---"
                break
        for z in self.wb.sheetnames[1:]:
            self.sheet = self.wb[z]
            self.sheet.cell(row=i-1, column=1).value += "---"
        self.wb.save(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")
        self.remove_widget(self.opt)
        self.remove_widget(self.clean2)

    def add_s2(self):
        try:
            self.wb = openpyxl.open(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")
            self.collator = icu.Collator.createInstance(icu.Locale('az_Latn_AZ'))
            self.sheetnames = sorted(self.wb.sheetnames[1:], key=self.collator.getSortKey)
            for i in self.sheetnames:
                self.subj = Button(text=i, font_size=Window.height//45,
                                   background_color=(12/255, 63/255, 82/255, 1),
                                   size_hint_y=None, height=37)
                self.subj.bind(on_release=lambda subj: self.dd2.select(subj))
                self.dd2.add_widget(self.subj)

            self.dd2.bind(on_select=lambda instance, x: self.release2(x))
            self.add_widget(self.dd2)
        except:
            pass

    def add_s(self):
        try:
            self.info3 = []
            self.wb = openpyxl.open(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")
            for i in range(1, len(self.wb.sheetnames)):
                info_sheet = self.wb[self.wb.sheetnames[i]]
                self.info3 += (info_sheet.cell(row=1, column=1).value+",")[14:].split(",")[:-1]

            self.collator = icu.Collator.createInstance(icu.Locale('az_Latn_AZ'))
            self.info3 = sorted(self.info3, key=self.collator.getSortKey)
            for z in self.info3:
                if "Sinif Rəhbəridir" in z:
                    self.head_teacher = z[:-18]
                else:
                    self.subj = Button(text=z.strip(" "), font_size=Window.height//45,
                                       background_color=(12/255, 63/255, 82/255, 1),
                                       size_hint_y=None, height=37)
                    self.subj.bind(on_release=lambda subj: self.dd.select(subj))
                    self.dd.add_widget(self.subj)
            self.dd.bind(on_select=lambda instance, x: self.release(x))
            self.add_widget(self.dd)
        except:
            pass

    def release(self, x):
        setattr(self.mnbutton,'text', self.mnbutton.text[:22]+x.text)
        x.text = self.head_teacher
        self.head_teacher = self.mnbutton.text[22:]

    def release2(self, x):
        self.dlt = x
        setattr(self.mnbutton2, "text",
                self.mnbutton2.text[:14]+x.text)

    def deyish(self):
        for i in range(1, len(self.wb.sheetnames)):
            info_sheet = self.wb[self.wb.sheetnames[i]]
            if "Rəhbər" in info_sheet.cell(row=1, column=1).value:
                info = (info_sheet.cell(row=1, column=1).value+",")[14:].split(",")[:-1]

                new_info = "Müəllim(lər):\n"
                for z in info:
                    new_info += (z.replace("(Sinif Rəhbəridir)", "") + ",")
                info_sheet.cell(row=1, column=1).value = new_info[:-1]
            if self.mnbutton.text[23:] in info_sheet.cell(row=1, column=1).value:
                info_names = (info_sheet.cell(row=1, column=1).value+",")[14:].split(",")[:-1]
                new_info = "Müəllim(lər):\n" + self.mnbutton.text[22:] + "(Sinif Rəhbəridir),"
                for z in info_names:
                    if self.mnbutton.text[22:] != z:
                        new_info += z.strip(" ") + ","
                info_sheet.cell(row=1, column=1).value = new_info[:-1]
        self.wb.save(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")

        # info = (info_sheet.cell(row=1, column=1).value+",")[14:].split(",")[:-1]

    def sil(self):
        self.wb.remove_sheet(self.wb[self.mnbutton2.text[14:]])
        self.remove_widget(self.dd2)
        self.dd2.remove_widget(self.dlt)
        # self.add_widget(self.dd2)
        self.wb.save(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l1)
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l2)
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l3)
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l4)

        ListSubjects.great_list.layout.l1 = Builder.load_string("""
GridLayout:
    id: l1
    cols:1""")
        ListSubjects.great_list.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    cols:1""")
        ListSubjects.great_list.layout.l3 = Builder.load_string("""
GridLayout:
    id: l3
    cols:1""")
        ListSubjects.great_list.layout.l4 = Builder.load_string("""
GridLayout:
    id: l4
    cols:1""")
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l1)
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l2)
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l3)
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l4)

        ListSubjects.great_list.layout.ids = {"l1": ListSubjects.great_list.layout.l1,
                                              "l2": ListSubjects.great_list.layout.l2,
                                              "l3": ListSubjects.great_list.layout.l3,
                                              "l4": ListSubjects.great_list.layout.l4}
        self.collator = icu.Collator.createInstance(icu.Locale('az_Latn_AZ'))
        self.sheetnames = sorted(self.wb.sheetnames[1:], key=self.collator.getSortKey)
        for i in range(0,  len(self.sheetnames)):
            ListSubjects.great_list.layout.ids[ListSubjects.id_s[i % 4]].add_widget(EditableImage(src="sfx/subject.png",
                                                                                                  dest="FM", id=ListSubjects.grade, subj=self.sheetnames[i]))
            ListSubjects.great_list.layout.ids[ListSubjects.id_s[i % 4]].add_widget(Label(text="{}{}".format(" "*(Window.width//55), self.sheetnames[i]),
                                                                                          size_hint_y=None, size_hint_x=None))


class Subject(FloatLayout):
    def __init__(self):
        super(Subject, self).__init__()
        self.dd = DropDown()
        self.dd.max_height = Window.height//5.5
        self.teachername = Builder.load_string("""
Label:
    text: "Müəllimin ASA"
    size_hint:  (0.27, 0.15)
    pos_hint: {"top": 0.88, "x": 0.12}
    font_size: self.size[1]//2.3
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        Rectangle:
            pos: self.pos
            size: self.size
        """)

        self.add_widget(self.teachername)
        self.add_widget(Label(text="Qruplarla işləyən Müəlimlər ASA"
                                   " vergül ilə ayrılmalıdır!\nRəhbərin"
                                   " ASA birinci yazılmalıdır!",
                              font_size=Window.height//54.5,
                              size_hint=(0.45, 0.05),
                              pos_hint={"x": 0.42, "top": 0.96}))
        self.teachername_entry = TextInput(text="", size_hint=(0.45, 0.15),
                                           pos_hint={"x": 0.42, "top": 0.88},
                                           multiline=False,
                                           padding=(Window.width//27,
                                           Window.height//58),
                                           font_size=Window.height//39)
        self.add_widget(self.teachername_entry)
        self.subjects_list = ['Alman Dili', 'Azərbaycan Dili',
                              'Azərbaycan Tarixi', 'Biologiya',
                              'Cağırışaqədərki Hazırlıq', 'Coğrafiya', 'Ədəbiyyat',
                              'Ərəb Dili', 'Fars Dili', 'Fizika', 'Fiziki mədəniyyət',
                              'Fiziki Tərbiyyə', 'Fransız Dili', 'Həyat Bilgisi',
                              'İnformatika', 'İngilis Dili', 'Kimya', 'Qarabağ Tarixi',
                              'Məntiqi və riyazi təfəkkürün inkişafı', 'Musiqi',
                              'Ortaq Türk Tarixi', 'Riyaziyyat', 'Rus Dili',
                              'Texnologiya', 'Təsviri İncəsənət', 'Ümümi Tarix']

        for i in self.subjects_list:
            self.subj = Button(text=i, font_size=Window.height//45,
                               background_color=(12/255, 63/255, 82/255, 1),
                               size_hint_y=None, height=37)
            self.subj.bind(on_release=lambda subj: self.dd.select(subj.text))
            self.dd.add_widget(self.subj)
        self.mainbutton = Button(text='Fənn', size_hint=(0.75, 0.15),
                                 pos_hint={"x": 0.12, "top": 0.69},
                                 font_size=Window.width//43,
                                 background_normal="sfx/selection2.png")

        self.mainbutton.bind(on_release=self.dd.open)
        self.dd.bind(on_select=lambda instance, x: setattr(self.mainbutton,
                                                           'text', x))
        self.add_widget(self.mainbutton)


        self.headteacher = Builder.load_string("""
Label:
    text: "Müəllim Sinif Rəhbəridirsə, soldaki opsiya seçilməlidir."
    size_hint:  (0.71, 0.15)
    pos_hint: {"top": 0.49, "x": 0.16}
    font_size: self.size[1]//2.6
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        Rectangle:
            pos: self.pos
            size: self.size
        """)
        self.add_widget(self.headteacher)
        self.cb = CheckBox(size_hint_x=0.05, pos_hint={"x": 0.11, "top": 0.92})
        self.add_widget(self.cb)

    def add_subject(self):
        if self.teachername_entry.text != "" and self.mainbutton.text != "Fənn":
            self.wb = openpyxl.open(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")

            self.mainsheet = self.wb.create_sheet(self.mainbutton.text)
            # self.teachername_entry.text += "."
            self.teachername_entry_text = self.teachername_entry.text + ","

            if not self.cb.active:
                self.mainsheet.cell(row=1, column=1).value = "Müəllim(lər):\n" + self.teachername_entry_text[:-1]
            else:
                self.mainsheet.cell(row=1, column=1).value =\
                    "Müəllim(lər):\n" + self.teachername_entry_text[:self.teachername_entry_text.find(",")] + \
                    "(Sinif Rəhbəridir)" + self.teachername_entry_text[self.teachername_entry_text.find(","):-1]

            self.jurnalsheet = self.wb[f"Sinif {ListSubjects.grade}"]
            for i in range(3, self.jurnalsheet.max_row+1):
                self.mainsheet.cell(row=i-1, column=1).value = self.jurnalsheet.cell(row=i, column=1).value
            self.wb.save(f"sfx/Jurnallar/{ListSubjects.grade}.xlsx")
            ListSubjects.great_list.layout.ids[ListSubjects.id_s[(len(self.wb.sheetnames)-2)% 4]].add_widget(EditableImage(src="sfx/subject.png",
                                                                                                                           dest="FM", id=ListSubjects.grade, subj=self.mainbutton.text))
            ListSubjects.great_list.layout.ids[ListSubjects.id_s[(len(self.wb.sheetnames)-2) % 4]].add_widget(Label(text="{}{}".format(" "*(Window.width//55), self.mainbutton.text),
                                                                                                                    size_hint_y=None, size_hint_x=None))
            self.cb.active = False
            if (len(self.wb.sheetnames)-1) % 4 == 1:
                ListSubjects.great_list.layout.height += Window.height*0.363
                ListSubjects.great_list.height += Window.height*0.363
            ListSubjects.subject_popup.dismiss()


class Calculator(FloatLayout):
    def __init__(self, Lists, grade, subj):
        super(Calculator, self).__init__()
        self.Lists = Lists
        self.subj = subj
        self.grade = grade
        self.dd = DropDown()
        self.dd.max_height = 0.44*Window.height
        self.mainbutton = Button(text="Əməliyyəti Seçin:\n", size_hint=(0.9, 0.2),
                                 pos_hint={"x": 0.05, "top": 0.88},
                                 font_size=Window.width//56,
                                 background_normal="sfx/selection2.png")
        self.add_widget(self.mainbutton)

        for i in ["KSQ əlavə elə", "BSQ əlavə elə",
                  "I və II Yarımil Qiymətləndirmə",
                  "İllik Qiymətləndirmə", "Qiymətləndirmə Sil", "Tarixləri göstər/Tarixi Dəyiş"]:
            self.opt = Button(text=i, font_size=Window.height//46,
                              background_color=(12/255, 63/255, 82/255, 1),
                              size_hint_y=None, height=44)
            self.opt.bind(on_release=lambda opt: self.dd.select(opt.text))
            self.dd.add_widget(self.opt)

        self.mainbutton.bind(on_release=self.dd.open)
        self.dd.bind(on_select=lambda instance, x: self.select(x))

        self.wb = openpyxl.open(f"sfx/Jurnallar/{grade}.xlsx")

    def select(self, x):
        try:
            self.remove_widget(self.date)
            self.remove_widget(self.date_input)
            self.remove_widget(self.test)
        except:
            pass
        try:
            self.remove_widget(self.date)
            self.remove_widget(self.date_input)
            self.remove_widget(self.test)
        except:
            pass
        try:
            self.remove_widget(self.ttbutton)
            self.remove_widget(self.testdate)
            self.remove_widget(self.testdate_input)
            self.remove_widget(self.tarix)
        except:
                pass
        try:
            self.remove_widget(self.ssbutton)
            self.remove_widget(self.sil_test)
        except:
            pass
        try:
            self.remove_widget(self.qqbutton)
            self.remove_widget(self.wo)
            self.remove_widget(self.hesabla)
        except:
            pass
        self.mainbutton.text = self.mainbutton.text[:18] + x
        if "KSQ əlavə elə" in x:
            self.date = Builder.load_string("""
Label:
    text: "KSQ Tarixi:"
    size_hint:(0.4, 0.1)
    pos_hint: {"x": .05, "top": 0.64}
    font_size: self.size[1]//1.9
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [17,17,17,17]""")
            self.add_widget(self.date)

            self.date_input = TextInput(size_hint=(0.45, 0.1),
                                        pos_hint={"x": 0.5, "top": 0.64},
                                        font_size=Window.height//30,
                                        padding=(Window.width//90, 7))
            self.add_widget(self.date_input)

            self.test = Builder.load_string("""
Button:
    text: "Əlavə Elə"
    background_color: (60/255, 58/255, 37/255, 1)
    size_hint: (0.9, 0.1)
    pos_hint: {"x": .05, "top": 0.5}
    font_size: self.size[1]//2.2
    """)
            self.test.on_press = self.preadding_test
            self.add_widget(self.test)

        if "BSQ əlavə elə" in x:
            self.date = Builder.load_string("""
Label:
    text: "BSQ Tarixi:"
    size_hint:(0.4, 0.1)
    pos_hint: {"x": .05, "top": 0.64}
    font_size: self.size[1]//1.9
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [17,17,17,17]""")
            self.add_widget(self.date)

            self.date_input = TextInput(size_hint=(0.45, 0.1),
                                        pos_hint={"x": 0.5, "top": 0.64},
                                        font_size=Window.height//30,
                                        padding=(Window.width//90, 7))
            self.add_widget(self.date_input)

            self.test = Builder.load_string("""
Button:
    text: "Əlavə Elə"
    background_color: (60/255, 58/255, 37/255, 1)
    size_hint: (0.9, 0.1)
    pos_hint: {"x": .05, "top": 0.5}
    font_size: self.size[1]//2.2
    """)
            self.test.on_press = self.preadding_test
            self.add_widget(self.test)

        if "göstər" in x:
            self.tt = DropDown()
            self.tt.max_height = 0.34*Window.height
            self.ttbutton = Button(text="Tarix:\n", size_hint=(0.9, 0.17),
                                   pos_hint={"x": 0.05, "top": 0.65},
                                   font_size=Window.width//56,
                                   background_normal="sfx/selection2.png")
            self.add_widget(self.ttbutton)

            for i in range(2, self.wb[self.subj].max_column+1):
                if "Yarımil" not in self.wb[self.subj].cell(row=1, column=i).value:
                    self.opt = Button(id=str(i), text=self.wb[self.subj].cell(row=1, column=i).value,
                                      font_size=Window.height//46,
                                      background_color=(12/255, 63/255, 82/255, 1),
                                      size_hint_y=None, height=44)
                    self.opt.bind(on_release=lambda opt: self.tt.select(opt))
                    self.tt.add_widget(self.opt)

            self.ttbutton.bind(on_release=self.tt.open)
            self.tt.bind(on_select=lambda instance, x: self.setting_date(x))

        if "Qiymətləndirmə Sil" in x:
            self.ss = DropDown()
            self.ss.max_height = 0.34*Window.height
            self.ssbutton = Button(text="Silmək:\n", size_hint=(0.9, 0.17),
                                   pos_hint={"x": 0.05, "top": 0.65},
                                   font_size=Window.width//56,
                                   background_normal="sfx/selection2.png")
            self.add_widget(self.ssbutton)

            for i in range(2, self.wb[self.subj].max_column+1):
                self.opt = Button(id=str(i), text=self.wb[self.subj].cell(row=1, column=i).value,
                                  font_size=Window.height//46,
                                  background_color=(12/255, 63/255, 82/255, 1),
                                  size_hint_y=None, height=44)
                self.opt.bind(on_release=lambda opt: self.ss.select(opt))
                self.ss.add_widget(self.opt)

            self.ssbutton.bind(on_release=self.ss.open)
            self.ss.bind(on_select=lambda instance, x: self.predelete_tests(x))

            self.sil_test = Builder.load_string("""
Button:
    text: "Sil"
    background_color: (60/255, 58/255, 37/255, 1)
    size_hint: (0.9, 0.1)
    pos_hint: {"x": .05, "top": 0.45}
    font_size: self.size[1]//2.2
    """)
            self.sil_test.on_press = self.delete_tests
            self.add_widget(self.sil_test)

        if "I və II Yarımil Qiymətləndirmə" in x:
            self.qq = DropDown()
            self.qq.max_height = 0.34*Window.height
            self.qqbutton = Button(text="Hesablanan Testlər:\n", size_hint=(0.9, 0.17),
                                   pos_hint={"x": 0.05, "top": 0.65},
                                   font_size=Window.width//56,
                                   background_normal="sfx/selection2.png")
            self.add_widget(self.qqbutton)

            for i in range(2, self.wb[self.subj].max_column+1):
                self.opt = Button(id=str(i), text=self.wb[self.subj].cell(row=1, column=i).value,
                                  font_size=Window.height//46,
                                  background_color=(12/255, 63/255, 82/255, 1),
                                  size_hint_y=None, height=44)
                self.opt.bind(on_release=lambda opt: self.qq.select(opt))
                self.qq.add_widget(self.opt)

            self.qqbutton.bind(on_release=self.qq.open)
            self.qq.bind(on_select=lambda instance, x: self.choose_wo(x.text))

            self.wo = TextInput(text="",  size_hint=(0.9, 0.13),
                                pos_hint={"x": 0.05, "top": 0.45},
                                font_size=Window.width//80)
            self.add_widget(self.wo)
            self.hesabla = Builder.load_string("""
Button:
    text: "Hesabla"
    background_color: (60/255, 58/255, 37/255, 1)
    size_hint: (0.9, 0.1)
    pos_hint: {"x": .05, "top": 0.27}
    font_size: self.size[1]//2.2
    """)
            self.hesabla.on_press = self.yarimhesabla
            self.add_widget(self.hesabla)

        if "İllik Qiymətləndirmə" in x:
            self.mainhesabla = Builder.load_string("""
Button:
    text: "Hesabla"
    background_color: (60/255, 58/255, 37/255, 1)
    size_hint: (0.9, 0.1)
    pos_hint: {"x": .05, "top": 0.57}
    font_size: self.size[1]//2.2
    """)
            self.mainhesabla.on_press = self.illikhesabla
            self.add_widget(self.mainhesabla)

    def illikhesabla(self):
        self.illik_index = []
        for i in range(2, self.wb[self.subj].max_column+1):
            if "Yarım" in self.wb[self.subj].cell(row=1, column=i).value:
                self.illik_index.append(i)
        if len(self.illik_index)==2:
            self.new_max_col = self.wb[self.subj].max_column + 1
            self.wb[self.subj].cell(row=1, column=self.new_max_col).value = "İllik Qiymətləndirmə"
            for j in range(2, self.wb[self.subj].max_row+1):
                self.wb[self.subj].cell(row=j, column=self.new_max_col).value =\
                    (float(self.wb[self.subj].cell(row=j, column=self.illik_index[0]).value) + float(self.wb[self.subj].cell(row=j, column=self.illik_index[1]).value))/2

        self.wb.save(f"sfx/Jurnallar/{self.grade}.xlsx")


    def yarimhesabla(self):
        if self.wo.text:
            self.mainindexes = []
            self.averagesum = []
            self.BSQ = None
            self.wb[self.subj].cell(row=1, column=self.wb[self.subj].max_column+1).value = "Yarımil Qiymətləndirmə"
            for g in self.wo.text.split(",")[:-1]:
                for f in range(2, self.wb[self.subj].max_column+1):
                    if g == self.wb[self.subj].cell(row=1, column=f).value:
                        if "BSQ" in g:
                            self.BSQ = f
                        else:
                            self.mainindexes.append(f)
                        break

            for j in range(2, self.wb[self.subj].max_row+1):
                self.sum = 0
                for f in self.mainindexes:
                    self.sum += int(self.wb[self.subj].cell(row=j, column=f).value)
                self.averagesum.append(self.sum/len(self.mainindexes))

            if self.BSQ is None:
                self.z = 0
                self.new_max_col = self.wb[self.subj].max_column
                for j in range(2, self.wb[self.subj].max_row+1):
                    self.wb[self.subj].cell(row=j, column=self.new_max_col).value = round(self.averagesum[self.z], 1)
                    self.z += 1
                self.wb.save(f"sfx/Jurnallar/{self.grade}.xlsx")
            else:
                self.z = 0
                self.new_max_col = self.wb[self.subj].max_column
                for j in range(2, self.wb[self.subj].max_row+1):
                    self.wb[self.subj].cell(row=j, column=self.new_max_col).value = round((self.averagesum[self.z]*0.4 + int(self.wb[self.subj].cell(row=j, column=self.BSQ).value)*0.6), 1)
                    self.z += 1
                self.wb.save(f"sfx/Jurnallar/{self.grade}.xlsx")

            self.wb = None
            self.wb = openpyxl.open(f"sfx/Jurnallar/{self.grade}.xlsx")

            self.Lists.layout.remove_widget(self.Lists.layout.l2)

            self.Lists.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    size_hint: None, 1
    cols:{0}
    spacing: 10, 21
            """.format(self.wb[self.subj].max_column-1))

            self.Lists.layout.ids["l2"] = self.Lists.layout.l2
            self.Lists.layout.add_widget(self.Lists.layout.l2)

            self.Lists.layout.l2.width = 52.5 * (self.wb[self.subj].max_column-1)
            self.Lists.layout.l2.pos = (.22*Window.width, 0)
            self.Lists.layout.width = Window.width*0.1955 + Window.width*0.043*(self.wb[self.subj].max_column-1)
            gl_ids.clear()

            for j in range(2, self.wb[self.subj].max_row+1):
                for f in range(2, self.wb[self.subj].max_column+1):
                    self.pallet = (12/255,63/255,82/255,1)
                    self.readonly = False
                    if "BSQ" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.pallet = (181/255, 8/255, 53/255, 1)
                    if "Yarımil" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.readonly = True
                        self.pallet = (37/255,124/255,22/255,1)
                    if "İllik Qiymətləndirmə" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.readonly = True
                        self.pallet = (24/255,24/255,24/255,1)
                    if "---" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.readonly = True
                        self.pallet  = (0,0,0,0)
                    if "---" not in self.wb[self.subj].cell(row=j, column=1).value:
                        FinalMarks.txt2 = Builder.load_string(f"""
TextInput:
    readonly: {self.readonly}
    canvas.before:
        Color:
            rgba: {self.pallet}
        Line:
            width: 2
            rectangle: self.x, self.y, self.width, self.height""")
                        FinalMarks.txt2.text = f"{self.wb[self.subj].cell(row=j, column=f).value}"
                        FinalMarks.Lists.layout.ids["l2"].add_widget(FinalMarks.txt2)
                        gl_ids[f"{j}_{f}"] = FinalMarks.txt2
                else:
                    continue

    def choose_wo(self, x):
        self.wo.text += f"{x},"

    def predelete_tests(self, x):
        setattr(self.ssbutton, "text", self.ssbutton.text[:8]+x.text)
        # self.wb.save(f"Jurnallar/{self.grade}.xlsx")
        self.wb = None
        self.wb = openpyxl.open(f"sfx/Jurnallar/{self.grade}.xlsx")
        self.dt = x

    def delete_tests(self):
        if self.ssbutton.text != "Silmək:\n":
            self.dt.text = self.dt.text[:4] + "---)"

            self.wb[self.subj].cell(row=1, column=int(self.dt.id)).value = self.wb[self.subj].cell(row=1, column=int(self.dt.id)).value[:4] + "---)"
            self.wb.save(f"sfx/Jurnallar/{self.grade}.xlsx")

            self.Lists.layout.remove_widget(self.Lists.layout.l2)

            self.Lists.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    size_hint: None, 1
    cols:{0}
    spacing: 10, 21
            """.format(self.wb[self.subj].max_column-1))

            self.Lists.layout.ids["l2"] = self.Lists.layout.l2
            self.Lists.layout.add_widget(self.Lists.layout.l2)

            self.Lists.layout.l2.width = 52.5 * (self.wb[self.subj].max_column-1)
            self.Lists.layout.l2.pos = (.22*Window.width, 0)
            self.Lists.layout.width = Window.width*0.1955 + Window.width*0.043*(self.wb[self.subj].max_column-1)
            gl_ids.clear()

            for j in range(2, self.wb[self.subj].max_row+1):
                for f in range(2, self.wb[self.subj].max_column+1):
                    self.pallet = (12/255,63/255,82/255,1)
                    self.readonly = False
                    if "BSQ" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.pallet = (181/255, 8/255, 53/255, 1)
                    if "Yarımil" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.readonly = True
                        self.pallet = (37/255,124/255,22/255,1)
                    if "İllik Qiymətləndirmə" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.readonly = True
                        self.pallet = (24/255,24/255,24/255,1)
                    if "---" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.readonly = True
                        self.pallet  = (0,0,0,0)
                    if "---" not in self.wb[self.subj].cell(row=j, column=1).value:
                        FinalMarks.txt2 = Builder.load_string(f"""
TextInput:
    readonly: {self.readonly}
    canvas.before:
        Color:
            rgba: {self.pallet}
        Line:
            width: 2
            rectangle: self.x, self.y, self.width, self.height""")
                        FinalMarks.txt2.text = f"{self.wb[self.subj].cell(row=j, column=f).value}"
                        FinalMarks.Lists.layout.ids["l2"].add_widget(FinalMarks.txt2)
                        gl_ids[f"{j}_{f}"] = FinalMarks.txt2
                else:
                    continue

    def setting_date(self, x):
        try:
            self.remove_widget(self.testdate)
            self.remove_widget(self.testdate_input)
            self.remove_widget(self.tarix)
        except:
            pass
        setattr(self.ttbutton,"text", self.ttbutton.text[:7] + x.text)
        self.testdate = Builder.load_string("""
Label:
    text: "Tarixi:"
    size_hint:(0.4, 0.1)
    pos_hint: {"x": .05, "top": 0.44}
    font_size: self.size[1]//1.9
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [17,17,17,17]""")
        self.add_widget(self.testdate)

        self.testdate_input = TextInput(text=x.text[4:-1],
                                        size_hint=(0.45, 0.1),
                                        pos_hint={"x": 0.5, "top": 0.44},
                                        font_size=Window.height//30,
                                        padding=(Window.width//90, 7))
        self.add_widget(self.testdate_input)
        self.d = x

        self.tarix = Builder.load_string("""
Button:
    text: "Tarixi Dəyiş"
    background_color: (60/255, 58/255, 37/255, 1)
    size_hint: (0.9, 0.1)
    pos_hint: {"x": .05, "top": 0.3}
    font_size: self.size[1]//2.2
    """)
        self.tarix.on_press = self.changing_date
        self.add_widget(self.tarix)

    def changing_date(self):
        if self.testdate_input.text != "":
            self.d.text = self.d.text[:3] + f"({self.testdate_input.text})"
            self.wb[self.subj].cell(row=1, column=int(self.d.id)).value = self.d.text
        setattr(self.ttbutton,"text", self.ttbutton.text[:7] + self.d.text)
        self.Lists.layout.remove_widget(self.Lists.layout.l2)

        self.Lists.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    size_hint: None, 1
    cols:{0}
    spacing: 10, 21
            """.format(self.wb[self.subj].max_column-1))
        self.Lists.layout.ids["l2"] = self.Lists.layout.l2
        self.Lists.layout.add_widget(self.Lists.layout.l2)

        self.Lists.layout.l2.width = 52.5 * (self.wb[self.subj].max_column-1)
        self.Lists.layout.l2.pos = (.22*Window.width, 0)
        self.Lists.layout.width = Window.width*0.1955 + Window.width*0.043*(self.wb[self.subj].max_column-1)
        for j in range(2, self.wb[self.subj].max_row+1):
            for f in range(2, self.wb[self.subj].max_column+1):
                self.pallet = (12/255,63/255,82/255,1)
                self.readonly = False

                if "BSQ" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.pallet = (181/255, 8/255, 53/255, 1)
                if "Yarımil" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.readonly = True
                    self.pallet = (37/255,124/255,22/255,1)
                if "İllik Qiymətləndirmə" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.readonly = True
                    self.pallet = (24/255,24/255,24/255,1)
                if "---" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.readonly = True
                    self.pallet  = (0,0,0,0)
                if "---" not in self.wb[self.subj].cell(row=j, column=1).value:
                    self.txt2 = Builder.load_string(f"""
TextInput:
    readonly: {self.readonly}
    canvas.before:
        Color:
            rgba: {self.pallet}
        Line:
            width: 2
            rectangle: self.x, self.y, self.width, self.height""")
                    self.txt2.text = f"{self.wb[self.subj].cell(row=j, column=f).value}"
                    self.Lists.layout.ids["l2"].add_widget(self.txt2)
                    gl_ids[f"{j}_{f}"] = self.txt2

            else:
                continue

        self.wb.save(f"sfx/Jurnallar/{self.grade}.xlsx")

    def preadding_test(self):
        if self.date_input.text != "":
            self.adding_test()

    def adding_test(self):
        self.new_max_col = self.wb[self.subj].max_column+1
        if "KSQ" in self.mainbutton.text:
            self.wb[self.subj].cell(row=1, column=self.new_max_col).value = f"KSQ({self.date_input.text})"
            gl_ids[f"1_{self.new_max_col}"] = TextInput(text=f"KSQ({self.date_input.text})")
        else:
            self.wb[self.subj].cell(row=1, column=self.new_max_col).value = f"BSQ({self.date_input.text})"
            gl_ids[f"1_{self.new_max_col}"] = TextInput(text=f"BSQ({self.date_input.text})")

        for g in gl_ids.keys():
            if str(gl_ids[g].text) != str(self.wb[self.subj].cell(row=int(g.split("_")[0]), column=int(g.split("_")[1])).value):
                self.wb[self.subj].cell(row=int(g.split("_")[0]), column=int(g.split("_")[1])).value = str(gl_ids[g].text)

        for j in range(2, self.wb[self.subj].max_row+1):
            self.wb[self.subj].cell(row=j, column=self.new_max_col).value = "0"
        self.wb.save(f"sfx/Jurnallar/{self.grade}.xlsx")

        self.Lists.layout.remove_widget(self.Lists.layout.l2)
        self.Lists.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    size_hint: None, 1
    cols:{0}
    spacing: 10, 21
        """.format(self.wb[self.subj].max_column-1))

        self.Lists.layout.ids["l2"] = self.Lists.layout.l2
        self.Lists.layout.add_widget(self.Lists.layout.l2)

        self.Lists.layout.l2.width = 52.5 * (self.wb[self.subj].max_column-1)
        self.Lists.layout.l2.pos = (.22*Window.width, 0)
        self.Lists.layout.width = Window.width*0.1955 + Window.width*0.043*(self.wb[self.subj].max_column-1)

        for j in range(2, self.wb[self.subj].max_row+1):
            for f in range(2, self.wb[self.subj].max_column+1):
                self.pallet = (12/255,63/255,82/255,1)
                self.readonly = False
                if "BSQ" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.pallet = (181/255, 8/255, 53/255, 1)
                if "---" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.readonly = True
                    self.pallet = (0,0,0,0)
                if "---" not in self.wb[self.subj].cell(row=j, column=1).value:
                    self.txt2 = Builder.load_string(f"""
TextInput:
    readonly: {self.readonly}
    canvas.before:
        Color:
            rgba: {self.pallet}
        Line:
            width: 2
            rectangle: self.x, self.y, self.width, self.height""")
                    self.txt2.text = f"{self.wb[self.subj].cell(row=j, column=f).value}"
                    self.Lists.layout.ids["l2"].add_widget(self.txt2)
                    gl_ids[f"{j}_{f}"] = self.txt2
                else:
                    continue


class EditableImage(Image, ButtonBehavior, FloatLayout):
    def __init__(self, id="GG", src="sfx/dir.png", dest="x", subj="None"):
        super(EditableImage, self).__init__()
        self.id = id
        self.subj = subj
        self.source = src
        self.size_hint_y = None
        self.size_hint_x = None
        self.size = (Window.width*0.15, Window.height*0.2)
        if dest == "x":
            self.on_press = self.print_id
        else:
            self.on_press = self.print_id2

    def print_id2(self):
        try:
            sm.remove_widget(sm.get_screen("FM"))
        except:
            pass
        self.show = Calculator(FinalMarks.Lists, self.id, self.subj)
        FinalMarks.prefinalpopup.content = self.show
        sm.add_widget(FinalMarks(name="FM", grade=self.id, subj=self.subj))
        self.list_superlist()
        sm.transition.direction = "left"
        sm.current = "FM"

    def print_id(self):
        global sm
        ListSubjects.logo.text = self.id
        ListSubjects.grade = self.id
        self.list_greatlist()
        sm.transition.direction = "left"
        sm.current = "LS"

    def list_superlist(self):
        self.wb = openpyxl.open(f"sfx/Jurnallar/{self.id}.xlsx")
        FinalMarks.tn.text = self.wb[self.subj].cell(row=1,column=1).value[14:]

        for i in range(self.wb[self.subj].max_row-1):
            if "---" not in self.wb[self.subj].cell(row=i+2, column=1).value:
                FinalMarks.txt = TextInput(id=f"{i}", text=self.wb[self.subj].cell(row=i+2, column=1).value)
                gl_ids[f"{i+2}_1"] = FinalMarks.txt
                FinalMarks.Lists.layout.ids["l1"].add_widget(FinalMarks.txt)

        for j in range(2, self.wb[self.subj].max_row+1):
            for f in range(2, self.wb[self.subj].max_column+1):
                self.readonly = False
                self.pallet = (12/255,63/255,82/255,1)
                if "BSQ" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.pallet = (181/255, 8/255, 53/255, 1)
                if "Yarımil" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.readonly = True
                    self.pallet = (37/255,124/255,22/255,1)
                if "İllik Qiymətləndirmə" in self.wb[self.subj].cell(row=1, column=f).value:
                        self.readonly = True
                        self.pallet = (24/255,24/255,24/255,1)
                if "---" in self.wb[self.subj].cell(row=1, column=f).value:
                    self.readonly = True
                    self.pallet = (0,0,0,0)
                if "---" not in self.wb[self.subj].cell(row=j, column=1).value:
                    FinalMarks.txt2 = Builder.load_string(f"""
TextInput:
    readonly: {self.readonly}
    canvas.before:
        Color:
            rgba: {self.pallet}
        Line:
            width: 2
            rectangle: self.x, self.y, self.width, self.height""")
                    FinalMarks.txt2.text = f"{self.wb[self.subj].cell(row=j, column=f).value}"
                    FinalMarks.Lists.layout.ids["l2"].add_widget(FinalMarks.txt2)
                    gl_ids[f"{j}_{f}"] = FinalMarks.txt2
                    # self.Lists.layout.ids[f"{i}"] = self.txt
                else:
                    continue

    def list_greatlist(self):
        self.wb = openpyxl.open(f"sfx/Jurnallar/{self.id}.xlsx")
        self.collator = icu.Collator.createInstance(icu.Locale('az_Latn_AZ'))
        self.sheetnames = sorted(self.wb.sheetnames[1:], key=self.collator.getSortKey)
        ListSubjects.great_list.layout.height = Window.height*0.37*math.ceil(len(self.sheetnames)/4)
        ListSubjects.great_list.height = Window.height*0.37*(math.ceil(len(self.sheetnames)/4))
        for i in range(0,  len(self.sheetnames)):
            ListSubjects.great_list.layout.ids[ListSubjects.id_s[i % 4]].add_widget(EditableImage(src="sfx/subject.png",
                                                                                                  dest="FM", id=ListSubjects.grade, subj=self.sheetnames[i]))
            ListSubjects.great_list.layout.ids[ListSubjects.id_s[i % 4]].add_widget(Label(text="{}{}".format(" "*(Window.width//55), self.sheetnames[i]),
                                                                                          size_hint_y=None, size_hint_x=None))
        self.wb.close()


class FinalMarks(Screen, FloatLayout):
    Lists = Builder.load_string("""
ScrollView:
    do_scroll_x: True
    do_scroll_y: True
    scroll_type: ['bars']
    bar_color: (60/255, 58/255, 37/255, 1)
    bar_inactive_color: (60/255, 58/255, 37/255, 0.78)

    size_hint: 0.85, 0.6
    pos_hint: {"x": 0.05, "top": 0.8}
    layout: layout
    bar_width: 9

    FloatLayout:
        id: layout
        l1:l1
        l2:l2
        pos: 0, 0
        
        size_hint_y: None
        size_hint_x: None

        GridLayout:
            id: l1    
            cols: 1
            size_hint: None, 1
            spacing: 0, 21

        GridLayout:
            id: l2 
            cols: 0
            size_hint: None, 1
            spacing: 10, 21
        """)
    tn = Builder.load_string("""
TextInput:
    id: tn
    size_hint: 0.13, 0.07
    pos_hint: {"x":0.115, "top":0.9}
    """)

    prefinalpopup = Popup(title="Qiymətləndirmə Əməliyyətlar.",
                                title_size=Window.size[0]*0.025,
                                title_color=(1, 43/255, 18/255, 1),
                                size_hint=(0.3, 0.73))

    def __init__(self, grade, subj, **kwargs):
        super(FinalMarks, self).__init__(**kwargs)
        self.subj = subj
        self.grade = grade
        self.exit = EditableImage()
        self.exit.source = "sfx/close.png"
        self.exit.size_hint = (0.06, 0.12)
        self.exit.pos_hint = {"x": 0.92, "top": 0.66}
        self.exit.on_press = sys.exit
        self.add_widget(self.exit)

        self.finish = EditableImage()
        self.finish.source = "sfx/tick.png"
        self.finish.size_hint = (0.06, 0.12)
        self.finish.pos_hint = {"x": 0.92, "top": 0.9}
        self.finish.on_press = self.finished
        self.add_widget(self.finish)

        self.back = EditableImage()
        self.back.source = "sfx/back.png"
        self.back.size_hint = (0.06, 0.12)
        self.back.pos_hint = {"x": 0.92, "top": 0.78}
        self.back.on_press = self.back_tolistsubjects
        self.add_widget(self.back)

        self.add = EditableImage()
        self.add.on_press = self.prefinalpopup.open
        self.add.source = "sfx/plus.png"
        self.add.size_hint = (0.1, 0.2)
        self.add.pos_hint = {"x": 0.44, "top": .1978}
        self.add_widget(self.add)
        self.wb = openpyxl.open(f"sfx/Jurnallar/{self.grade}.xlsx")

        if self.Lists.layout.l2.cols == 0:
            self.Lists.layout.l2.cols = self.wb[self.subj].max_column-1

        self.Lists.layout.l2.width = 52.5 * (self.wb[self.subj].max_column-1)
        self.Lists.layout.l1.width = Window.width*0.1955
        self.Lists.layout.l2.pos = (.22*Window.width, 0)
        self.Lists.layout.width = Window.width*0.1955 + Window.width*0.043*(self.wb[self.subj].max_column-1)

        self.add_widget(Builder.load_string("""
Label:
    text:"KSQ bal"
    size_hint:(0.045, 0.07)
    pos_hint:{"x": 0.395, "top": 0.98}
    canvas.before:
        Color:
            rgba: (12/255,63/255,82/255,1)
        Rectangle:
            size: self.size
            pos: self.pos"""))
        self.add_widget(Builder.load_string("""
Label:
    text: "BSQ bal"
    size_hint:(0.045, 0.07)
    pos_hint:{"x": 0.45, "top": 0.98}
    canvas.before:
        Color:
            rgba: (181/255, 8/255, 53/255, 1)
        Rectangle:
            size: self.size
            pos: self.pos"""))
        self.add_widget(Builder.load_string("""
Label:
    text:"Yarımil"
    size_hint:(0.045, 0.07)
    pos_hint:{"x": 0.505, "top": 0.98}
    canvas.before:
        Color:
            rgba: (37/255,124/255,22/255,1)
        Rectangle:
            size: self.size
            pos: self.pos"""))
        self.add_widget(Builder.load_string("""
Label:
    text:"İllik"
    size_hint:(0.045, 0.07)
    pos_hint:{"x": 0.56, "top": 0.98}
    canvas.before:
        Color:
            rgba: (24/255,24/255,24/255,1)
        Rectangle:
            size: self.size
            pos: self.pos"""))

        self.Lists.layout.height = 57*(self.wb[self.subj].max_row-1)
        # self.Lists.layout.l2.width = 52.5 * (self.wb[self.subj].max_column-1)
        self.add_widget(self.Lists)
        self.add_widget(self.tn)
        self.Lists.layout.ids = {"l1": self.Lists.layout.l1,
                                 "l2": self.Lists.layout.l2}

    def back_tolistsubjects(self):
        self.Lists.layout.remove_widget(self.Lists.layout.l1)
        self.Lists.layout.remove_widget(self.Lists.layout.l2)
        self.remove_widget(self.Lists)
        self.remove_widget(self.tn)
        self.Lists.layout.l1 = Builder.load_string("""
GridLayout:
    id: l1    
    cols: 1
    size_hint: None, 1
    spacing: 0, 21
        """)
        self.Lists.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2 
    cols: 0
    size_hint: None, 1
    spacing: 10, 21""")
        self.Lists.layout.ids = {"l1": self.Lists.layout.l1,
                                 "l2": self.Lists.layout.l2}
        self.Lists.layout.add_widget(self.Lists.layout.l1)
        self.Lists.layout.add_widget(self.Lists.layout.l2)
        sm.transition.direction = "right"
        sm.current = "LS"

    def finished(self):
        for g in gl_ids.keys():
            if str(gl_ids[g].text) != str(self.wb[self.subj].cell(row=int(g.split("_")[0]), column=int(g.split("_")[1])).value):
                self.wb[self.subj].cell(row=int(g.split("_")[0]), column=int(g.split("_")[1])).value = str(gl_ids[g].text)
        self.wb.save(f"sfx/Jurnallar/{self.grade}.xlsx")


class Yarat(FloatLayout):
    def __init__(self, num, grade, pow2, sector):
        super(Yarat, self).__init__()
        self.num = num
        self.grade = grade
        self.pow = pow2
        self.sector = sector
        self.Lists = Builder.load_string("""
ScrollView:
    scroll_type: ['bars']
    bar_color: (60/255, 58/255, 37/255, 1)
    bar_inactive_color: (60/255, 58/255, 37/255, 0.78)
    layout:layout
    size_hint: 0.6, 0.5
    pos_hint: {}
    bar_width: 9
    GridLayout:
        height: {}*58
        id: layout
        cols: 1
        spacing: 0, root.height//13
        size_hint_y: None""".format({"x": 0.2, "top": 0.77}, num))
        self.add_widget(self.Lists)
        for i in range(num):
            self.txt = TextInput(id=f"{i}", multiline=False)
            self.Lists.layout.add_widget(self.txt)
            self.Lists.layout.ids[f"{i}"] = self.txt

    def listing(self):
        with codecs.open("ls.txt", "r", "utf-8") as file:
            self.info = file.read().split("|")
            self.jurnalamount = int(self.info[2])

        if self.pow != "":
            self.new_grade = f"{self.grade}_{self.pow}_{self.sector},"
        else:
            self.new_grade = f"{self.grade}_{self.sector},"
        self.info[3] += self.new_grade

        if (self.jurnalamount-1) % 4 == 3:
            ListJournals.great_list.layout.height += Window.height*0.37
        self.jurnalamount += 1

        AddJournal().listjournal_window.dismiss()
        ListJournals.great_list.layout.ids[ListJournals.id_s[(self.jurnalamount-1) % 4]].add_widget(EditableImage(id=self.new_grade[:-1]))
        ListJournals.great_list.layout.ids[ListJournals.id_s[(self.jurnalamount-1) % 4]].add_widget(Label(
            text="{}{}".format(" "*(Window.width//55), self.new_grade[:-1]), size_hint_y=None, size_hint_x=None,))

        self.info[2] = self.jurnalamount
        with codecs.open("ls.txt", "w", "utf-8") as file:
            for i in self.info[:-1]:
                print(i, end="|", file=file)
        self.wb = xlsxwriter.Workbook(f"sfx/Jurnallar/{self.new_grade[:-1]}.xlsx")
        self.wb.close()
        self.wb = openpyxl.open(f"sfx/Jurnallar/{self.new_grade[:-1]}.xlsx")

        self.mainsheet = self.wb.create_sheet(title=f"Sinif {self.new_grade[:-1]}")
        self.wb.remove_sheet(self.wb["Sheet1"])
        self.mainsheet.cell(row=1, column=1).value = f"{self.info[0]}  \nDirektor:{self.info[1]}"
        self.mainsheet.cell(row=2, column=1).value = f"Şagirdlər ASA({self.num} nəfər)"
        self.class_list = []
        for i in range(self.num):
            self.class_list.append(self.Lists.layout.ids[f"{i}"].text.lower())
        self.collator = icu.Collator.createInstance(icu.Locale('az_Latn_AZ'))
        self.class_list = sorted(self.class_list, key=self.collator.getSortKey)
        for i in range(len(self.class_list)):
            self.mainsheet.cell(row=i+3, column=1).value = string.capwords(self.class_list[i])

        self.wb.save(f"sfx/Jurnallar/{self.new_grade[:-1]}.xlsx")


class ListSubjects(Screen, FloatLayout):
    logo = Builder.load_string(logotype)
    great_list = Builder.load_string('''
ScrollView:
    scroll_type: ['bars']
    bar_color: (60/255, 58/255, 37/255, 1)
    bar_inactive_color: (60/255, 58/255, 37/255, 0.78)
    size_hint: 0.7, 0.45
    pos_hint: {"x": 0.15, "top": 0.8}
    layout: layout
    bar_width: 9

    GridLayout:
        id: layout
        l1:l1
        l2:l2
        l3:l3
        l4:l4
        size_hint_y: None
        cols: 4
        
        spacing: 0,0
        GridLayout:
            id: l1
            cols:1
        GridLayout:
            id: l2
            cols:1      
        GridLayout:
            id: l3
            cols:1
        GridLayout:
            id: l4
            cols:1     ''')
    grade = "None"
    id_s = ["l1", "l2", "l3", "l4"]
    subject_popup = Popup(title="Yeni Jurnal səhifə yaradın.",
                          title_size=Window.size[0]*0.025,
                          title_color=(1, 43/255, 18/255, 1),
                          size_hint=(0.6, 0.55))

    config_popup = Popup(title="Prametrlər Dəyişmə",
                         title_size=Window.size[0]*0.025,
                         title_color=(1, 43/255, 18/255, 1),
                         size_hint=(0.96, 0.96))

    great_list.layout.ids = {"l1": great_list.layout.l1,
                             "l2": great_list.layout.l2,
                             "l3": great_list.layout.l3,
                             "l4": great_list.layout.l4}

    def __init__(self, **kwargs):
        super(ListSubjects, self).__init__(**kwargs)

        self.show4 = Subject()
        self.subject_popup.content = self.show4

        # self.show5 = Parameters()
        # self.config_popup.content = self.show5

        self.add = EditableImage()
        self.add.on_press = self.subject_popup.open
        self.add.source = "sfx/plus.png"
        self.add.size_hint = (0.12, 0.24)
        self.add.pos_hint = {"x": 0.44, "top": .33}
        self.add_widget(self.add)

        self.add_widget(self.logo)

        self.exit = EditableImage()
        self.exit.source = "sfx/close.png"
        self.exit.size_hint = (0.06, 0.12)
        self.exit.pos_hint = {"x": 0.88, "top": 0.66}
        self.exit.on_press = sys.exit
        self.add_widget(self.exit)

        self.back = EditableImage()
        self.back.source = "sfx/back.png"
        self.back.size_hint = (0.06, 0.12)
        self.back.pos_hint = {"x": 0.88, "top": 0.78}
        self.back.on_press = self.back_tolistjournals
        self.add_widget(self.back)

        self.confiq = EditableImage()
        self.confiq.source = "sfx/confiq.png"
        self.confiq.size_hint = (0.06, 0.12)
        self.confiq.pos_hint = {"x": 0.88, "top": 0.9}
        self.confiq.on_press = self.config_popup_open
        self.add_widget(self.confiq)

        self.add_widget(self.great_list)

    def config_popup_open(self):
        self.show5 = Parameters(self.grade)
        self.config_popup.content = self.show5
        self.config_popup.open()

    def back_tolistjournals(self):
        sm.transition.direction = "right"
        sm.current = "LJ"
        self.great_list.layout.remove_widget(self.great_list.layout.l1)
        self.great_list.layout.remove_widget(self.great_list.layout.l2)
        self.great_list.layout.remove_widget(self.great_list.layout.l3)
        self.great_list.layout.remove_widget(self.great_list.layout.l4)

        self.great_list.layout.l1 = Builder.load_string("""
GridLayout:
    id: l1
    cols:1""")
        self.great_list.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    cols:1""")
        self.great_list.layout.l3 = Builder.load_string("""
GridLayout:
    id: l3
    cols:1""")
        self.great_list.layout.l4 = Builder.load_string("""
GridLayout:
    id: l4
    cols:1""")
        self.great_list.layout.add_widget(self.great_list.layout.l1)
        self.great_list.layout.add_widget(self.great_list.layout.l2)
        self.great_list.layout.add_widget(self.great_list.layout.l3)
        self.great_list.layout.add_widget(self.great_list.layout.l4)

        self.great_list.layout.ids = {"l1": self.great_list.layout.l1,
                                      "l2": self.great_list.layout.l2,
                                      "l3": self.great_list.layout.l3,
                                      "l4": self.great_list.layout.l4}


class AddJournal(FloatLayout):
    listjournal_window = Popup(title_size=Window.size[0]*0.025,
                               title_color=(1, 43/255, 18/255, 1),
                               size_hint=(0.9, 0.9))

    def __init__(self):
        super(AddJournal, self).__init__()
        self.dd = DropDown()
        self.AZ = Button(text='AZ', font_size=Window.height//45,
                         background_color=(12/255, 63/255, 82/255, 1),
                         size_hint_y=None, height=44)
        self.RU = Button(text='RU', font_size=Window.height//45,
                         background_color=(12/255, 63/255, 82/255, 1),
                         size_hint_y=None, height=44)
        self.AZ.bind(on_release=lambda az: self.dd.select(az.text))
        self.RU.bind(on_release=lambda ru: self.dd.select(ru.text))
        self.dd.add_widget(self.AZ)
        self.dd.add_widget(self.RU)
        self.mainbutton = Button(text='Bölmə', size_hint=(None, None),
                                 pos_hint={"x": 0.7, "top": 0.77},
                                 font_size=Window.width//30,
                                 background_normal="sfx/selection.png")

        self.mainbutton.bind(on_release=self.dd.open)
        self.dd.bind(on_select=lambda instance, x: setattr(self.mainbutton,
                                                           'text', x))
        self.add_widget(self.mainbutton)

    def create(self, num, grade, pow2):
        if num != "" and grade != "" and self.mainbutton.text != "Bölmə":
            try:
                if not pow2:
                    self.wb = openpyxl.open(f"sfx/Jurnallar/{grade}_{self.mainbutton.text}.xlsx")
                    Digit().show_popup(" "*(Window.width//150) + f"Diqqət! {grade}_{self.mainbutton.text} Jurnalı  yaradıla bilməz, çünki artıq sistemdə mövcuddur!")
                else:
                    self.wb = openpyxl.open(f"sfx/Jurnallar/{grade}_{pow2}_{self.mainbutton.text}.xlsx")
                    Digit().show_popup(" "*(Window.width//150) + f"Diqqət! {grade}_{self.mainbutton.text}_{pow2} Jurnalı yaradıla bilməz, çünki artıq sistemdə mövcuddur!")
                Digit().show.next.text = "Qayıt"
                Digit().show.next.on_press = Digit().popup_window.dismiss

            except:
                self.listjournal_window.content = Yarat(int(num), grade, pow2, self.mainbutton.text)
                self.listjournal_window.title = "Jurnalı yarat({} nəfərdən ibarətdir)".format(num)
                for i in self.ids.keys():
                    if i != "yarat":
                        self.ids[i].text = ""
                self.mainbutton.text = "Bölmə"
                ListJournals.addjournal_window.dismiss()
                self.listjournal_window.open()


class ListJournals(Screen, FloatLayout):
    show2 = AddJournal()
    addjournal_window = Popup(content=show2, title="Jurnalı yarat",
                              title_size=Window.size[0]*0.025,
                              title_color=(1, 43/255, 18/255, 1),
                              size_hint=(0.6, 0.55))
    i = 0
    great_list = Builder.load_string('''
ScrollView:
    scroll_type: ['bars']
    bar_color: (60/255, 58/255, 37/255, 1)
    bar_inactive_color: (60/255, 58/255, 37/255, 0.78)
    size_hint: 0.7, 0.45
    pos_hint: {"x": 0.15, "top": 0.8}
    layout: layout
    bar_width: 9
    GridLayout:
        id: layout
        l1:l1
        l2:l2
        l3:l3
        l4:l4
        size_hint_y: None
        cols: 4
        
        spacing: 0,0
        GridLayout:
            id: l1
            cols:1
        GridLayout:
            id: l2
            cols:1      
        GridLayout:
            id: l3
            cols:1
        GridLayout:
            id: l4
            cols:1     ''')
    id_s = ["l1", "l2", "l3", "l4"]

    def __init__(self, **kwargs):
        super(ListJournals, self).__init__(**kwargs)
        self.logo = Builder.load_string(logotype)
        self.logo.text = "Jurnalların Siyahısı"
        self.add_widget(self.logo)

        self.add = EditableImage()
        self.add.on_press = self.show_addjounral
        self.add.source = "sfx/plus.png"
        self.add.size_hint = (0.12, 0.24)
        self.add.pos_hint = {"x": 0.44, "top": .33}
        self.add_widget(self.add)

        self.confiq = EditableImage()
        self.confiq.source = "sfx/confiq.png"
        self.confiq.size_hint = (0.06, 0.12)
        self.confiq.pos_hint = {"x": 0.88, "top": 0.9}
        self.confiq.on_press = self.change_confiq
        self.add_widget(self.confiq)

        self.exit = EditableImage()
        self.exit.source = "sfx/folder.png"
        self.exit.size_hint = (0.06, 0.12)
        self.exit.pos_hint = {"x": 0.88, "top": 0.78}
        self.exit.on_press = self.open_mainfolder
        self.add_widget(self.exit)

        self.exit = EditableImage()
        self.exit.source = "sfx/close.png"
        self.exit.size_hint = (0.06, 0.12)
        self.exit.pos_hint = {"x": 0.88, "top":0.66}
        self.exit.on_press = sys.exit
        self.add_widget(self.exit)

        with codecs.open("ls.txt", "r", "utf-8") as file:
            self.info = file.read().split("|")
        self.jurnalamount = int(self.info[2])
        self.jurnallist = self.info[3].split(",")

        self.great_list.layout.height = Window.height*0.37*math.ceil(self.jurnalamount/4)

        self.great_list.height = Window.height*0.37*(math.ceil(self.jurnalamount/4))
        self.great_list.layout.ids = {"l1": self.great_list.layout.l1,
                                      "l2": self.great_list.layout.l2,
                                      "l3": self.great_list.layout.l3,
                                      "l4": self.great_list.layout.l4}

        while self.i < self.jurnalamount:
            self.great_list.layout.ids[self.id_s[self.i % 4]].add_widget(EditableImage(id=self.jurnallist[self.i]))
            self.great_list.layout.ids[self.id_s[self.i % 4]].add_widget(Label(
                text="{}{}".format(" "*(Window.width//55), self.jurnallist[self.i]), size_hint_y=None, size_hint_x=None,
                ))

            self.i += 1
        self.add_widget(self.great_list)

    def show_addjounral(self):
        self.addjournal_window.open()

    def change_confiq(self):
        global sm
        sm.transition.direction = "right"
        sm.current = "LM"

    def open_mainfolder(self):
        os.startfile("sfx/Jurnallar")


class P(FloatLayout):
    pass


class ListNames(Screen, FloatLayout):
    def __init__(self, **kwargs):
        super(ListNames, self).__init__(**kwargs)
        self.logo = Builder.load_string(logotype)
        self.add_widget(self.logo)

        self.exit = EditableImage()
        self.exit.source = "sfx/close.png"
        self.exit.size_hint = (0.06, 0.12)
        self.exit.pos_hint = {"x": 0.88, "top": 0.9}
        self.exit.on_press = sys.exit
        self.add_widget(self.exit)

        with codecs.open("ls.txt", "r", "utf-8") as file:
            self.info = file.read().split("|")

        self.schoolname = Builder.load_string(lb)
        self.schoolname.pos_hint = {"top": 0.75, "x": 0.2}
        self.add_widget(self.schoolname)
        self.schoolname_entry = TextInput(text=self.info[0], size_hint=(0.35, 0.1),
                                          pos_hint={"x": 0.45, "top": 0.75},
                                          multiline=False, padding=(Window.width//27,
                                                                   Window.height//58),
                                          font_size=Window.height//39)
        self.add_widget(self.schoolname_entry)

        self.directorname = Builder.load_string(lb)
        self.directorname.text = "Direktorun ASA"
        self.directorname.pos_hint = {"top": 0.6, "x": 0.2}
        self.add_widget(self.directorname)
        self.directorname_entry = TextInput(text=self.info[1], size_hint=(0.35, 0.1),
                                            pos_hint={"x": 0.45, "top": 0.6},
                                            multiline=False, padding=(Window.width//27,
                                                                     Window.height//58),
                                            font_size=Window.height//35)
        self.add_widget(self.directorname_entry)

        self.submitnames = Builder.load_string("""
Button:
    text: "Qeyd Elə"
    color: (112/255,8/255,5/255,1)
    bold: True
    font_size: self.size[0]//10
    size_hint: 0.3, 0.1
    pos_hint: {"x":0.35, "top":0.4}
    background_color: (0,0,0,0)
    canvas.before:
        Color:
            rgba: (112/255,8/255,5/255,1)
        Line:
            rounded_rectangle: (self.pos[0],self.pos[1], self.size[0], self.size[1],12)
            width: 3
""")
        self.submitnames.on_press = self.save_names
        self.add_widget(self.submitnames)

    def save_names(self):
        with codecs.open("ls.txt", "r+", "utf-8") as file:
            self.info = file.read().split("|")
            self.info[0] = self.schoolname_entry.text
            self.info[1] = self.directorname_entry.text
            self.info[4] = "NO"
        with codecs.open("ls.txt", "w", "utf-8") as file:
            for i in self.info[:-1]:
                print(i, file=file, end="|")
        Digit().show.next.on_press = self.close_popup
        Digit().show_popup(f"""
{self.schoolname_entry.text}, {self.directorname_entry.text} parametrlər qeyd olundu.
Gələcəkdə parametrləri dəyişmək mümkündür""")

    def close_popup(self):
        global sm
        Digit().popup_window.dismiss()
        sm.transition.direction = "left"
        sm.current = "LJ"


class Digit(App):
    global sm
    Window.size = (GetSystemMetrics(0), GetSystemMetrics(1))
    Window.fullscreen = 'auto'
    Window.clearcolor = (210/255, 140/255, 8/255, 1)
    with codecs.open("ls.txt", "r", "utf-8") as file:
        info = file.read().split("|")
    if int(info[2]) + 1 != len(info[3].split(",")):
        with codecs.open("ls.txt", "w", "utf-8") as file:
            info[2] = len(info[3].split(",")) - 1
            for i in info[:-1]:
                print(i, end="|", file=file)
    if info[4] == "NO":
        sm.add_widget(ListJournals(name="LJ"))
        sm.add_widget(ListNames(name="LM"))
    else:
        sm.add_widget(ListNames(name="LM"))
        sm.add_widget(ListJournals(name="LJ"))
    sm.add_widget(ListSubjects(name="LS"))

    show = P()
    popup_window = Popup(content=show, title="Diqqət!",
                         title_size=Window.size[0]*0.025,
                         title_color=(1, 43/255, 18/255, 1),
                         size_hint=(0.6, 0.55))

    def build(self):
        try:
            os.mkdir("sfx/Jurnallar")
        except:
            pass
        Config.set('kivy', 'exit_on_escape', '0')
        return sm

    def show_popup(self, comment):
        self.show.com.text = comment
        self.popup_window.open()

    def delete_journal(self, x):
        with codecs.open("ls.txt", "r", "utf-8") as file:
            self.info = file.read().split("|")
        self.z = self.info[3].split(",")[:-1]
        self.z.remove(x)
        self.info[2] = len(self.z)
        self.info[3] = ""
        os.remove(f'sfx/Jurnallar/{x}.xlsx')
        for x in self.z:
            self.info[3] += f"{x},"
        with codecs.open("ls.txt", "w", "utf-8") as file:
            for i in self.info[:-1]:
                print(i, end="|", file=file)
            file.close()
        ListJournals.great_list.layout.remove_widget(ListJournals.great_list.layout.l1)
        ListJournals.great_list.layout.remove_widget(ListJournals.great_list.layout.l2)
        ListJournals.great_list.layout.remove_widget(ListJournals.great_list.layout.l3)
        ListJournals.great_list.layout.remove_widget(ListJournals.great_list.layout.l4)

        ListJournals.great_list.layout.l1 = Builder.load_string("""
GridLayout:
    id: l1
    cols:1""")
        ListJournals.great_list.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    cols:1""")
        ListJournals.great_list.layout.l3 = Builder.load_string("""
GridLayout:
    id: l3
    cols:1""")
        ListJournals.great_list.layout.l4 = Builder.load_string("""
GridLayout:
    id: l4
    cols:1""")

        ListJournals.great_list.layout.add_widget(ListJournals.great_list.layout.l1)
        ListJournals.great_list.layout.add_widget(ListJournals.great_list.layout.l2)
        ListJournals.great_list.layout.add_widget(ListJournals.great_list.layout.l3)
        ListJournals.great_list.layout.add_widget(ListJournals.great_list.layout.l4)

        ListJournals.great_list.layout.ids = {"l1": ListJournals.great_list.layout.l1,
                                              "l2": ListJournals.great_list.layout.l2,
                                              "l3": ListJournals.great_list.layout.l3,
                                              "l4": ListJournals.great_list.layout.l4}

        for i in range(0, int(self.info[2])):
            print(self.info[3])
            ListJournals.great_list.layout.ids[ListJournals.id_s[i % 4]].add_widget(EditableImage(id=self.z[i]))
            ListJournals.great_list.layout.ids[ListJournals.id_s[i % 4]].add_widget(Label(
                text="{}{}".format(" "*(Window.width//55), self.z[i]), size_hint_y=None, size_hint_x=None,
            ))
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l1)
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l2)
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l3)
        ListSubjects.great_list.layout.remove_widget(ListSubjects.great_list.layout.l4)

        ListSubjects.great_list.layout.l1 = Builder.load_string("""
GridLayout:
    id: l1
    cols:1""")
        ListSubjects.great_list.layout.l2 = Builder.load_string("""
GridLayout:
    id: l2
    cols:1""")
        ListSubjects.great_list.layout.l3 = Builder.load_string("""
GridLayout:
    id: l3
    cols:1""")
        ListSubjects.great_list.layout.l4 = Builder.load_string("""
GridLayout:
    id: l4
    cols:1""")
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l1)
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l2)
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l3)
        ListSubjects.great_list.layout.add_widget(ListSubjects.great_list.layout.l4)

        ListSubjects.great_list.layout.ids = {"l1": ListSubjects.great_list.layout.l1,
                                              "l2": ListSubjects.great_list.layout.l2,
                                              "l3": ListSubjects.great_list.layout.l3,
                                              "l4": ListSubjects.great_list.layout.l4}

        ListSubjects.config_popup.dismiss()
        sm.transition.direction = "right"
        sm.current = "LJ"


if __name__ == "__main__":
    Digit().run()