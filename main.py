import pathlib
import os
import time
import concurrent.futures
import subprocess
import win32com.client
import signal

from kivy.metrics import dp
from kivymd.app import MDApp
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.lang import Builder
from kivy.clock import mainthread
from kivymd.uix.dialog import MDDialog
from kivymd.uix.card import MDCard
from kivy.core.window import Window
from kivymd.uix.snackbar import Snackbar
from kivymd.uix.behaviors import HoverBehavior
from kivymd.uix.button import MDRaisedButton, MDFloatingActionButton, MDRectangleFlatButton
from kivy.properties import ObjectProperty, StringProperty

app_path = os.getcwd()
Builder.load_file('design.kv')


class RootPopup(Popup):
    drop_id = ObjectProperty(None)
    filePath = StringProperty('')
    folderpath = []

    def __init__(self, **kwargs):
        super(RootPopup, self).__init__(**kwargs)
        del self.folderpath[:]
        Window.bind(on_dropfile=self._on_file_drop)

    def _on_file_drop(self, window, file_path):
        path = file_path.decode("utf-8")
        dirpath = pathlib.PureWindowsPath(str(path)).as_posix()
        self.ids.input_textbox.text = dirpath
        del self.folderpath[:]
        self.folderpath.append(self.ids.input_textbox.text)

    def template_download(self, name):
        if self.folderpath != []:
            os.chdir(self.folderpath[0])
            filepath = pathlib.Path(self.folderpath[0]).joinpath(name)
            if not filepath.exists():
                self.create_template(filepath)
            else:
                self.alert = MDDialog(title='File Already Exists',
                                      text='Do you want to replace existing file?', buttons=[
                                          MDRaisedButton(text="YES",
                                                         on_release=self.create_template),
                                          MDRaisedButton(text="NO",
                                                         on_release=self.self_dismiss
                                                         )])
                self.alert.open()
        else:
            alert2 = MDDialog(title='No Input Found',
                              text='Please provide the path for data modules')
            alert2.open()

    def create_template(self, button=None):
        Snackbar(text='    Downloading Template...',
                 snackbar_x=dp(20), snackbar_y=dp(20),
                 size_hint_x=(Window.width - (dp(20) * 2)) / Window.width).open()
        os.chdir(pathlib.PureWindowsPath(os.path.join(
            app_path, r'source/create_template')).as_posix())
        template_download = pathlib.PureWindowsPath(os.path.join(
            app_path, r'source/create_template/launch_creator.bat')).as_posix()
        subprocess.call(
            [template_download, self.ids.input_textbox.text])
        try:
            self.alert.dismiss()
        except AttributeError:
            pass

    def self_dismiss(self, button=None):
        self.alert.dismiss()

    def change_attribs(self):
        if pathlib.Path(self.ids.input_textbox.text).joinpath('dm_label.xlsx').exists():
            self.pop_up1()
            executor = concurrent.futures.ThreadPoolExecutor()
            if self.ids.infoname_checkbox.state == 'down':
                executor.submit(self.mod_process, False, False)
            elif self.ids.techname_checkbox.state == 'down':
                executor.submit(self.mod_process, False, True)
            else:
                executor.submit(self.mod_process, False, None)
        else:
            MDDialog(title="Template Not Exist",
                     text="Please check for the excel file named \"dm_label.xlsx\" in the folder").open()

    def update_attribs(self):
        self.pop_up1()
        executor = concurrent.futures.ThreadPoolExecutor()
        if self.ids.infoname_checkbox.state == 'down':
            executor.submit(self.mod_process, True, False)
        elif self.ids.techname_checkbox.state == 'down':
            executor.submit(self.mod_process, True, True)
        else:
            executor.submit(self.mod_process, True, None)

    def pop_up1(self):
        self.dialog = MDDialog(
            size_hint=(None, None),
            width=dp(200),
            auto_dismiss=True,
            type="custom",
            content_cls=Matter(),
        )
        self.dialog.open()

    @mainthread
    def mod_process(self, choice, option):
        if self.folderpath != []:
            try:
                os.chdir(pathlib.PureWindowsPath(os.path.join(
                    app_path, r'source/infotech')).as_posix())
                infotech_edit = pathlib.PureWindowsPath(os.path.join(
                    app_path, r'source/infotech/launch_infotech.bat')).as_posix()
                if (choice == True) and (option == True):
                    subprocess.call(
                        [infotech_edit, self.ids.input_textbox.text, "1", "1"])
                elif (choice == True) and (option == False):
                    subprocess.call(
                        [infotech_edit, self.ids.input_textbox.text, "1", "0"])
                elif (choice == False) and (option == True):
                    subprocess.call(
                        [infotech_edit, self.ids.input_textbox.text, "0", "1"])
                elif (choice == False) and (option == False):
                    subprocess.call(
                        [infotech_edit, self.ids.input_textbox.text, "0", "0"])
                elif (choice == True) and (option == None):
                    subprocess.call(
                        [infotech_edit, self.ids.input_textbox.text, "1", "2"])
                elif (choice == False) and (option == None):
                    subprocess.call(
                        [infotech_edit, self.ids.input_textbox.text, "0", "2"])
                time.sleep(2)
                self.dialog.dismiss()
            except NotADirectoryError:
                self.dialog.dismiss()
                incorrect_ip = MDDialog(title='Not A Directory',
                                        text='Please provide the path for folder, not a file')
                incorrect_ip.open()
        else:
            alert = MDDialog(title='No Input Found',
                             text='Please provide the path for data modules')
            self.dialog.dismiss()
            alert.open()

    def clear_input(self):
        del self.folderpath[:]
        self.ids.input_textbox.text = 'Please drop your folder here'


class Tooltip(Label):
    pass


class HoverFlat(MDRectangleFlatButton, HoverBehavior):

    def __init__(self, **kwargs):
        super(HoverFlat, self).__init__(**kwargs)

    def on_enter(obj):
        obj.text_color = [66/255, 175/255, 1, 1]
        obj.md_bg_color = [50/255, 100/255, 150/255, 1]

    def on_leave(obj):
        obj.text_color = [66/255, 175/255, 1, 1]
        obj.md_bg_color = [0.15, 0.15, 0.15, 1]

    def on_click(obj, parent):
        for each in parent.children:
            if each != obj:
                each.md_bg_color = [0.07, 0.07, 0.07, 1]
                each.elevation = 0
            else:
                obj.elevation = 10
                obj.md_bg_color = [0.4, 0.4, 0.4, 1]


class HoverCard(MDRaisedButton, HoverBehavior):

    def __init__(self, **kwargs):
        # Window.bind(mouse_pos=self.on_mouse_pos)
        super(HoverCard, self).__init__(**kwargs)

    def on_enter(obj):
        obj.text_color = [66/255, 175/255, 1, 1]
        obj.md_bg_color = [50/255, 100/255, 150/255, 1]

    def on_leave(obj):
        obj.text_color = [0, 0, 0, 1]
        obj.md_bg_color = [25/255, 154/255, 229/255, 1]

    def on_click(obj, parent):
        for each in parent.children:
            if each != obj:
                each.md_bg_color = [0.07, 0.07, 0.07, 1]
                each.elevation = 0
            else:
                obj.elevation = 10
                obj.md_bg_color = [0.4, 0.4, 0.4, 1]


class HoverFloat(MDFloatingActionButton, HoverBehavior):

    def __init__(self, **kwargs):
        # Window.bind(mouse_pos=self.on_mouse_pos)
        super(HoverFloat, self).__init__(**kwargs)

    def on_enter(obj):
        obj.text_color = [66/255, 175/255, 1, 1]
        obj.md_bg_color = [50/255, 100/255, 150/255, 1]

    def on_leave(obj):
        obj.text_color = [0, 0, 0, 1]
        obj.md_bg_color = [25/255, 154/255, 229/255, 1]


class Matter(MDCard):
    pass


class Label_Composer(MDApp):

    def build(self):
        self.theme_cls.theme_style = 'Dark'
        self.theme_cls.material_style = "M3"
        self.theme_cls.primary_palette = 'LightBlue'
        return RootPopup().open()

    def action_quit(self, function: str):
        wmi = win32com.client.GetObject('winmgmts:')

        def getpid(process_name):
            return [item.split()[1] for item in os.popen('tasklist').read().splitlines()[4:] if process_name in item.split()]

        for p in wmi.InstancesOf('win32_process'):
            if p.Name == function:
                process_id = getpid(function)
                if len(process_id) > 1:
                    for x in process_id:
                        pid = int(x)
                        os.kill(pid, signal.SIGTERM)
                else:
                    pid = int("".join(getpid(function)))
                    os.kill(pid, signal.SIGTERM)


def reset():
    import kivy.core.window as window
    from kivy.base import EventLoop
    if not EventLoop.event_listeners:
        from kivy.cache import Cache
        window.Window = window.core_select_lib(
            'window', window.window_impl, True)
        Cache.print_usage()
        for cat in Cache._categories:
            Cache._objects[cat] = {}


if __name__ == '__main__':
    reset()
    Label_Composer().run()
