#coding=UTF-8

import DaVinciResolveScript as dvr_script
import time
import re
import os
import threading
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askdirectory

resolve = dvr_script.scriptapp("Resolve")
projectManager = resolve.GetProjectManager()
project = projectManager.GetCurrentProject()

def find_marker(idx):
    """load timeline by idx, and find the marker at the beginning"""
    timeline_by_idx = project.GetTimelineByIndex(idx) 
    project.SetCurrentTimeline(timeline_by_idx) 
    timeline = project.GetCurrentTimeline()
    timeline_name = timeline.GetName()
    markers = timeline.GetMarkers()
    if 0 in markers:
        # because the marker use to store info is at the beginning of the timeline, the frame is 0
        info = markers[0]
        info["timeline_name"] = timeline_name
        return info
    elif 0 not in markers:
        # if the timelien have no marker, need to add a new marker
        return idx

def load_timeline(info):
    """load maker info and save it in the dict named episode_info"""
    episode_info = {}
    if info["color"] == "Blue":
        episode_info["timeline_name"] = info["timeline_name"]
        episode_info["episode"] = info["name"]
        # marker name store the info of episode
        episode_info["version"] = int(info["note"])
        # marker note store the info of version
        episode_info["color"] = info["color"]
        # maker color "Blue" means it is a useful timeline and need to be render
    elif info["color"] == "Red":
        episode_info["timeline_name"] = info["timeline_name"]
        episode_info["episode"] = None
        episode_info["version"] = None
        episode_info["color"] = "Red"
        # maker color "Red" means this timeline don't need to be render
    return episode_info

def define_timeline(timeline_version):
    """determining episodes base on the name of timeline"""
    episode_info = {}
    timeline = project.GetCurrentTimeline() 
    name = timeline.GetName() 
    if name[0].lower() == "e":
        numbers = [float(s) for s in re.findall(r"-?\d+\.?\d*", name)]
        episode = int(numbers[0])
        confirm = [name, episode]
        # Ask for user confirmation
        if user_confirm(confirm):
            version = add_version(timeline_version,episode) + 1
            episode_info["timeline_name"] = name
            episode_info["episode"] = str(episode)
            episode_info["version"] = version
            episode_info["color"] = "Blue"
            add_marker("Blue", str(episode), str(version))
            return episode_info
    else:
        episode_info["timeline_name"] = name
        episode_info["episode"] = None
        episode_info["version"] = None
        episode_info["color"] = "Red"
        add_marker("Red", " ", " ")
        return episode_info

def add_version(timeline_version,episode):
    """Load all information for timelines with markers, and derive the version with newly added timeline."""
    version = 0
    for idx in timeline_version:
        info = timeline_version[idx]
        if info["episode"] == str(episode) and info["version"] >= version:
            version = info["version"]
    return version

def user_confirm(confirm):
    """Ask for user confirmation"""
    root = Tk() 
    root.geometry("400x100+200+300")
    confirm_text = "Timeline name: " + confirm[0] + "\n"
    episode_confirm = IntVar()
    episode_confirm.set(confirm[1])
    label01 = Label(root, text = confirm_text)
    label01.grid(row=0, column=0)
    entry01 = Entry(root, textvariable=episode_confirm)
    entry01.grid(row=1, column=0)
    btn01 = Button(root, text="confirm", command = root.destroy)
    btn01.grid(row=1, column=1)
    root.mainloop()
    return episode_confirm.get()

def add_marker(color,episode,version):
    """Add markers to identified timelines for easier retrieval next time"""
    timeline = project.GetCurrentTimeline()
    timeline.AddMarker(0, color, episode, version, 1, "")

def load_all_timeline():
    """main program to load all timelines"""
    timeline_version = {}
    unmark_timeline = []
    tc = project.GetTimelineCount() 
    for idx in range(1, tc + 1):
        info = find_marker(idx)
        if type(info) == dict:
            episode_info = load_timeline(info)
            timeline_version[idx] = episode_info
        if type(info) == int:
            unmark_timeline.append(idx)
    for idx in unmark_timeline:
        timeline_by_idx = project.GetTimelineByIndex(idx) 
        project.SetCurrentTimeline(timeline_by_idx) 
        episode_info = define_timeline(timeline_version)
        timeline_version[idx] = episode_info
    return timeline_version

class Application(Frame):
    """GUI main program"""
    def __init__(self, timeline_version, master=None):
        super().__init__(master) 
        self.master = master 
        self.timeline_version = timeline_version
        self.episode_number = self.episode_number(timeline_version)
        self.time_str = self.get_time()
        self.pack()
        self.createWidget(timeline_version) 
        self.createWidget2()

    def createWidget(self, timeline_version):
        """创建组件"""
        episode_number = self.episode_number
        episode_list = []
        self.version_vars = []
        self.render_vars = []
        self.name_vars = []
        # create the list of episode string
        for ep in range(1, episode_number + 1):
            format_episode = "EP" + "{:0>2d}".format(ep)
            episode_list.append(format_episode)
            # create the list of version
            version_number = self.version_number(timeline_version, ep)
            version_options = self.version_options(version_number)
            # create the optionmenu of version
            version_variable = StringVar()
            version_variable.set(version_options[-1])
            self.version_vars.append(version_variable)
            OptionMenu(self, self.version_vars[ep-1], *version_options)\
            .grid(row = ep - 1, column = 2)
            # create the check button for each episode
            render_variable = IntVar()
            render_variable.set(1)
            self.render_vars.append(render_variable)
            Checkbutton(self, variable=self.render_vars[ep-1], onvalue=1, offvalue=0)\
            .grid(row = ep - 1, column = 0, sticky = "w")
        # create the button to load specific episode.
        for inx, episode_name in enumerate(episode_list):
            Button(self, text = episode_name, width=5, command = lambda arg=int(episode_name[-2:]):self.set_timeline(arg, self.version_vars[arg-1].get()))\
            .grid(row = inx, column = 1, sticky = "e")
            # create the Entry to display timeline name
            timelinename = StringVar()
            name = self.timeline_name(inx + 1, self.version_vars[inx].get())
            timelinename.set(name)
            self.name_vars.append(timelinename)
            Entry(self, textvariable = self.name_vars[inx], width = 8)\
            .grid(row = inx, column = 3, sticky = "e")
        # the entry to select render path
        self.path = StringVar()
        self.path.set("")
        entry01 = Entry(self, textvariable = self.path)
        entry01.grid(row = episode_number + 1, column = 2, columnspan = 2, sticky = "we")
        btn_01 = Button(self, text = "choose", command = self.select_path)
        btn_01.grid(row = episode_number + 1, column = 1, sticky = "we")
        # the render button
        btn_render = Button(self, text = "render", width=5, command = self.render_all)
        btn_render.grid(row = episode_number + 2, column = 1)
    
    def createWidget2(self):
        """another fuctuon to import xml as timeline"""
        self.importpath = StringVar()
        self.importpath.set("")
        entry_importpath = Entry(self, textvariable=self.importpath)
        entry_importpath.grid(row = 0, column=5)
        btn_pathselect = Button(self, text = "选择", command = self.select_importpath)
        btn_pathselect.grid(row = 0, column=6)
        btn_import = Button(self, text = "开始", command = self.import_beginning)
        btn_import.grid(row = 1, column=5)
            
    def episode_number(self, timeline_version):
        """How many episodes are there in total"""
        episodes = []
        for idx in timeline_version:
            info = timeline_version[idx]
            if info["episode"] != None:
                episode = int(info["episode"])
                episodes.append(episode)
        episode_number = max(episodes)
        return episode_number 

    def version_number(self, timeline_version, episode):
        """How many versions are there for a particular episode"""
        versions = [1]
        for idx in timeline_version:
            info = timeline_version[idx]
            if info["episode"] == str(episode):
                versions.append(info["version"])
        version_number = max(versions)
        return version_number

    def version_options(self, version_number):
        """create a list for optionmenu"""
        version_options = []
        for v in range(1, version_number + 1):
            version_names = "version" + str(v)
            version_options.append(version_names)
        return version_options
    
    def set_timeline(self, episode, version_variable):
        """Set timeline based on episode and version."""
        timeline_version = self.timeline_version
        version = int(version_variable[7:])
        action = False
        for idx in timeline_version:
            info = timeline_version[idx]
            if info["episode"] == str(episode) and info["version"] == version:
                print("ok")
                timeline_by_idx = project.GetTimelineByIndex(idx) 
                project.SetCurrentTimeline(timeline_by_idx)
                timeline = project.GetCurrentTimeline() 
                name = timeline.GetName() 
                self.name_vars[episode - 1].set(name)
                action = True
            else:
                pass
        if action == False:
            print("Wrong,can't find the timeline")

    def timeline_name(self, episode, version_variable):
        """find timeline name based on episode and version."""
        timeline_version = self.timeline_version
        version = int(version_variable[7:])
        for idx in timeline_version:
            info = timeline_version[idx]
            if info["episode"] == str(episode) and info["version"] == version:
                print("find name")
                return info["timeline_name"]
            else:
                pass
    

    def render_all(self):
        """Render all selected episodes."""
        episode_number = self.episode_number
        for ep in range(1, episode_number + 1):
            if self.render_vars[ep - 1].get() == 1:
                self.set_timeline(ep, self.version_vars[ep - 1].get())
                self.render_timeline(ep)
            elif self.render_vars[ep - 1].get() == 0:
                print(str(ep) + ": 不渲染")
            else:
                print(self.render_vars[ep - 1])
    
    def get_time(self):
        """Automatically get the time string based on the current date."""
        time_str = time.strftime("%m%d", time.localtime())
        return time_str
    
    def render_timeline(self, ep):
        """Render current episode"""
        print("开始添加" + str(ep))
        time_str = self.time_str
        name = "{:0>2d}".format(ep) + "_" + time_str
        path = self.path.get()
        mode = project.GetCurrentRenderMode() 
        if mode == 1:
            project.SetRenderSettings({"CustomName": name})
            project.SetRenderSettings({"TargetDir": path})
            project.AddRenderJob()
        elif mode == 0:
            project.SetRenderSettings({"TargetDir": path + "/" + name})
            project.AddRenderJob()
    
    def select_path(self):
        """select the path for render"""
        path = askdirectory()
        if path == "":
            self.path.get()
        else:
            self.path.set(path)
    
    def select_importpath(self):
        """select the path for import xml"""
        path = askdirectory()
        if path == "":
            self.importpath.get()
        else:
            self.importpath.set(path)
    
    def import_beginning(self):
        """import all xml and file as a timeline"""
        mediaPool = project.GetMediaPool() 
        folder = mediaPool.GetCurrentFolder()
        path_list = []
        for item in os.listdir(self.importpath.get()):
            if item[0]==".":
                    pass
            elif item[0]!=".":
                path_list.append(item)
        for filepath in path_list:
            path = os.path.join(self.importpath.get(), filepath)
            list = os.listdir(path)
            numbers = [float(s) for s in re.findall(r"-?\d+\.?\d*", filepath)]
            episode = int(numbers[0])
            format_episode = "EP" + "{:0>2d}".format(episode)
            for file in list:
                if file.endswith(".xml") or file.endswith(".fcpxml"):
                    xmlpath = os.path.join(path, file)

            import_folder = mediaPool.AddSubFolder(folder, format_episode + "_" +self.time_str)
            mediaPool.SetCurrentFolder(import_folder) 
            t = threading.Thread(target=self.import_timeline(xmlpath, path))
            t.start()
            t.join()
            mediaPool.SetCurrentFolder(folder) 
    
    def import_timeline(self, xmlpath, path):
        """import timeline by xml"""
        mediaPool = project.GetMediaPool() 
        items = mediaPool.ImportMedia(path)
        for item in items:
            item.SetClipColor("Orange") 
        folder = mediaPool.GetCurrentFolder()
        list = [folder]
        mediaPool.ImportTimelineFromFile(xmlpath,{
            "importSourceClips":False,
            "sourceClipsFolders":list,
        })

if __name__ == "__main__": 
    timeline_version = load_all_timeline()
    root = Tk() 
    root.geometry("400x500+200+300")
    app = Application(timeline_version,root) 
    root.mainloop()



