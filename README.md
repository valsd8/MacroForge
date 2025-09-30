# ⚔️ MacroForge

> Generate simple Office macros . Created by `val.sd8`.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Status](https://img.shields.io/badge/status-dev-red)

---

## 🔍 Description

**MacroForge** is an Office macro generator designed to generate simple reverse shell and powershell commands in Word / Excel macros.

- CLI and questionnary usage

-More feature comming soon !!

> ⚠️ For **educational and authorized** use only.

---

## 🛠️ Installation (You may need to use a virtual python env)

```bash
git clone https://github.com/valsd8/MacroForge
cd macroforge
pip install -r requirements.txt
```

## Usage

For now this only run on windows where office is installed. Dont forget to check this case in Word and Excel or else it will fails. 

## Help Command

```bash
python3 main.py --help
╭───────────────────────── MacroForge v0.1 ─────────────────────────╮
│                                                                   │
│        __  ___                      ______                        │
│       /  |/  /___ _______________  / ____/___  _________ ____     │
│      / /|_/ / __ `/ ___/ ___/ __ \/ /_  / __ \/ ___/ __ `/ _ \    │
│     / /  / / /_/ / /__/ /  / /_/ / __/ / /_/ / /  / /_/ /  __/    │
│    /_/  /_/\__,_/\___/_/   \____/_/    \____/_/   \__, /\___/     │
│                                                  /____/           │
│                                                                   │
│                                                                   │
╰─────────────────────────── by val.sd8 ────────────────────────────╯
usage: MacroForge [-h] [--file {Word,Excel}] [--mode {Revshell,Cmd}] [--ip IP] [--port PORT]
                  [--payload PAYLOAD] [--yes]

Generate reverse shell and custom commands for Excel / Word macros. Use interactive mode if no args
supplied.

options:
  -h, --help            show this help message and exit
  --file {Word,Excel}   Target file type
  --mode {Revshell,Cmd}
                        Operation mode
  --ip IP               Listener IP (for revshell)
  --port PORT           Listener port (for revshell)
  --payload PAYLOAD     Payload string / command (for 'cmd' modes)
  --yes, -y             Skip prompts
  
```


![Enabling VBA object model](content/objectModelVBA.png)


**Getting a reverse shell using the cli**
![Simple example on how to get a reverse shell using the tool with the cli](content/reverse_shell_demo.png)

**Getting a reverse shell using the interactive prompt**
![same with the interactive mode](content/reverse_shell_using_interactive.png)
