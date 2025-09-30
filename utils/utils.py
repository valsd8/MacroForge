import logging
import base64
from os import sys
from rich.console import Console
from rich.panel import Panel
from pyfiglet import Figlet

def powershell_base64(command: str, encoding: str = 'utf-16le') -> str:
        
        b = command.encode(encoding)         
        return base64.b64encode(b).decode('ascii')

from rich.console import Console
from rich.panel import Panel
from rich.align import Align
from rich.text import Text
from pyfiglet import Figlet

def banner():
    console = Console()
    f = Figlet(font='slant')
    
    ascii_banner = f.renderText('MacroForge')
    ascii_banner = Align.center(ascii_banner)

    subtitle = Text("by val.sd8", style="bold cyan")
    version_tag = Text("MacroForge v0.1", style="bold magenta")

    panel = Panel.fit(
        ascii_banner,
        title=version_tag,
        subtitle=subtitle,
        border_style="bright_magenta",
        padding=(1, 4),
    )

    console.print(panel)

# logging_setup.py


