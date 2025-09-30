import win32com.client as win32
from os import sys
from utils.questionary import interactive
from utils.utils import banner
from utils.argparse import parse_args



def main():
    banner()
    
    if len(sys.argv) == 1:
        try:
            interactive()
        except KeyboardInterrupt:
            print("\nOperation cancelled, CRTL + C was pressed !!")
            sys.exit(0)
    else:
       parse_args()
        
    

        

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        
        print("\n[+] Exiting...")
        sys.exit(0)


