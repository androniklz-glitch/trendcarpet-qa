Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\AndronikLindgren\.gemini\antigravity\scratch\Projects\Trendcarpet_Interiors"
WshShell.Run """C:\Users\AndronikLindgren\miniconda3\python.exe"" start_hub_with_tunnel.py", 0, False

